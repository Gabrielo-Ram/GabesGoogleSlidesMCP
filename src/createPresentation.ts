/**
 * 'createPresentation.ts' is a central hub for logic that creates a new Google
 * Slides presentation into the user's Google Drive.
 *
 * This file handles authorization and authentication of Google details and
 * it handles creating a new presentation and adding a 'Title' slide.
 *
 */

import {
  presentationID,
  auth,
  setPresentationID,
  setAuthToken,
} from "./globals.js";
import { addCustomSlide } from "./addSlides.js";
import { CompanyData } from "./types.js";
import { google, slides_v1 } from "googleapis";
import { OAuth2Client } from "google-auth-library";
import { authenticate } from "@google-cloud/local-auth";
import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const SCOPES = ["https://www.googleapis.com/auth/presentations"];
const CREDENTIALS_PATH = path.join(__dirname, "../credentials.json");
const TOKEN_PATH = path.join(__dirname, "../token.json");

/* ----———————————————————————————————————————————————--- */
/**
 * Code taken from Google's Node.js quickstart tutorial
 * on 'developers.google.com'. Creates and configures necessary tokens
 * and data to use the Google Slides API.
 */

/**
 * Reads previously authorized credentials from the save file.
 *
 * @return {Promise<OAuth2Client|null>}
 */
async function loadSavedCredentialsIfExist() {
  try {
    const content = await fs.readFile(TOKEN_PATH, "utf-8");
    const credentials = JSON.parse(content);
    return google.auth.fromJSON(credentials) as any;
  } catch (err) {
    return null;
  }
}

/**
 * Serializes credentials to a file compatible with GoogleAuth.fromJSON.
 *
 * @param {OAuth2Client} client
 * @return {Promise<void>}
 */
async function saveCredentials(client: OAuth2Client) {
  const content = await fs.readFile(CREDENTIALS_PATH, "utf-8");
  const keys = JSON.parse(content);
  const key = keys.installed || keys.web;
  const payload = JSON.stringify({
    type: "authorized_user",
    client_id: key.client_id,
    client_secret: key.client_secret,
    refresh_token: client.credentials.refresh_token,
  });
  await fs.writeFile(TOKEN_PATH, payload);
}

/**
 * Load client credentials to access the Google API.
 */
export async function authorize() {
  let client = await loadSavedCredentialsIfExist();
  if (client) {
    return client;
  }
  client = await authenticate({
    scopes: SCOPES,
    keyfilePath: CREDENTIALS_PATH,
  });
  if (client.credentials) {
    await saveCredentials(client);
  }
  return client;
}

/* ----———————————————————————————————————————————————--- */
// Logic for creating and a Google Slides Presentations

/**
 * Creates a Google Slide presentation and edits the 'Title' slide to
 * contain the company name and date the presentation was created.
 * @param {string} companyName The name of the company.
 * @param {slides_v1.Slides} service The Google Slides API instance.
 * @returns A presentationId string.
 */
async function createPresentation(
  companyName: string,
  service: slides_v1.Slides
) {
  //Creates the presentation and adds the Company name and current date into the title
  try {
    const presentation = await service.presentations.create({
      requestBody: {
        title: `${companyName} Slide Deck`,
      },
    });

    //Gets the presentationId and sets it to the global variable
    const presentationId = presentation.data.presentationId;
    if (!presentationId) {
      throw new Error("presentationId is null or undefined");
    }
    setPresentationID(presentationId);

    //Gets a reference to the 'Title Slide'
    const slideData = await service.presentations.get({
      presentationId,
    });
    const firstSlide = slideData.data.slides?.[0];

    if (!firstSlide) {
      throw new Error("No slides found");
    }
    if (!firstSlide.pageElements || firstSlide.pageElements.length === 0) {
      throw new Error("No elements found on First Slide");
    }

    //Get the objectIDs of the two automatically generated textboxes.
    const titleId = firstSlide.pageElements[0].objectId;
    const subtitleId = firstSlide.pageElements[1].objectId;
    if (!subtitleId || !titleId) {
      throw new Error("titleID or subtitleId not set");
    }

    //Replace the text inside each textbox with the company name and subheader respectively
    await service.presentations.batchUpdate({
      presentationId,
      requestBody: {
        requests: [
          {
            insertText: {
              objectId: titleId,
              text: companyName,
              insertionIndex: 0,
            },
          },
          {
            insertText: {
              objectId: subtitleId,
              text: `Created: ${new Date().toISOString().split("T")[0]}`,
              insertionIndex: 0,
            },
          },
        ],
      },
    });

    return presentationId;
  } catch (err) {
    throw new Error("Failed to create presentation in createPresentation()");
  }
}

/**
 * Initiates the creation of an entire Pitch Deck.
 * Includes a 'Title' slide and any supplemental slides thereafter.
 *
 * @param {any} companyData The object containing raw data for the company.
 * @param {string} companyName The name of the company you want to create a presentation for.
 */
export async function initiateSlides(
  companyName: string,
  companyData?: object
) {
  try {
    //Creates and sets authorization token
    const auth = await authorize();
    const service = google.slides({ version: "v1", auth });
    setAuthToken(service);

    //Creates a new presentation and updates 'Title' slide
    const newPresentationId = await createPresentation(companyName, service);
    // companyData.presentationId = newPresentationId;
    setPresentationID(newPresentationId);
    if (!newPresentationId || !presentationID) {
      throw new Error("Failed to create presentation in initiateSlides()");
    }

    /*
    —————————————————————————————————————————————————————————————————————————————
      Add some pre-set slides below this line. Comment out this code-block if
      you don't want to add pre set slides. 
    */

    //Adds a new Bullet Slide
    const preSetBulletTitle = "[Insert pre-set title]";
    const preSetBulletContent =
      "[Insert pre-set content]\nThis is sample bullet content";
    await addCustomSlide(
      preSetBulletTitle,
      preSetBulletContent,
      presentationID,
      "Bullet"
    );

    //Formats data and creates a 'Paragraph' slide
    const preSetParagraphTitle = "[Insert pre-set title]";
    const preSetParagraphContent =
      "[Insert pre-set content]\nThis is sample paragraph content";
    await addCustomSlide(
      preSetBulletTitle,
      preSetBulletContent,
      presentationID,
      "Bullet"
    );

    console.error(`\nCreated presentation with ID: ${newPresentationId}`);
    return newPresentationId;
  } catch (error) {
    throw new Error(`There was an error in initiateSlides(): \n${error}`);
  }
}

//TESTING: Main function used for testing
async function main() {
  try {
    const testData = {
      companyName: "Hax @ Newark",
      location: "Newark NJ",
      foundedYear: 2011,
      arr: 20000000,
      industry: "Venture Capital",
      burnRate: 8000000,
      exitStrategy: "IPO",
      dealStatus: "Due Diligence",
      fundingStage: "Series C",
      investmentAmount: 200,
      investmentDate: "May 10, 2024",
      keyMetrics:
        "This is a really really cool company with really really cool people. Let's see how long I can make this test string to test how Google Slides handles text wrap. My name is Gabe Ramirez, an intern at Hax. I love music (I play guitar and dance), and honestly forgot how much I enjoyed coding and problem solving until this project.",
    };

    await initiateSlides("Hax @ Newark", testData);
  } catch (error) {
    console.error("\nThere was an error in the main function!\n", error);
  }
}

main();
