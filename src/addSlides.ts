import {
  presentationID,
  auth,
  setPresentationID,
  setAuthToken,
} from "./globals.js";
import { authorize } from "./createPresentation.js";
import { SlideType } from "./types.js";
import { google, slides_v1 } from "googleapis";

/**
 * Creates a new slide with a title and a paragraph who's content
 * is created entirely by an LLM.
 *
 * @param {string} title The Title of the slide
 * @param {string} content The content we are writing into the new slide
 */
async function addNewParagraphSlide(title: string, content: string) {
  if (!presentationID) {
    throw new Error(
      "addnewParagraphSlide(): 'presentationID' {string} does not exist! Have you created a presentation yet?"
    );
  }
  if (!auth) {
    throw new Error(
      "addNewParagraphSlide(): 'auth' {slides_v1.Slide} is not set!"
    );
  }

  //Adds a new slide in our presentation with a unique ID
  const uniqueID = `id_${Date.now()}_${Math.floor(Math.random() * 10000)}`;
  const newSlide = await auth.presentations.batchUpdate({
    presentationId: presentationID,
    requestBody: {
      requests: [
        {
          createSlide: {
            objectId: uniqueID,
            slideLayoutReference: {
              predefinedLayout: "TITLE_AND_BODY",
            },
          },
        },
      ],
    },
  });

  //Get a reference to the presentation and then the most recently created slide
  const presentation = await auth.presentations.get({
    presentationId: presentationID,
  });

  const slidesArray = presentation.data.slides;
  if (!slidesArray) {
    throw new Error(
      `addNewParagraphSlide(): Could not retrieve a reference to the slide array`
    );
  }

  const currentSlide = slidesArray[slidesArray.length - 1];
  if (
    !currentSlide ||
    !currentSlide.pageElements ||
    currentSlide.pageElements.length === 0
  ) {
    throw new Error(
      "addNewParagraphSlide(): Recently created slide not found or has no elements"
    );
  }

  //There are automatically two textbox elements in this layout.
  //Get the ID's of those SHAPES.
  const titleId = currentSlide.pageElements[0].objectId;
  const bodyId = currentSlide.pageElements[1].objectId;
  if (!bodyId || !titleId) {
    throw new Error("titleID or subtitleId not set");
  }

  //Add text to the slide and style them.
  const updateRes = await auth.presentations.batchUpdate({
    presentationId: presentationID,
    requestBody: {
      requests: [
        {
          insertText: {
            objectId: titleId,
            text: title,
            insertionIndex: 0,
          },
        },
        {
          updateTextStyle: {
            objectId: titleId,
            style: {
              bold: true,
              fontFamily: "Times New Roman",
              fontSize: {
                magnitude: 25,
                unit: "PT",
              },
            },
            textRange: { type: "ALL" },
            fields: "bold,fontFamily,fontSize",
          },
        },
        {
          insertText: {
            objectId: bodyId,
            text: content,
            insertionIndex: 0,
          },
        },
      ],
    },
  });
}

/**
 * Creates a new slide with a title and bullet points for key metrics.
 * The 'content' of this slide is organized into bullet points. The LLM knows to create
 * a new bullet point by adding a new line '\n' to the string it provides.
 *
 * @param {string} title The title of this new slide
 * @param {string} content The content of this new slide
 */
async function addNewBulletSlide(title: string, content: string) {
  if (!presentationID) {
    throw new Error(
      "addnewParagraphSlide(): 'presentationID' {string} does not exist! Have you created a presentation yet?"
    );
  }
  if (!auth) {
    throw new Error(
      "addNewParagraphSlide(): 'auth' {slides_v1.Slide} is not set!"
    );
  }

  //Adds a new slide in our presentation with a unique ID
  const uniqueID = `id_${Date.now()}_${Math.floor(Math.random() * 10000)}`;
  const newSlide = await auth.presentations.batchUpdate({
    presentationId: presentationID,
    requestBody: {
      requests: [
        {
          createSlide: {
            objectId: uniqueID,
            slideLayoutReference: {
              predefinedLayout: "TITLE_AND_BODY",
            },
          },
        },
      ],
    },
  });

  //Get a reference to the presentation and then the most recently created slide
  const presentation = await auth.presentations.get({
    presentationId: presentationID,
  });
  const slidesArray = presentation.data.slides;
  if (!slidesArray) {
    throw new Error(
      `addNewParagraphSlide(): Could not retrieve a reference to the slide array`
    );
  }

  const currentSlide = slidesArray[slidesArray.length - 1];
  if (
    !currentSlide ||
    !currentSlide.pageElements ||
    currentSlide.pageElements.length === 0
  ) {
    throw new Error(
      "addNewParagraphSlide(): Recently created slide not found or has no elements"
    );
  }

  //There are automatically two textbox elements in this layout.
  //Get the ID's of those SHAPES.
  const titleId = currentSlide.pageElements[0].objectId;
  const bodyId = currentSlide.pageElements[1].objectId;
  if (!bodyId || !titleId) {
    throw new Error("titleID or subtitleId not set");
  }

  //Add text to the slide and style them.
  const updateRes = await auth.presentations.batchUpdate({
    presentationId: presentationID,
    requestBody: {
      requests: [
        {
          insertText: {
            objectId: titleId,
            text: title,
            insertionIndex: 0,
          },
        },
        {
          updateTextStyle: {
            objectId: titleId,
            style: {
              bold: true,
              fontFamily: "Times New Roman",
              fontSize: {
                magnitude: 25,
                unit: "PT",
              },
            },
            textRange: { type: "ALL" },
            fields: "bold,fontFamily,fontSize",
          },
        },
        {
          insertText: {
            objectId: bodyId,
            text: content,
            insertionIndex: 0,
          },
        },
        {
          createParagraphBullets: {
            objectId: bodyId,
            textRange: {
              type: "ALL",
            },
            bulletPreset: "BULLET_DISC_CIRCLE_SQUARE",
          },
        },
      ],
    },
  });
}

/**
 * This function allows the user (or the LLM) to create a custom slide using
 * the functions defined above. The LLM will format a 'content' block and
 * place it into the new slide along with a custom title.
 *
 * @param {string} title The title of the custom slide.
 * @param {string} content The content placed inside the slide (formatted by the LLM)
 * @param {string} presentationId The ID of the presentation we want to add a slide to.
 * @param {SlideType} type The way the LLM identifies whether it wants to create a 'Paragraph' or a 'Bullet' slide
 */
export async function addCustomSlide(
  title: string,
  content: string,
  presentationId: string,
  type: SlideType
) {
  try {
    //Creates and sets authorization token
    const auth = await authorize();
    const service = google.slides({ version: "v1", auth });
    setAuthToken(service);

    //Sets the presentationID
    if (!presentationId) {
      throw new Error(
        "addCustomSlide(): The inputted 'presentationId' is not properly set."
      );
    }
    setPresentationID(presentationId);

    //Creates the appropriate type of slide depending on 'type'
    switch (type) {
      case "Paragraph":
        await addNewParagraphSlide(title, content);
        break;
      case "Bullet":
        await addNewBulletSlide(title, content);
        break;
      default:
        throw new Error(
          `addCustomSlide(): You must input a proper 'SlideType' string. ${type} is not valid`
        );
    }
  } catch (error) {
    throw new Error(
      `addCustomSlide(): Could not create a new slide: \n${error}`
    );
  }
}
