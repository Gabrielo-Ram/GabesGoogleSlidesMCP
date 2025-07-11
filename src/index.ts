/**
 * `src/index.ts` is the entry file for the MCP server. This file registers the
 * tools and resources that an MCP client (LLM) will be able to reference.
 * Logic that uses the Google Slides API is not found in this file.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import process from "process";
import path from "path";
import fs from "fs";
import { parse } from "csv-parse";
import dotenv from "dotenv";
import { fileURLToPath } from "url";
import { initiateSlides } from "./createPresentation.js";
import { addCustomSlide } from "./addSlides.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.resolve(__filename);

dotenv.config({ path: path.resolve(__dirname, "../env") });

// Create server instance
const server = new McpServer({
  name: "Gabe's Google Slides MCP",
  version: "1.0.0",
  capabilities: {
    resources: {},
    tools: {},
  },
});

/**
 * A server tool that validates raw JSON input from the LLM. This tool is
 * subject to change depending on where we are reading in data from
 * (eg. Notion database, CSV files, Excel spreadsheets, Google Sheets)
 * Maybe make a validation tool for each file type? An LLM is able to
 * restructure data to fit the tool it is calling.
 *
 * server.tool("validate-data", "A server tool that validates...",
 * {
 *  data: z.array.object()
 * }
 * )
 */

const extractDataToolDescription = `This tool extracts data from one row in a provided CSV file. This tool is totally optional, only call this tool if the user provides a CSV file to be read. 

This tool accepts two parameters:
  - name: The name/title of the row of data the user wants to extract from the CSV string.
  - csvFile: The CSV string itself. 

The user will provide a CSV file, and you will recieve a string value representing that CSV file. You must pass in that CSV string as input into this tool.

IMPORTANT: The user MUST specify the what row of data they want to extract from the CSV. If unclear, please prompt the user to fulfill the missing information. The user can only create one presentation at a time.  

The output of this tool will be one JSON object that represents data for one 
row. Each key will represent one property in that CSV row. Each value will 
be the data associated with that property. Here is a sample JSON object this tool
could output: 

{
  companyName: string;
  location: string;
  foundedYear: number;
  arr: number;
  industry: string;
  burnRate: number;
  exitStrategy: string;
  dealStatus: string;
  fundingStage: string;
  investmentAmount: number;
  investmentDate: string;
  keyMetrics: string;
  presentationId?: string;
};

You must use the information in this object as context to create your custom slides in a Google Slides presentation. Carefully retain and reference the returned data throughout the slide creation process.`;

server.tool(
  "extract-data",
  extractDataToolDescription,
  {
    name: z
      .string()
      .describe(
        "The name (the first property) of the row who's data we want to extract."
      ),
    csvFile: z.string().describe("The CSV string we are extracting data from"),
  },
  async ({ name, csvFile }) => {
    try {
      const lines = csvFile.trim().split("\n");

      const headers = lines[0].split(",").map((header) => header.trim());
      const rows = lines
        .slice(1)
        .map((line) => line.split(",").map((cell) => cell.trim()));

      const selectedRow = rows.find(
        (row) => row[0].trim().toLowerCase() === name.trim().toLowerCase()
      );
      if (!selectedRow) {
        return {
          content: [
            {
              type: "text",
              text: `Could not find a title with the name ${name}. Is it spelled correctly?`,
            },
          ],
        };
      }

      const rowData: Record<string, string> = {};

      headers.forEach((header, index) => {
        rowData[header] = selectedRow[index] ?? "";
      });

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(rowData),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `An error occurred while extracting company data:\n${
              (error as Error).message
            }`,
          },
        ],
      };
    }
  }
);

const createPresentationToolDescription = `A tool used to create and style a Google Slides presentation into the user's account. 

This tool creates ONE presentation. The user can only create one presentation at a time. 

This tool's only paramater is a the title of the presentation you want to create (usually the name of the company).

This tool will return a string representing the presentationId for the created presentation.
This value will be important if the user decides to add a custom slide using
the addCustomSlide() server tool. Remember this data point.`;

//Registers a tool that creates a Google Slides presentation based on data recieved.
server.tool(
  "create-presentation",
  createPresentationToolDescription,
  {
    companyName: z
      .string()
      .describe(
        "The name of the company, and therefore the title of the newly created slide"
      ),
  },
  async ({ companyName }) => {
    //Validates input
    if (!companyName) {
      throw new Error("Missing or invalid input data for create-presentation");
    }

    try {
      const presentationId = await initiateSlides(companyName);

      return {
        content: [
          {
            type: "text",
            text: `Presentation ID: ${presentationId}`,
          } as { [x: string]: unknown; type: "text"; text: string },
        ],
      };
    } catch (error) {
      console.error(
        "create-presentation: Error creating presentation:\n",
        error
      );

      return {
        content: [
          {
            type: "text",
            text: `Failed to create presentation\n${error}`,
          } as { [x: string]: unknown; type: "text"; text: string },
        ],
      };
    }
  }
);

const addCustomSlideToolDescription = `This tool allows you to create a custom slide based on any content/context you create/have.

The user can create a maximum of 10 custom slides at at time. The content and styling of custom slide is entirely up to your discretion.

This tool's parameters are the following: 
- A custom title for this new slide you create. Make up your own title.
- Text content you compose based on the user's request. You must use the data extracted from extract-company-data as context for the writing you produce. If you decide the content of this slide is writting in bullet-form, create a new line to seperate each "bullet point" 
- The presentationID for the most recent, or most appropriate, presentation you've created. This is a string value.
- A string value that one the following: "Paragraph" or "Bullet". The user may or may not specify what slide type they want to create. If they do not, adjust this input so that it will best fit the content that will go inside this slide.

You DO NOT NEED to create a new presentation when you run this tool. Simply retrieve the presentationID that was returned when you ran the create-presentation tool.`;

server.tool(
  "add-custom-slide",
  addCustomSlideToolDescription,
  {
    slideTitle: z.string().describe("Your title for this custom-made slide"),
    slideContent: z
      .string()
      .describe("The content made specifically for this slide by the LLM."),
    presentationId: z
      .string()
      .describe(
        "The presentationID for the most recently created Google Slides presentation."
      ),
    slideType: z
      .string()
      .describe(
        "One of the literal strings 'Paragraph' or 'Bullet' that specifies what kind of slide this tool will create"
      ),
  },
  async ({ slideTitle, slideContent, presentationId, slideType }) => {
    try {
      //Validates 'slideType' input
      if (slideType !== "Paragraph" && slideType !== "Bullet") {
        return {
          content: [
            {
              type: "text",
              text: "add-custom-slide: Invalid slideType parameter. You must pass in either 'Paragraph' or 'Bullet' as literal strings.",
            },
          ],
        };
      }

      //Adds a new slide to the presentation.
      await addCustomSlide(slideTitle, slideContent, presentationId, slideType);

      return {
        content: [
          {
            type: "text",
            text: `Succesfully created a new custom slide!`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `There was an error creating a custom slide:\n${error}`,
          },
        ],
      };
    }
  }
);

//Main function used to connect to an MCP Client.
export async function engageServer() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error(
    "\nGoogle Slides MCP Server:\nOpened a connection at new Stdio Transport..."
  );
}

engageServer().catch((error) => {
  console.error("Fatal error in main(): ", error);
  process.exit(1);
});
