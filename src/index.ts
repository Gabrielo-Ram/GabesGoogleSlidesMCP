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

const extractCompanyDataToolDescription = `This tool extracts data from a specific company from a CSV file.
The user will provide a file or a file path to the CSV file to you. You must pass in that user input into this tool.
This tool MUST be called before create-presentation. It returns all the necessary data to generate a Google Slides presentation.
The user must specify the name of one company. If unclear, ask the user which company they want to create a presentation for. A user can only create a presentation for ONE company. 

The output of this tool will be one JSON object that represents data for one 
company. Each key will represent one property in that CSV row. Each value will 
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

This tool provides essential company data that you will use to generate content slides for a Google Slides presentation. Carefully retain and reference the returned data throughout the slide creation process.`;

server.tool(
  "extract-company-data",
  extractCompanyDataToolDescription,
  {
    companyName: z
      .string()
      .describe("The name of the company who's data we want to extract."),
    csvFile: z.string().describe("The file path to a CSV file"),
  },
  async ({ companyName, csvFile }) => {
    try {
      const pathToCSV = csvFile; //THIS PARAMTER TOTATLLY DEPENDS ON HOW YOU ARE DEPLOYING YOUR APPLICATION
      if (!pathToCSV) {
        throw new Error(
          "extract-company-data: csvFile not set properly. Please check how you are inputting the CSV file into this tool"
        );
      }

      const company = await new Promise<Record<string, any> | null>(
        (resolve, reject) => {
          const stream = fs
            .createReadStream(pathToCSV)
            .pipe(parse({ columns: true, trim: true }));

          stream.on("data", (row) => {
            if (
              row.companyName && //HOW WILL YOU KNOW THE PROPERTIES OF THE CSV FILE?
              row.companyName.trim().toLowerCase() ===
                companyName.trim().toLowerCase()
            ) {
              stream.destroy(); // stop the stream once a match is found resolve(row);
              resolve(row);
            }
          });

          stream.on("end", () => resolve(null)); // no match found
          stream.on("error", reject); // on error
        }
      );

      if (!company) {
        return {
          content: [
            {
              type: "text",
              text: `No company found with the name "${companyName}"`,
            },
          ],
        };
      }

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(company),
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
You must call 'extract-company-data' before running this tool.

This tool creates ONE presentation for ONE set of company data. 

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

const addCustomSlideToolDescription = `This tool allows the user to create a custom slide based on company data.
If you already don't have data for a specific company, you MUST call extract-company-data before you run this tool. You will use the output of extract-company-data to create content for a custom slide in this server tool. 

This tool's parameters are the following: 
- A custom title for this new slide you create. Make up your own title.
- A paragraph you compose based on the user's request. You must use the data extracted from extract-company-data as context for the writing you produce. 
- The presentationID for the most recent presentation you've created. This is a string value.
- A string value that is either: "Paragraph" or "Bullet". The user will specify what kind of slide they want to create (either a paragraph or bulleted slide). If they do not, ask the user to fill in this information before you call this tool.

You DO NOT NEED to create a new presentation when you run this tool. Simply retrieve the presentationID that was returned when you first ran the create-presentation tool.`;

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
        "The literal strings 'Paragraph' or 'Bullet' that specifies what kind of slide this tool will create"
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

// engageServer().catch((error) => {
//   console.error("Fatal error in main(): ", error);
//   process.exit(1);
// });
