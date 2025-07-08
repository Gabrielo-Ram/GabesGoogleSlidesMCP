/**
 * 'global.ts' is a dedicated module that contains global variables that are used across
 * this application such as:
 *  - The current presentation ID
 *  - The Authorization token
 */

import { slides_v1 } from "googleapis";

export let presentationID: string | null = null;

export function setPresentationID(newID: string) {
  presentationID = newID;
}

export let auth: slides_v1.Slides | null = null;

export function setAuthToken(token: slides_v1.Slides) {
  auth = token;
}
