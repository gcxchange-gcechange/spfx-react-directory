/* eslint-disable @typescript-eslint/no-var-requires */
import * as strings from "SpfxReactDirectoryWebpartWebPartStrings";
// import * as english from "../loc/en-us.js";
// import * as french from "../loc/fr-fr.js";


const english = require('../loc/en-en.js');
const french = require('../loc/fr-fr.js');


// eslint-disable-next-line @typescript-eslint/explicit-function-return-type
export function SelectLanguage(lang:unknown) {
  switch (lang) {
    case "en-us": {
      return english;
    }
    case "fr-fr": {
      return french;
    }
    default: {
      return strings;
    }
  }
}