import * as strings from "SpfxReactDirectoryWebpartWebPartStrings";

import * as english from "../loc/en-us.js";

import * as french from "../loc/fr-fr.js";

export function SelectLanguage(lang: string): ISpfxReactDirectoryWebpartWebPartStrings {
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
