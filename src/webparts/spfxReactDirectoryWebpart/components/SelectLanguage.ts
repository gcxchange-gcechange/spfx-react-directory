import * as strings from "SpfxReactDirectoryWebpartWebPartStrings";

const english: ISpfxReactDirectoryWebpartWebPartStrings = strings;
const french: ISpfxReactDirectoryWebpartWebPartStrings = strings;

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
