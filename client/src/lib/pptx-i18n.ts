import { createInstance } from "i18next";
import { keyToLabel, translationsEn } from "pptx-react-viewer/i18n";
import { initReactI18next } from "react-i18next";

const pptxI18n = createInstance();

void pptxI18n.use(initReactI18next).init({
  resources: {
    en: {
      translation: translationsEn,
    },
  },
  lng: "en",
  fallbackLng: "en",
  interpolation: {
    escapeValue: false,
  },
  parseMissingKeyHandler: (key) => keyToLabel(key),
  missingKeyHandler: false,
});

export default pptxI18n;
