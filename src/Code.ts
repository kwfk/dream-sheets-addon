const API_SERVER_URL = "http://m1.zamfi.net:8080/";
const IMAGE_NOT_FOUND = `${API_SERVER_URL}ftp/image-not-found.png`;
const DEFAULT_SEED = "1234";
const ERRORS = [
  "#REF!",
  "#ERROR!",
  "#NAME?",
  "#DIV/0",
  "#VALUE!",
  "#NUM!",
  "#NULL!",
  "#N/A",
];
type SPREADSHEET_INPUT = string | number | boolean | Date;

function _validatePrompt(prompt) {
  if (ERRORS.includes(prompt)) {
    throw `Invalid input: ${prompt} error was passed as prompt`;
  }
  if (typeof prompt !== "string") {
    console.log(typeof prompt);
    throw `Invalid input: prompt is not a string`;
  }
  if (prompt === "") {
    throw `Invalid input: prompt is empty`;
  }
}

function _validateLengthAndTranspose(length, transpose) {
  if (length && typeof length !== "number")
    throw `Invalid input: length=${length} is not a number`;
  if (transpose && typeof transpose !== "boolean")
    throw `Invalid input: tranpose=${transpose} is not a boolean`;
}

function onInstall(e: GoogleAppsScript.Events.AddonOnInstall) {
  PropertiesService.getDocumentProperties().setProperty("seed", DEFAULT_SEED);
  // @ts-ignore
  onOpen(e);
}

function onOpen(e: GoogleAppsScript.Events.SheetsOnOpen) {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .createMenu("Dream Sheets")
    .addItem("Set/view global seed", "_setGlobalSeed")
    .addItem("Show sidebar", "_showSidebar")
    .addItem("Download image", "downloadImageFromURL")
    .addItem("Rerun selected cell(s)", "_rerun")
    .addItem("Rerun all TTI", "_rerunAllTTIConfirm")
    .addToUi();
}

function _setGlobalSeed() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  const globalSeed = parseInt(
    PropertiesService.getDocumentProperties().getProperty("seed") ||
      DEFAULT_SEED
  );
  const result = ui.prompt(
    "Current global seed is: " + globalSeed,
    "Set a global seed:\n(Changing the global seed will rerun all TTI functions)",
    ui.ButtonSet.OK_CANCEL
  );

  // Process the user's response.
  const button = result.getSelectedButton();
  const seed = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    const seedNum = parseInt(seed);
    if (!seedNum) {
      ui.alert("Seed must be an integer");
    } else if (seedNum === globalSeed) {
      ui.alert("Given seed is the same as current global seed");
    } else {
      PropertiesService.getDocumentProperties().setProperty(
        "seed",
        seedNum.toString()
      );
      _rerunAllTTI();
    }
  }
}

function _showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Page").setTitle(
    "Spreadsheet Diffusion"
  );
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showSidebar(html);
}

function _rerunAllTTI() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getDataRange();
  const formulas = range?.getFormulas();
  Logger.log(formulas);

  for (let r = 0; r < formulas.length; r++) {
    for (let c = 0; c < formulas[r].length; c++) {
      const formula = formulas[r][c];
      if (formula && /TTI\(/i.test(formula)) {
        const curr = sheet.getRange(r + 1, c + 1);
        curr.clear();
        SpreadsheetApp.flush();
        curr.setFormula(formula);
      }
    }
  }
}

function _rerunAllTTIConfirm() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    "Are you sure you want to rerun all TTI functions?",
    ui.ButtonSet.YES_NO
  );

  if (result == ui.Button.YES) {
    _rerunAllTTI();
  }
}

function _rerun() {
  const range = SpreadsheetApp.getActiveRange();
  const formulas = range?.getFormulas();
  const values = range?.getValues();
  let copyVals: string[][] = [];

  for (let r = 0; r < formulas.length; r++) {
    let row: string[] = [];
    for (let c = 0; c < formulas[r].length; c++) {
      const formula = formulas[r][c];
      if (formula) {
        row.push(formula);
      } else {
        const val = values[r][c];
        row.push(val);
      }
    }
    copyVals.push(row);
  }

  range.clear();
  SpreadsheetApp.flush();
  Logger.log(copyVals);
  range.setValues(copyVals);
}

function downloadImageFromURL() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const value = range?.getValue();
  console.log(value, value.getUrl());
  return value;
}

function FETCH_IMAGE(ready) {
  const rand = Math.random();
  console.log(rand);
  return "https://www.google.com/images/srpr/logo3w.png";
  if (ready) {
    if (rand > 0.5) {
      return "https://zamfi.net/img/me.jpg";
    } else {
      return "https://www.google.com/images/srpr/logo3w.png";
    }
  } else {
    return "https://handsontable.com/docs/img/examples/javascript-the-good-parts.jpg";
  }
}

function test(input) {
  console.log(input, typeof input);
  // const range = SpreadsheetApp.getActiveRange();
  // const note = range.getNote();
  // const location = range.getA1Notation();
  // console.log(location, note);
}

function TTI(
  prompt,
  seed: number | string = "",
  guidance: number | string = ""
) {
  const cfg = guidance;
  Logger.log(`input) seed: ${seed}; cfg: ${cfg}; prompt: ${prompt}`);
  _validatePrompt(prompt);

  // if no seed arg, use global seed
  let reqSeed = parseInt(
    PropertiesService.getDocumentProperties().getProperty("seed") ||
      DEFAULT_SEED
  );
  if (seed) {
    if (typeof seed === "number") reqSeed = seed;
    else throw `Invalid input: seed=${seed} is not a number`;
  }
  if (cfg) {
    if (typeof cfg !== "number") {
      throw `Invalid input: guidance=${cfg} is not a number`;
    } else if (cfg < 0 && cfg > 35) {
      throw `Invalid input: guidance=${cfg} must be between 0 and 35`;
    }
  }

  const encodedPrompt = encodeURIComponent(prompt);
  // const seedAndPrompt = `seed=${reqSeed}&${encodedPrompt}`;

  const response = UrlFetchApp.fetch(
    `${API_SERVER_URL}?prompt=${encodedPrompt}&seed=${reqSeed}&cfg=${cfg}`,
    {
      muteHttpExceptions: true,
    }
  );
  if (response.getResponseCode() !== 200) {
    const errorMsg = response.getContentText();
    const errorCode = response.getResponseCode();
    Logger.log("server fail (%s)", errorCode);
    throw `(${errorCode}) ${errorMsg}`;
  } else {
    const url = response.getContentText();
    Logger.log("url: %s", url);
    return url;
  }
}

// GPT FORMULAS =============================================
function GPT(prompt, stop = "") {
  Logger.log("prompt: %s", prompt);
  _validatePrompt(prompt);

  const encodedPrompt = encodeURIComponent(prompt);
  let params = `prompt=${encodedPrompt}`;
  if (stop) {
    params += `&stop=${stop}`;
  }
  const response = UrlFetchApp.fetch(`${API_SERVER_URL}gpt?${params}`, {
    muteHttpExceptions: true,
  });
  Logger.log(
    "gpt res: %s",
    response.getContentText(),
    response.getResponseCode()
  );
  if (response.getResponseCode() !== 200) {
    const errorMsg = response.getContentText();
    const errorCode = response.getResponseCode();
    Logger.log("server fail (%s)", errorCode);
    throw `(${errorCode}) ${errorMsg}`;
  } else {
    return response.getContentText();
  }
}

function GPT_LIST(prompt, length = 5, transpose = false) {
  Logger.log("prompt: %s", prompt);
  _validatePrompt(prompt);
  _validateLengthAndTranspose(length, transpose);

  const encodedPrompt = encodeURIComponent(prompt);
  let params = `prompt=${encodedPrompt}&length=${length}`;
  const res = UrlFetchApp.fetch(`${API_SERVER_URL}listgpt?${params}`, {
    muteHttpExceptions: true,
  });
  Logger.log(`list gpt (${res.getResponseCode()}): ${res.getContentText()}`);

  if (res.getResponseCode() !== 200) {
    const errMsg = res.getContentText();
    const errCode = res.getResponseCode();
    throw `(${errCode}) ${errMsg}`;
  } else {
    let list: string[];
    try {
      list = JSON.parse(res.getContentText());
    } catch (err) {
      throw `GPT failed to generate a proper list, try another response`;
    }

    if (transpose) return [list];
    return list;
  }
}

function GPT_LIST_T(prompt, length = 5) {
  return GPT_LIST(prompt, length, true);
}

function LIST_COMPLETION(prompt, length = 5, transpose = false) {
  return GPT_LIST(
    `similar items to this list without repeating "[${prompt}]"`,
    length,
    transpose
  );
}

function LIST_COMPLETION_T(prompt, length = 5) {
  return LIST_COMPLETION(prompt, length, true);
}

function SYNONYMS(prompt, length = 5, transpose = false) {
  return GPT_LIST(`synonyms of "${prompt}"`, length, transpose);
}
function SYNONYMS_T(prompt, length = 5) {
  return SYNONYMS(prompt, length, true);
}

function ANTONYMS(prompt, length = 5, transpose = false) {
  return GPT_LIST(`antonyms of "${prompt}"`, length, transpose);
}
function ANTONYMS_T(prompt, length = 5) {
  return ANTONYMS(prompt, length, true);
}

function DIVERGENTS(prompt, length = 5, transpose = false) {
  return GPT_LIST(`divergent words to "${prompt}"`, length, transpose);
}
function DIVERGENTS_T(prompt, length = 5) {
  return DIVERGENTS(prompt, length, true);
}

function ALTERNATIVES(prompt, length = 5, transpose = false) {
  return GPT_LIST(`alternative ways to say "${prompt}"`, length, transpose);
}
function ALTERNATIVES_T(prompt, length) {
  return ALTERNATIVES(prompt, length, true);
}

function EMBELLISH(prompt) {
  _validatePrompt(prompt);
  prompt = `Embellish this sentence: ${prompt}`;
  const res = GPT(prompt);
  return res;
}
