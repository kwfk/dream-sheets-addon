// const API_SERVER_URL = "http://m1.zamfi.net:8080/";
const API_SERVER_URL = "http://dreamsheets.zamfi.net/";

const IMAGE_NOT_FOUND = `${API_SERVER_URL}ftp/image-not-found.png`;
const DEFAULT_SEED = "1234";
const DEFAULT_CFG = "13";
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
    .addItem("Get document ID", "_showDocumentId")
    .addItem("Set/view global seed", "_setGlobalSeed")
    .addSeparator()
    .addItem("Show prompt", "_showPrompt")
    .addItem("Download image", "_downloadImage")
    .addSeparator()
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

function _getDocumentId() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const id = sheet.getId();
  return id;
}

function _showDocumentId() {
  SpreadsheetApp.getUi().alert(_getDocumentId());
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

type TTIParams = {
  prompt: string;
  seed: string | number;
  guidance: string | number;
};

function _getParams(
  prompt: string,
  seed: number | string = "",
  guidance: number | string = ""
) {
  const params: TTIParams = {
    prompt,
    seed,
    guidance,
  };
  return JSON.stringify(params);
}

function _getPrompt(): TTIParams {
  // get the params for TTI, pass them into a different function, get their values
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  const value = cell?.getValue().toString();

  // check if value === CellImage
  if (value !== "CellImage") {
    throw new Error("Selected cell is not an image");
  }

  // check if formula is =IMAGE(TTI())
  const formula = cell?.getFormula();
  if (formula === undefined) {
    throw new Error("Selected cell is not a formula");
  }

  const regex = /=IMAGE\(TTI\((.*?)\)\)/i;
  const matches = formula.match(regex);

  if (matches && matches.length >= 2) {
    const rawPrompt = matches[1];

    const prompt: string = cell
      .setValue("=_getParams(" + rawPrompt + ")")
      .getValue()
      .toString();
    cell.setValue(formula);

    return JSON.parse(prompt);
  } else {
    throw new Error("Selected cell must use TTI formula");
  }
}

function _showPrompt() {
  const loc = SpreadsheetApp.getCurrentCell().getA1Notation();
  const prompt = _getPrompt();

  let seed =
    PropertiesService.getDocumentProperties().getProperty("seed") ||
    DEFAULT_SEED;
  if (prompt.seed) {
    seed = prompt.seed.toString();
  }

  let cfg = DEFAULT_CFG;
  if (prompt.guidance) {
    cfg = prompt.guidance.toString();
  }

  const htmlOutput = HtmlService.createHtmlOutput(
    `<p>Prompt: ${prompt.prompt}</p><p>Seed: ${seed}</p><p>Guidance: ${cfg}</p>`
  )
    .setWidth(250)
    .setHeight(250);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, `Prompt at ${loc}`);
}

function _downloadImage() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  const value = cell?.getValue().toString();

  // check if value === CellImage
  if (value !== "CellImage") {
    throw new Error("Selected cell is not an image");
  }

  // check if formula is =IMAGE(TTI())
  const formula = cell?.getFormula();
  if (formula === undefined) {
    throw new Error("Selected cell is not a formula");
  }

  const regex = /=IMAGE\((TTI\(.*?\))\)/i;
  const matches = formula.match(regex);

  if (matches && matches.length >= 2) {
    const tti = matches[1];

    const url: string = cell
      .setValue("=" + tti)
      .getValue()
      .toString();
    cell.setValue(formula);

    const htmlOutput = HtmlService.createHtmlOutput(
      `<div style="overflow-wrap: anywhere;"><a href="${url}">${url}</a></div>`
    )
      .setWidth(250)
      .setHeight(150);
    SpreadsheetApp.getUi().showModelessDialog(
      htmlOutput,
      `URL at ${cell.getA1Notation()}`
    );
  } else {
    throw new Error("Selected cell must use TTI formula");
  }
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

/**
 * Generates an image using a text-to-image model
 * @param prompt
 * @param seed
 * @param guidance
 * @returns An image URL of the generated image
 * @customfunction
 */
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
  const id = _getDocumentId();

  const response = UrlFetchApp.fetch(
    `${API_SERVER_URL}?prompt=${encodedPrompt}&seed=${reqSeed}&cfg=${cfg}&id=${id}`,
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
/**
 * An open-ended function to GPT-3
 * @param prompt
 * @param stop
 * @returns
 * @customfunction
 */
function GPT(prompt, stop = "") {
  Logger.log("prompt: %s", prompt);
  _validatePrompt(prompt);

  const encodedPrompt = encodeURIComponent(prompt);
  let params = `prompt=${encodedPrompt}`;
  if (stop) {
    params += `&stop=${stop}`;
  }
  const id = _getDocumentId();
  const response = UrlFetchApp.fetch(
    `${API_SERVER_URL}gpt?${params}&id=${id}`,
    {
      muteHttpExceptions: true,
    }
  );
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

/**
 * Gives list of items based on the given prompt
 * @param prompt
 * @param length
 * @param transpose
 * @returns
 * @customfunction
 */
function GPT_LIST(prompt, length = 5, transpose = false) {
  Logger.log("prompt: %s", prompt);
  _validatePrompt(prompt);
  _validateLengthAndTranspose(length, transpose);

  const encodedPrompt = encodeURIComponent(prompt);
  let params = `prompt=${encodedPrompt}&length=${length}`;
  const id = _getDocumentId();
  const res = UrlFetchApp.fetch(`${API_SERVER_URL}listgpt?${params}&id=${id}`, {
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
/**
 * Transposed version of GPT_LIST
 * @param prompt
 * @param length
 * @returns
 * @customfunction
 */
function GPT_LIST_T(prompt, length = 5) {
  return GPT_LIST(prompt, length, true);
}

/**
 * Gives list of items with elements similar to given list
 * @param prompt
 * @param length
 * @param transpose
 * @returns
 * @customfunction
 */
function LIST_COMPLETION(prompt, length = 5, transpose = false) {
  return GPT_LIST(
    `similar items to this list without repeating "[${prompt}]"`,
    length,
    transpose
  );
}
/**
 * Transposed version of LIST_COMPLETION
 * @param prompt
 * @param length
 * @returns
 * @customfunction
 */
function LIST_COMPLETION_T(prompt, length = 5) {
  return LIST_COMPLETION(prompt, length, true);
}

/**
 * Gives list of synonyms of given word
 * @param prompt
 * @param length
 * @param transpose
 * @returns
 * @customfunction
 */
function SYNONYMS(prompt, length = 5, transpose = false) {
  return GPT_LIST(`synonyms of "${prompt}"`, length, transpose);
}
/**
 * Transposed version of SYNONYMS
 * @param prompt
 * @param length
 * @returns
 * @customfunction
 */
function SYNONYMS_T(prompt, length = 5) {
  return SYNONYMS(prompt, length, true);
}

/**
 * Gives list of antonyms of given word
 * @param prompt
 * @param length
 * @param transpose
 * @returns
 * @customfunction
 */
function ANTONYMS(prompt, length = 5, transpose = false) {
  return GPT_LIST(`antonyms of "${prompt}"`, length, transpose);
}
/**
 * Transposed version of ANTONYMS
 * @param prompt
 * @param length
 * @returns
 * @customfunction
 */
function ANTONYMS_T(prompt, length = 5) {
  return ANTONYMS(prompt, length, true);
}

/**
 * Gives list of different but related words to given word
 * @param prompt
 * @param length
 * @param transpose
 * @returns
 * @customfunction
 */
function DIVERGENTS(prompt, length = 5, transpose = false) {
  return GPT_LIST(`divergent words to "${prompt}"`, length, transpose);
}
/**
 * Transposed version of DIVERGENTS
 * @param prompt
 * @param length
 * @returns
 * @customfunction
 */
function DIVERGENTS_T(prompt, length = 5) {
  return DIVERGENTS(prompt, length, true);
}

/**
 * Gives list of alternative ways to say the given word or phrase
 * @param {string} prompt A word or phrase
 * @param {number} length Number of alternatives
 * @returns The alternative ways to say the word or phrase
 * @customfunction
 */
function ALTERNATIVES(prompt, length = 5, transpose = false) {
  return GPT_LIST(`alternative ways to say "${prompt}"`, length, transpose);
}
/**
 * Transposed version of ALTERNATIVES
 * @param prompt
 * @param length
 * @returns
 * @customfunction
 */
function ALTERNATIVES_T(prompt, length) {
  return ALTERNATIVES(prompt, length, true);
}

/**
 * Embellishes a phrase
 * @param {string} prompt A phrase to embellish
 * @returns The embellished phrase
 * @customfunction
 */
function EMBELLISH(prompt) {
  _validatePrompt(prompt);
  prompt = `Embellish this sentence: ${prompt}`;
  const res = GPT(prompt);
  return res;
}
