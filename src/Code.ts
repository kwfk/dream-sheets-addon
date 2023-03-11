const API_SERVER_URL = "http://m1.zamfi.net:8080/";
const IMAGE_NOT_FOUND = `${API_SERVER_URL}ftp/image-not-found.png`;
const DEFAULT_SEED = "1234";
const ERRORS = ["#REF!", "#ERROR!", "#NAME?"];

function onInstall(e: GoogleAppsScript.Events.AddonOnInstall) {
  PropertiesService.getDocumentProperties().setProperty("seed", DEFAULT_SEED);
  // @ts-ignore
  onOpen(e);
}

function onOpen(e: GoogleAppsScript.Events.SheetsOnOpen) {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .createMenu("Spreadsheet Diffusion")
    .addItem("Set/view global seed", "setGlobalSeed")
    .addItem("Show sidebar", "showSidebar")
    .addItem("Download image", "downloadImageFromURL")
    .addItem("Rerun selected cell(s)", "rerun")
    .addToUi();
}

function setGlobalSeed() {
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
      rerunAllTTI();
    }
  }
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Page").setTitle(
    "Spreadsheet Diffusion"
  );
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showSidebar(html);
}

function rerunAllTTI() {
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

function rerun() {
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
  const value = range.getValue();
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

function test() {
  console.log(
    Object.entries(PropertiesService.getDocumentProperties().getProperties())
  );
}

function TTI(prompt, seed = null) {
  // check if the prompt is a Google Sheets error message
  if (ERRORS.includes(prompt)) {
    return IMAGE_NOT_FOUND;
  }

  // if no seed arg, use global seed
  let reqSeed = parseInt(
    PropertiesService.getDocumentProperties().getProperty("seed") ||
      DEFAULT_SEED
  );
  if (seed) {
    reqSeed = seed;
  }

  const encodedPrompt = encodeURIComponent(prompt);
  // const seedAndPrompt = `seed=${reqSeed}&${encodedPrompt}`;

  const response = UrlFetchApp.fetch(
    `${API_SERVER_URL}?prompt=${encodedPrompt}&seed=${reqSeed}`
  );
  Logger.log("first cut: %s", response.getResponseCode());
  if (response.getResponseCode() !== 200) {
    return IMAGE_NOT_FOUND;
  } else {
    const url = response.getContentText();
    Logger.log("url: %s", url);
    return url;
  }
}

// GPT FORMULAS =============================================
function GPT(prompt, stop = "") {
  Logger.log("prompt: %s", prompt);
  const encodedPrompt = encodeURIComponent(prompt);
  let params = `prompt=${encodedPrompt}`;
  if (stop) {
    params += `&stop=${stop}`;
  }
  const response = UrlFetchApp.fetch(`${API_SERVER_URL}gpt?${params}`);
  Logger.log(
    "gpt res: %s",
    response.getContentText(),
    response.getResponseCode()
  );
  if (response.getResponseCode() !== 200) {
    return "null-completion";
  } else {
    return response.getContentText();
  }
}

function GPT_LIST(prompt, length = 5, transpose = false) {
  prompt = `Javascript array literal length ${length} with ${prompt} ["`;

  const res = GPT(prompt, "%5D");
  const list = JSON.parse(`["${res}]`);

  if (transpose) {
    return [list];
  }
  return list;
}

function GPT_LIST_T(prompt, length = 5) {
  return GPT_LIST(prompt, length, true);
}

function LIST_COMPLETION(prompt, length = 5, transpose = false) {
  prompt = `Extend the Javascript array literal with ${
    length + prompt.length
  } similar items [${prompt}, "`;
  const res = GPT(prompt, "%5D");
  const list = JSON.parse(`["${res}]`);
  if (transpose) {
    return [list];
  }
  return list;
}

function SYNONYM(prompt, length = 5, transpose = false) {
  prompt = `Javascript array literal length ${length} with synonyms of ${prompt} ["`;
  const res = GPT(prompt, "%5D");
  const list = JSON.parse(`["${res}]`);
  if (transpose) {
    return [list];
  }
  return list;
}

function ANTONYM(prompt, length = 5, transpose = false) {
  prompt = `Javascript array literal length ${length} with antonyms of ${prompt} ["`;
  const res = GPT(prompt, "%5D");
  const list = JSON.parse(`["${res}]`);
  if (transpose) {
    return [list];
  }
  return list;
}

function ALTERNATIVE(prompt, length = 5, transpose = false) {
  prompt = `Javascript array literal length ${length} of alternative ways to say "${prompt}" ["`;
  const res = GPT(prompt, "%5D");
  const list = JSON.parse(`["${res}]`);
  if (transpose) {
    return [list];
  }
  return list;
}

function EMBELLISH(prompt, tranpose = false) {
  prompt = `Embellish this sentence: ${prompt}`;
  const res = GPT(prompt);
  if (tranpose) {
    return [res];
  }
  return res;
}
