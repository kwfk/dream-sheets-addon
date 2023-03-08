const API_SERVER_URL = "http://m1.zamfi.net:8080/"

function onInstall(e) {
  PropertiesService.getDocumentProperties().setProperty('seed', 1234)
  onOpen(e);
}

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Spreadsheet Diffusion')
      .addItem('Set/view global seed', 'setGlobalSeed')
      .addItem('Show sidebar', 'showSidebar')
      .addItem('Download image', 'downloadImageFromURL')
      .addToUi();
}

function setGlobalSeed() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  const globalSeed = parseInt(PropertiesService.getDocumentProperties().getProperty('seed'))
  const result = ui.prompt(
      'Current global seed is: ' + globalSeed,
      'Set a global seed:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  const button = result.getSelectedButton();
  const seed = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    const seedNum = parseInt(seed)
    if (!seedNum) {
      ui.alert('Seed must be an integer')
    } else {
      PropertiesService.getDocumentProperties().setProperty('seed', seedNum)
    }
  }
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Page')
      .setTitle('Spreadsheet Diffusion');
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}

function downloadImageFromURL() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const value = range.getValue();
  console.log(value, value.getUrl())
  return value;
}

function FETCH_IMAGE(ready) {
  const rand = Math.random()
  console.log(rand)
  return 'https://www.google.com/images/srpr/logo3w.png'
  if (ready) {
    if (rand > 0.5) {
      return 'https://zamfi.net/img/me.jpg'
    } else {
      return 'https://www.google.com/images/srpr/logo3w.png'
    }
  } else {
    return 'https://handsontable.com/docs/img/examples/javascript-the-good-parts.jpg'
  }
}

function TTI(prompt, seed=null) {
  let reqSeed = parseInt(PropertiesService.getDocumentProperties().getProperty('seed'))
  if (seed) {
    reqSeed = seed
  }

  Logger.log("prompt: %s", prompt)
  const encodedPrompt = encodeURIComponent(prompt)
  const response = UrlFetchApp.fetch(`${API_SERVER_URL}?prompt=${encodedPrompt}&seed=${reqSeed}`)
  Logger.log("first cut: %s", response.getResponseCode());
  if (response.getResponseCode() !== 200) {
    return 'null-image'
  } else {
    return response.getContentText()
  }
}

function GPT(prompt, stop="") {
  Logger.log("prompt: %s", prompt)
  const encodedPrompt = encodeURIComponent(prompt)
  let params = `prompt=${encodedPrompt}`
  if (stop) {
    params += `&stop=${stop}`
  }
  const response = UrlFetchApp.fetch(`${API_SERVER_URL}gpt?${params}`);
  Logger.log("gpt res: %s", response.getContentText(), response.getResponseCode());
  if (response.getResponseCode() !== 200) {
    return 'null-completion'
  } else {
    return response.getContentText()
  }
}

function GPT_LIST(prompt, length=5, transpose=false) {
  prompt = `Javascript array literal length ${length} with ${prompt} ["`

  const res = GPT(prompt, "%5D")
  const list = JSON.parse(`["${res}]`)

  if (transpose) {
    return [list]
  }
  return list
}

function GPT_LIST_T(prompt, length=5) {
  return GPT_LIST(prompt, length, true)
}

function LIST_COMPLETION(prompt, length=5, transpose=false) {
  prompt = `Extend the Javascript array literal with ${length+prompt.length} similar items [${prompt}, "`
  const res = GPT(prompt, "%5D")
  const list = JSON.parse(`["${res}]`)
  if (transpose) {
    return [list]
  }
  return list
}


function SYNONYM(prompt, length=5, transpose=false) {
  prompt = `Javascript array literal length ${length} with synonyms of ${prompt} ["`
  const res = GPT(prompt, "%5D")
  const list = JSON.parse(`["${res}]`)
  if (transpose) {
    return [list]
  }
  return list
}

function ANTONYM(prompt, length=5, transpose=false) {
  prompt = `Javascript array literal length ${length} with antonyms of ${prompt} ["`
  const res = GPT(prompt, "%5D")
  const list = JSON.parse(`["${res}]`)
  if (transpose) {
    return [list]
  }
  return list
}

function ALTERNATIVE(prompt, length=5, transpose=false) {
  prompt = `Javascript array literal length ${length} of alternative ways to say "${prompt}" ["`
  const res = GPT(prompt, "%5D")
  const list = JSON.parse(`["${res}]`)
  if (transpose) {
    return [list]
  }
  return list
}
