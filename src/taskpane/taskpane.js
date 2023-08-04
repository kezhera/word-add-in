/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */


const CheckTextTest = async (text) => {
  const res = await fetch("https://enagramm.com/API/SpellChecker/CheckTextTest", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      Text: text,
      ToolsGlobalLangID: 1,
      ToolsLangCode: "GEO"
    })
  }).then(r => r.json())
  return res.lstWR
}

const getTextItem = (text) => {
  return `<li>${text}</li>`
}

const showText = (text) => {
  let html = ""
  text.forEach(word => {
    if(word.Correct){
      html += getTextItem(word.Word)
    }
  });
  document.querySelector(".word-list").innerHTML = ""
  document.querySelector(".word-list").innerHTML = html
}

function extractWordsFromText(text) {
  const words = text.split(/\s+/); // Split by whitespace to get words
  return words;
}
/* global document, Office, Word */

Office.onReady((info) => {
  run()
});

export async function run() {
  return Word.run(async (context) => {
    const documentBody = context.document.body
    context.load(documentBody);
    const text = await context.sync().then(() => documentBody.text);
    const chechedTextArr = await CheckTextTest(text)
    console.log("chechedTextArr", chechedTextArr)
    showText(chechedTextArr)
    
    const paragraphs = context.document.body.paragraphs;
    context.load(paragraphs, 'text');

    context.sync()
      .then(function () {
        let allWords = [];
        for (let paragraph of paragraphs.items) {
          const wordsInParagraph = extractWordsFromText(paragraph.text);
          allWords = allWords.concat(wordsInParagraph);
        }

        console.log("allWords", allWords)
      });
  });
}
