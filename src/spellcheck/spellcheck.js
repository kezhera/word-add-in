const CheckTextTest = async (text) => {
  const cleanedText = text.replace(/[\s\u000b]/g, " ");
  
  const body = JSON.stringify({
    Text: cleanedText,
    ToolsGlobalLangID: 1,
    ToolsLangCode: "GEO"
  })

  const res = await fetch("https://enagramm.com/API/SpellChecker/CheckTextTest", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: body
  }).then(r => r.json())
  return res.lstWR
}

const showWord = (word, context) => {
  const container = document.querySelector(".word-list")
  
  const wordElement = document.createElement("li");
  wordElement.textContent = word.text; 

  container.appendChild(wordElement);

  wordElement.addEventListener("click", async () => {
    word.select()
    await context.sync()
  })
  return wordElement
}

function getWordsFromText(text) {
  const specialChars = /[ „“`!@#$%^&*()_+\-=\[\]{};':"\\|,.<>\/?~]/;
  const words = text.split(specialChars).filter(s => s != "");  
  return words;
}

Office.onReady((info) => {
  document.getElementById("check-text").onclick = checkText
});

export const getDocumentText = async (context) => {
  const documentBody = context.document.body
  context.load(documentBody);
  await context.sync()
  return documentBody.text
}

export const getDocumentWordsByGetRange = async (context) =>  {
  const body = context.document.body.getRange().getTextRanges([' '], true);
  context.load(body, ['text', 'font']);
  await context.sync(); 
  return body.items
}

export const showCorretWords = async (words,correctWords, context) => {
  document.querySelector(".word-list").innerHTML = ""
  
  for (let i = 0; i < words.length; i += 1) {
    const word = words[i];
    
    if (correctWords[i].Correct){
      word.font.highlightColor = 'yellow'
      showWord(word, context) 
    }else{
      word.font.highlightColor = null
    };
  }

  await context.sync(); 
}

export async function checkText() {
  return Word.run(async (context) => {
    // const text = await getDocumentText(context)
    const documentWords = await getDocumentWordsByGetRange(context)
    const text = documentWords.map((word) => word.text).join(" ")
    const chechedWordArr = await CheckTextTest(text)

    await showCorretWords(documentWords, chechedWordArr, context)
  });
}
