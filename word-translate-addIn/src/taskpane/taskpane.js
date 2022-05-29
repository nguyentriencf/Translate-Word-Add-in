/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
// npm cache --force clean  
// npm install --force
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("insert").onclick = writeData;
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, async (eventArgs) => {
      await translates();
    });
  }
});

const API_KEY = "b0ac586fa2mshf0687d63e8ec41cp13f635jsn560a4554a34a";
const select = document.getElementById("selectCountry");
const textArea = document.getElementById("textTranSlated");

function writeData(){
  Office.context.document.setSelectedDataAsync(textArea.value, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      console.log(asyncResult.error.message);
    }
  });
}

const countryListAlpha2 = {
  auto: "Automatic",
  af: "Afrikaans",
  sq: "Albanian",
  am: "Amharic",
  ar: "Arabic",
  hy: "Armenian",
  az: "Azerbaijani",
  eu: "Basque",
  be: "Belarusian",
  bn: "Bengali",
  bs: "Bosnian",
  bg: "Bulgarian",
  ca: "Catalan",
  ceb: "Cebuano",
  ny: "Chichewa",
  "zh-CN": "Chinese (Simplified)",
  "zh-TW": "Chinese (Traditional)",
  co: "Corsican",
  hr: "Croatian",
  cs: "Czech",
  da: "Danish",
  nl: "Dutch",
  en: "English",
  eo: "Esperanto",
  et: "Estonian",
  tl: "Filipino",
  fi: "Finnish",
  fr: "French",
  fy: "Frisian",
  gl: "Galician",
  ka: "Georgian",
  de: "German",
  el: "Greek",
  gu: "Gujarati",
  ht: "Haitian Creole",
  ha: "Hausa",
  haw: "Hawaiian",
  he: "Hebrew",
  iw: "Hebrew",
  hi: "Hindi",
  hmn: "Hmong",
  hu: "Hungarian",
  is: "Icelandic",
  ig: "Igbo",
  id: "Indonesian",
  ga: "Irish",
  it: "Italian",
  ja: "Japanese",
  jw: "Javanese",
  kn: "Kannada",
  kk: "Kazakh",
  km: "Khmer",
  ko: "Korean",
  ku: "Kurdish (Kurmanji)",
  ky: "Kyrgyz",
  lo: "Lao",
  la: "Latin",
  lv: "Latvian",
  lt: "Lithuanian",
  lb: "Luxembourgish",
  mk: "Macedonian",
  mg: "Malagasy",
  ms: "Malay",
  ml: "Malayalam",
  mt: "Maltese",
  mi: "Maori",
  mr: "Marathi",
  mn: "Mongolian",
  my: "Myanmar (Burmese)",
  ne: "Nepali",
  no: "Norwegian",
  ps: "Pashto",
  fa: "Persian",
  pl: "Polish",
  pt: "Portuguese",
  pa: "Punjabi",
  ro: "Romanian",
  ru: "Russian",
  sm: "Samoan",
  gd: "Scots Gaelic",
  sr: "Serbian",
  st: "Sesotho",
  sn: "Shona",
  sd: "Sindhi",
  si: "Sinhala",
  sk: "Slovak",
  sl: "Slovenian",
  so: "Somali",
  es: "Spanish",
  su: "Sundanese",
  sw: "Swahili",
  sv: "Swedish",
  tg: "Tajik",
  ta: "Tamil",
  te: "Telugu",
  th: "Thai",
  tr: "Turkish",
  uk: "Ukrainian",
  ur: "Urdu",
  uz: "Uzbek",
  vi: "Vietnamese",
  cy: "Welsh",
  xh: "Xhosa",
  yi: "Yiddish",
  yo: "Yoruba",
  zu: "Zulu",
};;

function loadCountry(){
  Object.entries(countryListAlpha2).forEach(([key,value]) => {
       var option = document.createElement("option");
       key.toLowerCase() === 'vi' ? option.selected ="selected" : value;
       option.value = key.toLowerCase();
       option.text = value;
       select.add(option);
  });
}
loadCountry();

async function getSelectionText(){
   const result=  await Word.run(async (context)=>{
    let paragraph = context.document.getSelection();
    paragraph.load('text');
    await context.sync(); 
    return paragraph.text;
  });
  return result;
}
async function checkSelectedText(){
  var result =""
  let text = await getSelectionText();
  text===""?console.log('empty'): result = text;
 return result;

}


 
async function translates(){
 let text = await checkSelectedText();
 if(text!==""){
    const options = {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "X-RapidAPI-Host": "microsoft-translator-text.p.rapidapi.com",
        "X-RapidAPI-Key": API_KEY,
      },
      body: '[{"Text":"' + text + '"}]',
    };

    fetch(
      `https://microsoft-translator-text.p.rapidapi.com/translate?to%5B0%5D=${select.value}&api-version=3.0&profanityAction=NoAction&textType=plain`,
      options
    )
      .then((response) => response.json())
      .then((response) => {
        console.log(response);
        var textTranSlated = response[0].translations[0].text;

        textArea.innerHTML = textTranSlated;
      })
      .catch((err) => console.error(err));
 }
}

 


