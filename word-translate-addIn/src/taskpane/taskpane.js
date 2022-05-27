/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
// npm cache --force clean  
// npm install --force
// const translate = require("@vitalets/google-translate-api");
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("insert").onclick =  writeData();
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, async (eventArgs) => {
      // await translates();
    });
  }
});
// const API_KEY = '9ac32b0592msh11de28e6258e1cap14cf7fjsn332299a82fc1';
// const API_KEY = "93d67f0abamsh53f0dd7a562f07cp12fe30jsnb6ce5425e9ec";
const select = document.getElementById("selectCountry");
const textArea = document.getElementById("textTranSlated");

function writeData(){
  Office.context.document.setSelectedDataAsync(textArea.value, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      write(asyncResult.error.message);
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
  let text = await getSelectionText();
  console.log(text);
  // if(text ===''){
  //   console.log('empty')
  // }else{
  //   console.log(text);
  //   return text;
  // }
}




async function autoDetect(textSlection){
  const encodedParams = new URLSearchParams();
  encodedParams.append("q", textSlection);

  const options = {
    method: "POST",
    headers: {
      "content-type": "application/x-www-form-urlencoded",
      "Accept-Encoding": "application/gzip",
      "X-RapidAPI-Host": "google-translate1.p.rapidapi.com",
      "X-RapidAPI-Key": API_KEY,
    },
    body: encodedParams,
  };

  const objectDectect= fetch("https://google-translate1.p.rapidapi.com/language/translate/v2/detect", options)
  .then((result) => result.json()).then((result) =>{return result.data.detections[0][0].language} );
 return objectDectect;
}



async function translate(){
  const detectSource = await autoDetect();
  let optionTarget = select.value;
  let textSelection = await checkSelectedText();
  const encodedParams = new URLSearchParams();
  encodedParams.append("q", textSelection);
  encodedParams.append("target", optionTarget);
  encodedParams.append("source", detectSource);
  const options = {
    method: "POST",
    headers: {
      "content-type": "application/x-www-form-urlencoded",
      "Accept-Encoding": "application/gzip",
      "X-RapidAPI-Host": "google-translate1.p.rapidapi.com",
      "X-RapidAPI-Key": API_KEY,
    },
    body: encodedParams,
  };

   fetch("https://google-translate1.p.rapidapi.com/language/translate/v2", options)
    .then((response) => response.json())
    .then((response) => {
      var result = response.data.translations[0].translatedText;
      textArea.innerHTML = result;
    })
    .catch((err) => console.error(err));
}

// function googleTranslateElementInit() {
//   new google.translate.TranslateElement({ pageLanguage: "en" }, "vi");
// }

 
async function translates(){
 let text = await readData();
 const options = {
   method: "POST",
   headers: {
     "content-type": "application/json",
     "X-RapidAPI-Host": "microsoft-translator-text.p.rapidapi.com",
     "X-RapidAPI-Key": "9ac32b0592msh11de28e6258e1cap14cf7fjsn332299a82fc1",
   },
   body: '[{"Text":"hello"}]',
 };

 fetch(
   "https://microsoft-translator-text.p.rapidapi.com/translate?to%5B0%5D=en&api-version=3.0&profanityAction=NoAction&textType=plain",
   options
 )
   .then((response) => response.json())
   .then((response) => console.log(response))
   .catch((err) => console.error(err));
}
translates();
 
 function readData(){
 Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
   if (asyncResult.status == Office.AsyncResultStatus.Failed) {
     console.log("Action failed. Error: " +  asyncResult.error.message);
   } else {
     return  asyncResult.value;
   }
 });

}
