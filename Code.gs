/**
 * Apps Script backend for Inglês para Trabalhar nos EUA.
 * Set Script Property OPENAI_API_KEY before deployment.
 * Apps Script > Project Settings > Script Properties.
 */
const SHEET_ID = "COLE_ID_DA_PLANILHA_AQUI";
const OPENAI_BASE = "https://api.openai.com/v1";

function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (d.action === "chat") return json_(chat_(d));
    if (d.action === "tts") return json_(tts_(d));
    return json_(sync_(d));
  } catch (err) {
    return json_({ok:false, error:String(err && err.message || err)});
  }
}
function apiKey_() {
  const k = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!k) throw new Error("Configure OPENAI_API_KEY nas Script Properties.");
  return k;
}
function chat_(d) {
  const key=apiKey_();
  const topic=d.topic||"Inglês geral", level=d.level||"Iniciante";
  const system="You are a friendly American English teacher for a Brazilian Portuguese-speaking adult learning to work in the United States. Topic: "+topic+". Student level: "+level+". Reply primarily in natural American English, then give a concise Portuguese explanation when useful. Correct the student's mistakes gently: show corrected phrase, short reason in Brazilian Portuguese, and ask one natural follow-up question to keep the conversation going. Roleplay realistic professional situations when topic is a job. Avoid long lectures. Never claim to assess phonemes from text.";
  const payload={model:"gpt-4o-mini",messages:[{role:"system",content:system}].concat((d.messages||[]).slice(-20)),temperature:0.7};
  const r=UrlFetchApp.fetch(OPENAI_BASE+"/chat/completions",{method:"post",contentType:"application/json",headers:{Authorization:"Bearer "+key},payload:JSON.stringify(payload),muteHttpExceptions:true});
  const status=r.getResponseCode(), body=JSON.parse(r.getContentText()||"{}");
  if(status<200||status>=300) throw new Error(body.error&&body.error.message||"OpenAI chat HTTP "+status);
  return {ok:true,reply:body.choices[0].message.content};
}
function tts_(d) {
  const key=apiKey_(), text=String(d.text||"").slice(0,4000);
  if(!text) throw new Error("Texto vazio.");
  const payload={model:"gpt-4o-mini-tts",voice:"alloy",input:text, instructions:"Speak naturally in clear American English at a learner-friendly pace."};
  const r=UrlFetchApp.fetch(OPENAI_BASE+"/audio/speech",{method:"post",contentType:"application/json",headers:{Authorization:"Bearer "+key},payload:JSON.stringify(payload),muteHttpExceptions:true});
  const status=r.getResponseCode();
  if(status<200||status>=300) throw new Error("OpenAI TTS HTTP "+status+": "+r.getContentText().slice(0,500));
  return {ok:true,audio:Utilities.base64Encode(r.getBlob().getBytes()),mime:"audio/mpeg"};
}
function sync_(d) {
  if(SHEET_ID==="COLE_ID_DA_PLANILHA_AQUI") throw new Error("Configure SHEET_ID no Code.gs.");
  const ss=SpreadsheetApp.openById(SHEET_ID);
  let sh=ss.getSheetByName("Progresso");if(!sh)sh=ss.insertSheet("Progresso");
  if(sh.getLastRow()===0)sh.appendRow(["Data","Aluno","Dados JSON"]);
  sh.appendRow([new Date(),d.student||"Aluno",JSON.stringify(d.state||d)]);
  return {ok:true};
}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
