import { BANK_URL, BANK_VERSION, APP_VERSION } from "./config.js";
import { loadQuestionBank } from "./data/question-repository.js";
import { openProgressStore, createMemoryStore } from "./data/progress-store.js";
import { createBackup, parseBackup } from "./data/backup.js";
import { gradeQuestion } from "./domain/questions.js";
import { createSession } from "./domain/session.js";
import { applyAttempt, buildStats } from "./domain/progress.js";
import * as view from "./ui/screens.js";
import { selected, shouldSubmitOnSelection } from "./ui/controller.js";

const root=document.querySelector("#app");
const state={questions:[],records:[],store:null,persistent:true,session:[],mode:"",index:0,locked:false,lastCorrect:false,results:[]};
const renderHome=()=>root.innerHTML=view.home(buildStats(state.records,state.questions.length));
const chapters=()=>[...new Map(state.questions.map(q=>[q.chapterId,{id:q.chapterId,title:q.chapterTitle}])).values()];
function start(items,mode){if(!items.length){alert("暂无可练习题目");return;}Object.assign(state,{session:items,mode,index:0,locked:false,results:[]});renderQuiz();}
function renderQuiz(){root.innerHTML=view.quiz(state.session[state.index],state.index,state.session.length,state.locked,state.lastCorrect);}
async function submit(){const answers=selected();if(!answers.length){alert("请先选择答案");return;}const q=state.session[state.index];state.lastCorrect=gradeQuestion(q,answers);state.results.push(state.lastCorrect);const old=state.records.find(r=>r.questionId===q.id);const record=applyAttempt(old,{questionId:q.id,selected:answers,isCorrect:state.lastCorrect,mode:state.mode,at:new Date().toISOString()});state.records=state.records.filter(r=>r.questionId!==q.id).concat(record);await state.store.put(record);state.locked=true;renderQuiz();}
function next(){if(state.index+1>=state.session.length){root.innerHTML=view.results(state.results);return;}state.index++;state.locked=false;renderQuiz();}
function download(){const blob=new Blob([createBackup(state.records,{exportedAt:new Date().toISOString(),bankVersion:BANK_VERSION})],{type:"application/json"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=`毛概刷题备份-${new Date().toISOString().slice(0,10)}.json`;a.click();URL.revokeObjectURL(a.href);}
root.addEventListener("click",async event=>{const action=event.target.closest("[data-action]")?.dataset.action;if(!action)return;
  if(action==="home")renderHome();
  if(action==="chapters")root.innerHTML=view.setup(chapters());
  if(action==="random")root.innerHTML=view.setup(chapters(),true);
  if(action==="start-chapter"){const id=document.querySelector("#chapter").value,type=document.querySelector("#type").value;start(createSession(state.questions,{mode:"chapter",chapterId:id,types:type?[type]:[]}),"chapter");}
  if(action==="start-random")start(createSession(state.questions,{mode:"random",count:Number(document.querySelector("#count").value)}),"random");
  if(action==="wrong")start(createSession(state.questions,{mode:"wrong",wrongIds:state.records.filter(r=>r.isWrong).map(r=>r.questionId)}),"wrong");
  if(action==="submit")await submit();
  if(action==="next")next();
  if(action==="settings")root.innerHTML=view.settings(state.persistent,BANK_VERSION);
  if(action==="export")download();
  if(action==="clear"&&confirm("确定清空全部学习记录吗？")){await state.store.clear();state.records=[];renderHome();}
});
root.addEventListener("change",async event=>{
  if(event.target.name==="answer"&&!state.locked&&shouldSubmitOnSelection(state.session[state.index]?.type)){await submit();return;}
  if(event.target.id!=="import")return;
  try{const backup=parseBackup(await event.target.files[0].text(),BANK_VERSION);await state.store.replaceAll(backup.records);state.records=backup.records;alert("学习记录已恢复");renderHome();}catch(error){alert(error.message);}
});

try{state.questions=await loadQuestionBank(BANK_URL,BANK_VERSION);try{state.store=await openProgressStore();state.records=await state.store.all();}catch{state.persistent=false;state.store=createMemoryStore();}renderHome();}catch(error){root.innerHTML=`<main class="screen"><h1>无法启动</h1><p>${error.message}</p></main>`;}
if("serviceWorker" in navigator)window.addEventListener("load",()=>navigator.serviceWorker.register("./service-worker.js"));
