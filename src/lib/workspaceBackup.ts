import DOMPurify from 'dompurify';
import type { OfficeFile, FileVersion, DocumentPageSetup, DocumentComment, SlideData, SheetContent } from '../types';

const id = () => crypto.randomUUID();
export function snapshot(file: OfficeFile, name: string): FileVersion {
  return {id:id(),name:name.trim().slice(0,100)||'Saved version',createdAt:Date.now(),content:structuredClone(file.content),pageSetup:file.pageSetup?structuredClone(file.pageSetup):undefined,comments:file.comments?structuredClone(file.comments):undefined};
}
export function createBackup(files: OfficeFile[]): Blob {
  return new Blob([JSON.stringify({format:'folio-workspace',version:1,exportedAt:new Date().toISOString(),files},null,2)],{type:'application/json'});
}
const record = (v: unknown): v is Record<string,unknown> => !!v && typeof v==='object'&&!Array.isArray(v);
function text(v:unknown,limit=100000){return typeof v==='string'?v.slice(0,limit):'';}
function pageSetup(v:unknown):DocumentPageSetup|undefined{if(!record(v))return;return {landscape:v.landscape===true,margin:['normal','narrow','wide'].includes(String(v.margin))?v.margin as DocumentPageSetup['margin']:'normal',size:['a4','letter','legal'].includes(String(v.size))?v.size as DocumentPageSetup['size']:'a4',header:text(v.header,500),footer:text(v.footer,500),pageNumbers:v.pageNumbers===true};}
function comments(v:unknown):DocumentComment[]{if(!Array.isArray(v))return [];return v.slice(0,1000).filter(record).map(c=>({id:text(c.id,100)||id(),text:text(c.text,5000),quote:text(c.quote,5000),createdAt:typeof c.createdAt==='number'?c.createdAt:Date.now(),resolved:c.resolved===true}));}
function checkContent(kind:OfficeFile['kind'],content:unknown):string|SheetContent|SlideData[]{
 if(kind==='document'){if(typeof content!=='string')throw Error('A document in this backup has invalid content.');return DOMPurify.sanitize(content,{USE_PROFILES:{html:true},FORBID_TAGS:['script','iframe','object','embed','form']});}
 if(kind==='spreadsheet'){
  if(!record(content)||!record(content.cells))throw Error('A spreadsheet in this backup has invalid content.');
  const checkCells=(cells:unknown)=>{if(!record(cells))throw Error('A worksheet in this backup has invalid cells.');for(const [address,value] of Object.entries(cells)){if(!/^[A-Z]{1,3}[1-9]\d{0,6}$/.test(address)||typeof value!=='string')throw Error('A worksheet contains an invalid cell address or value.');}};
  checkCells(content.cells);
  if(content.sheets!==undefined){if(!Array.isArray(content.sheets)||content.sheets.length>200||!content.sheets.length)throw Error('This backup has an invalid worksheet list.');for(const sheet of content.sheets){if(!record(sheet)||typeof sheet.name!=='string'||typeof sheet.id!=='string')throw Error('This backup has an invalid worksheet.');checkCells(sheet.cells);}}
  return structuredClone(content) as unknown as SheetContent;
 }
 if(!Array.isArray(content)||!content.length||content.length>2000)throw Error('A presentation in this backup has an invalid slide list.');
 for(const slide of content){if(!record(slide)||typeof slide.title!=='string'||typeof slide.body!=='string'||!['title','content','split','blank'].includes(String(slide.layout)))throw Error('This backup has an invalid slide.');if(slide.elements!==undefined){if(!Array.isArray(slide.elements)||slide.elements.length>500)throw Error('A slide has too many objects.');for(const e of slide.elements){if(!record(e)||!['text','image','shape'].includes(String(e.type))||['x','y','width','height'].some(k=>typeof e[k]!=='number'||!Number.isFinite(e[k])))throw Error('A slide contains an invalid object.');if(e.type==='image'&&typeof e.src==='string'&&!/^(data:image\/(?:png|jpeg|gif|webp);base64,|https?:\/\/)/i.test(e.src))throw Error('A slide image uses an unsupported address.');}}}
 return structuredClone(content) as SlideData[];
}
export async function readBackup(file: File): Promise<OfficeFile[]> {
 let data:unknown;try{data=JSON.parse(await file.text());}catch{throw Error('This file is not a valid Folio backup.');}
 if(!record(data)||data.format!=='folio-workspace'||data.version!==1||!Array.isArray(data.files)||data.files.length>2000)throw Error('Choose a workspace backup created by Folio.');
 const restored:OfficeFile[]=[];
 for(const item of data.files){
  if(!record(item)||!['document','spreadsheet','presentation'].includes(String(item.kind))||typeof item.name!=='string')throw Error('This backup contains an invalid file.');
  const kind=item.kind as OfficeFile['kind'];
  const versions:FileVersion[]=Array.isArray(item.versions)?item.versions.slice(0,10).filter(record).map(v=>({id:id(),name:text(v.name,100)||'Saved version',createdAt:typeof v.createdAt==='number'?v.createdAt:Date.now(),content:checkContent(kind,v.content),pageSetup:pageSetup(v.pageSetup),comments:comments(v.comments)})):[];
  restored.push({id:id(),name:item.name.slice(0,160)||'Restored file',kind,content:checkContent(kind,item.content),createdAt:typeof item.createdAt==='number'?item.createdAt:Date.now(),updatedAt:Date.now(),starred:item.starred===true,trashed:item.trashed===true,pageSetup:pageSetup(item.pageSetup),comments:comments(item.comments),versions});
 }
 return restored;
}
