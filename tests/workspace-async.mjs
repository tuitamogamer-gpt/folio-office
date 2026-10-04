import { readFile } from 'node:fs/promises';
import { chromium, expect } from '@playwright/test';

const browser=await chromium.launch({headless:true,executablePath:'/usr/bin/chromium',args:['--no-sandbox']});
const page=await browser.newPage({viewport:{width:1440,height:1000}});
const errors=[];page.on('pageerror',error=>errors.push(error.message));
const fixture={id:'race-document',name:'Race document',kind:'document',content:'<p>Draft one before saving.</p>',createdAt:1,updatedAt:1,starred:false,trashed:false};
await page.addInitScript(file=>{
 if(!sessionStorage.getItem('race-seeded')){localStorage.setItem('folio-files-v1',JSON.stringify([file]));sessionStorage.setItem('race-seeded','true');}
},fixture);

async function openHistory(){await page.getByRole('button',{name:'Review',exact:true}).click();await page.getByRole('button',{name:'Version history',exact:true}).click();}
async function saved(){await expect(page.locator('.doc-status')).toContainText('All changes saved',{timeout:10000});}
async function delaySaves(milliseconds){await page.evaluate(delay=>{window.__saveDelay=delay;},milliseconds);}

try{
 await page.goto(process.env.BASE_URL||'http://localhost:5174');
 await page.locator('.file-name').filter({hasText:'Race document'}).click();
 await saved();
 await page.evaluate(()=>{
  const transaction=IDBDatabase.prototype.transaction;
  const complete=Object.getOwnPropertyDescriptor(IDBTransaction.prototype,'oncomplete');
  IDBDatabase.prototype.transaction=function(...args){
   const result=transaction.apply(this,args);
   const delay=window.__saveDelay||0;
   if(args[1]==='readwrite'&&delay)Object.defineProperty(result,'oncomplete',{configurable:true,set(handler){complete.set.call(this,event=>setTimeout(()=>handler.call(result,event),delay));}});
   return result;
  };
 });
 await delaySaves(1200);
 await openHistory();
 await page.getByLabel('Save the current version').fill('Original snapshot');
 await page.getByRole('button',{name:'Save version',exact:true}).click();
 await expect(page.getByRole('button',{name:'Saving…',exact:true})).toBeDisabled();
 await expect(page.getByRole('button',{name:'Restore version',exact:true})).toBeDisabled();
 await expect(page.locator('.version-item')).toHaveCount(1);
 await page.getByRole('button',{name:'Close version history'}).click();
 await page.locator('.tiptap').fill('Draft two edited while the saved version is writing.');
 await delaySaves(0);
 await saved();
 await expect(page.locator('.tiptap')).toHaveText('Draft two edited while the saved version is writing.');
 await openHistory();
 await expect(page.locator('.version-item')).toContainText('Draft one before saving.');
 await delaySaves(1200);
 await page.getByRole('button',{name:'Restore version',exact:true}).click();
 await expect(page.getByRole('dialog',{name:'Version history'})).toHaveCount(0);
 await expect(page.locator('.tiptap')).toHaveText('Draft one before saving.');
 await page.locator('.tiptap').fill('Draft three edited immediately after restoration.');
 await delaySaves(0);
 await saved();
 await expect(page.locator('.tiptap')).toHaveText('Draft three edited immediately after restoration.');
 await openHistory();
 await expect(page.locator('.version-item')).toHaveCount(2);
 await expect(page.locator('.version-item').first()).toContainText('Draft two edited while the saved version is writing.');
 await page.getByRole('button',{name:'Close version history'}).click();
 await page.reload();
 await page.locator('.file-name').filter({hasText:'Race document'}).click();
 await expect(page.locator('.tiptap')).toHaveText('Draft three edited immediately after restoration.');
 await saved();
 console.log('PASS: delayed milestone save/restore retains intervening editor changes, preserves pre-restore draft, and disables duplicate actions.');

 await page.evaluate(()=>{
  window.__transaction=IDBDatabase.prototype.transaction;window.__setItem=Storage.prototype.setItem;
  IDBDatabase.prototype.transaction=function(...args){if(args[1]==='readwrite')throw new DOMException('Simulated storage failure','QuotaExceededError');return window.__transaction.apply(this,args);};
  Storage.prototype.setItem=function(){throw new DOMException('Simulated storage failure','QuotaExceededError');};
 });
 await openHistory();
 await page.getByLabel('Save the current version').fill('Unsaved milestone');
 await page.getByRole('button',{name:'Save version',exact:true}).click();
 await expect(page.locator('.version-item')).toHaveCount(3);
 await page.getByRole('button',{name:'Close version history'}).click();
 await expect(page.locator('.save-error-banner')).toBeVisible();
 const downloadEvent=page.waitForEvent('download');
 await page.getByRole('button',{name:'Back up workspace',exact:true}).click();
 const backup=JSON.parse(await readFile(await (await downloadEvent).path(),'utf8'));
 expect(backup.files[0].content).toContain('Draft three edited immediately after restoration.');
 expect(backup.files[0].versions[0].name).toBe('Unsaved milestone');
 await page.evaluate(()=>{IDBDatabase.prototype.transaction=window.__transaction;Storage.prototype.setItem=window.__setItem;});
 await page.locator('.tiptap').fill('Draft three with storage recovered.');
 await saved();
 console.log('PASS: a failed critical save keeps current content and the new milestone available in a workspace backup.');

 await page.evaluate(()=>{
  const original=File.prototype.text;
  File.prototype.text=async function(){if(this.name==='race.folio.json')await new Promise(resolve=>setTimeout(resolve,1000));return original.call(this);};
 });
 const restored={...fixture,id:'incoming',name:'Restored notes',content:'<p>Imported backup text</p>'};
 await page.getByLabel('Restore workspace backup').setInputFiles({name:'race.folio.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify({format:'folio-workspace',version:1,files:[restored]}))});
 await page.locator('.tiptap').fill('Draft four edited while the backup is being read.');
 await expect(page.getByRole('status')).toContainText('1 files restored as copies',{timeout:10000});
 await expect(page.locator('.tiptap')).toHaveText('Draft four edited while the backup is being read.');
 await saved();
 await page.getByTitle('Back to workspace').click();
 await expect(page.locator('.file-row')).toHaveCount(2);
 await page.reload();
 await expect(page.locator('.file-row')).toHaveCount(2);
 await page.locator('.file-name').filter({hasText:'Race document'}).click();
 await expect(page.locator('.tiptap')).toHaveText('Draft four edited while the backup is being read.');
 expect(errors).toEqual([]);
 console.log('PASS: backup restore merges with the latest live workspace after asynchronous reading, preserving concurrent edits and existing files.');
}finally{await browser.close();}
