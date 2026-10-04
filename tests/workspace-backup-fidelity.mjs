import { chromium, expect } from '@playwright/test';

const browser=await chromium.launch({headless:true,executablePath:'/usr/bin/chromium',args:['--no-sandbox']});
const page=await browser.newPage();
await page.route('**/__backup-fidelity',route=>route.fulfill({contentType:'text/html',body:'<!doctype html><title>Backup fidelity</title>'}));
try{
 await page.goto(`${process.env.BASE_URL||'http://localhost:5174'}/__backup-fidelity`);
 const result=await page.evaluate(async()=>{
  const {createBackup,readBackup,snapshot}=await import('/src/lib/workspaceBackup.ts');
  const comments=Array.from({length:1005},(_,index)=>({id:`comment-${index}`,text:`Comment ${index}`,quote:`Selected words ${index}`,createdAt:1735689600000+index,resolved:index%2===0}));
  comments[0]={...comments[0],id:'long-id-'.repeat(25),text:'Čuvaj cijeli komentar & <literal text>.\n'.repeat(4000)+'END OF LONG COMMENT',quote:'A lengthy selected passage 📝 '.repeat(400)+'END OF LONG SELECTION'};
  const setup={landscape:true,margin:'wide',size:'legal',header:'Imported header '.repeat(100)+'HEADER END',footer:'Imported footer '.repeat(100)+'FOOTER END',pageNumbers:true};
  const file={id:'source-file',name:'Long imported filename '.repeat(10)+'NAME END',kind:'document',content:`<p><span data-comment-id="${comments[0].id}">Anchored passage</span></p>`,createdAt:1735689600000,updatedAt:1735689600001,starred:true,trashed:false,pageSetup:setup,comments};
  file.versions=[snapshot(file,'Before edits')];
  const restored=(await readBackup(new File([createBackup([file])],'test.folio.json',{type:'application/json'})))[0];
  const checks={
   copiedId:restored.id!==file.id,
   name:restored.name===file.name,
   content:restored.content===file.content,
   commentCount:restored.comments.length,
   allComments:JSON.stringify(restored.comments)===JSON.stringify(comments),
   header:restored.pageSetup.header===setup.header,
   footer:restored.pageSetup.footer===setup.footer,
   snapshotComments:JSON.stringify(restored.versions[0].comments)===JSON.stringify(comments),
   snapshotHeader:restored.versions[0].pageSetup.header===setup.header,
   snapshotFooter:restored.versions[0].pageSetup.footer===setup.footer,
   commentAnchor:restored.content.includes(`data-comment-id="${comments[0].id}"`),
  };
  const unsafe={...file,content:'<p onclick="alert(1)">Safe body</p><script>alert(1)</script><iframe src="https://example.com"></iframe>',comments:[{id:'invalid-types',text:42,quote:{bad:true},createdAt:1735689600000,resolved:false}]};
  const safe=(await readBackup(new File([createBackup([unsafe])],'unsafe.folio.json')))[0];
  checks.sanitized=!/<script|<iframe|onclick=/i.test(safe.content)&&safe.content.includes('Safe body');
  checks.textTypes=safe.comments[0].text===''&&safe.comments[0].quote==='';
  return checks;
 });
 expect(result).toEqual({copiedId:true,name:true,content:true,commentCount:1005,allComments:true,header:true,footer:true,snapshotComments:true,snapshotHeader:true,snapshotFooter:true,commentAnchor:true,sanitized:true,textTypes:true});
 console.log('PASS: workspace backups retain 1,005 comments, long Unicode comments/selections/anchors, imported names and headers/footers in current files and saved versions; HTML sanitization and type validation remain active.');
}finally{await browser.close();}
