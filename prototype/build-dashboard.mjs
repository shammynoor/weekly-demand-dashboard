import fs from 'node:fs/promises';import path from 'node:path';
export async function buildDashboard(root,destination){
 await fs.mkdir(destination,{recursive:true});
 let app=await fs.readFile(path.join(root,'app.js'),'utf8');
 // Hoisted adapter definitions replace the legacy integration before initAuth runs.
 // Keep the proven parsing, merging, chart calculations and all control IDs intact.
 const finalStart=app.indexOf(' // Date the newest Slot Report CSV');
 const finalEnd=app.indexOf("$('signIn').onclick",finalStart);
 const originalFinalization=app.slice(finalStart,finalEnd);
 if(finalStart<0||finalEnd<0||!originalFinalization.includes('records=[...aggregate]'))throw Error('Dashboard integration markers changed. Review the adapter.');
 const finalization=originalFinalization.replace(' }catch(e){',' }catch(e){').replace(/OneDrive refresh error/g,'Practice report refresh error');
 const snapshotStart=app.indexOf(' const snapshot=JSON.stringify({');
 const snapshotEnd=app.indexOf('\n try{',snapshotStart);
 const snapshot=app.slice(snapshotStart,snapshotEnd).replace(' const snapshot=JSON.stringify({',' return JSON.stringify({').replace('folder:selectedFolder||cloudFolder||null,','');
 if(!snapshot.includes('version:9'))throw Error('Settings snapshot marker changed.');
 const adapter=await fs.readFile(path.join(root,'prototype/dashboard-adapter.js'),'utf8');
 app+='\nfunction analyticsSettingsSnapshot(){\n'+snapshot+'\n}\n'+adapter.replace('/* REPORT_FINALIZATION */',finalization);
 // All browser-local preferences are scoped to the practice AND signed-in account.
 app=app.replace(/\blocalStorage\b/g,'iqLocalStorage').replace(/\bsessionStorage\b/g,'iqSessionStorage');
 app=app.replace(/event\.key===/g,"(event.key?.startsWith(iqStorageScope)?event.key.slice(iqStorageScope.length):null)===");
 app=app.replace("const GRAPH='https://graph.microsoft.com/v1.0'","const GRAPH='disabled:'");
 app=app.replace(/OneDrive/g,'practice workspace').replace(/Microsoft sign-in/g,'practice sign-in');
 await fs.writeFile(path.join(destination,'app.js'),app);
 let html=await fs.readFile(path.join(root,'index.html'),'utf8');
 html=html.replace(/<script[^>]*src="[^"]*msal-browser[^"]*"[^>]*><\/script>/g,'');
 html=html.replace('<head>','<head><script src="scope.js"></script>');
 html=html.replace(/OneDrive/g,'practice workspace').replace(/Microsoft/g,'Darwin IQ');
 html=html.replace(/CSV reports are read from[\s\S]*?loaded on every computer\./,'Uploaded exports are reduced to reporting aggregates and stored for this practice. Patient identifiers and the original CSV are not retained. Clinician names support People and Workforce settings.');
 html=html.replace('</head>','<link rel="stylesheet" href="adapter.css"></head>');
 html=html.replace(/app\.js\?v=[^" ]+/g,'app.js?v=20261105-full-mvp');
 await fs.writeFile(path.join(destination,'index.html'),html);
 for(const file of ['forecast.js','routine-capacity.js','data-correctness.js','direct-rh.js'])await fs.copyFile(path.join(root,file),path.join(destination,file));
 await fs.copyFile(path.join(root,'prototype/dashboard-scope.js'),path.join(destination,'scope.js'));
 await fs.copyFile(path.join(root,'prototype/dashboard-adapter.css'),path.join(destination,'adapter.css'));
}
