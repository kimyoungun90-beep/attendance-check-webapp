'use strict';
const http=require('http'),fs=require('fs'),path=require('path');
const root=__dirname,port=8766;
const types={'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.png':'image/png','.svg':'image/svg+xml','.json':'application/json; charset=utf-8'};
const server=http.createServer((req,res)=>{let file;try{file=path.resolve(root,'.'+decodeURIComponent(req.url.split('?')[0]));}catch{res.writeHead(400);return res.end();}if(file!==root&&!file.startsWith(root+path.sep)){res.writeHead(403);return res.end();}if(file===root)file=path.join(root,'index.html');fs.readFile(file,(error,data)=>{if(error){res.writeHead(404);return res.end('Not found');}res.setHeader('Content-Type',types[path.extname(file)]||'application/octet-stream');res.setHeader('X-Content-Type-Options','nosniff');res.setHeader('Cache-Control','no-cache');res.end(data);});});
server.on('error',e=>{console.error(e.code==='EADDRINUSE'?'이미 실행 중입니다. http://127.0.0.1:8766 을 여세요.':e.message);process.exitCode=1;});
server.listen(port,'127.0.0.1',()=>console.log('\nZENIEL 근태 시스템\nhttp://127.0.0.1:'+port+'\n이 창을 닫으면 웹앱이 종료됩니다.\n'));
