const http = require('http');
const fs = require('fs');
const path = require('path');
const server = http.createServer((req,res)=>{
  let p = decodeURIComponent(req.url.split('?')[0]);
  if(p==='/') p='/index.html';
  const filePath = path.join(__dirname, p);
  fs.readFile(filePath, (err,data)=>{
    if(err){ res.writeHead(404); res.end('not found'); return; }
    const ext = path.extname(filePath);
    const types = {'.html':'text/html','.docx':'application/vnd.openxmlformats-officedocument.wordprocessingml.document','.js':'text/javascript','.png':'image/png','.jpg':'image/jpeg','.jpeg':'image/jpeg'};
    res.writeHead(200, {'Content-Type': types[ext]||'application/octet-stream'});
    res.end(data);
  });
});
server.listen(8787, ()=>console.log('listening on 8787'));
