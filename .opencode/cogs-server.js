const http = require('http');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..', process.env.COGS_ROOT || 'docs');
const port = parseInt(process.env.COGS_PORT || '8080', 10);
const user = process.env.COGS_USER || '';
const pass = process.env.COGS_PASS || '';

function ok(req) {
  if (!user) return true;
  const h = req.headers.authorization || '';
  if (!h.startsWith('Basic ')) return false;
  const dec = Buffer.from(h.slice(6), 'base64').toString('utf8');
  const i = dec.indexOf(':');
  return dec.slice(0, i) === user && dec.slice(i + 1) === pass;
}

const CT = { '.html': 'text/html; charset=utf-8', '.csv': 'text/csv; charset=utf-8',
             '.json': 'application/json; charset=utf-8', '.css': 'text/css', '.js': 'application/javascript' };

http.createServer((req, res) => {
  if (!ok(req)) { res.writeHead(401, { 'WWW-Authenticate': 'Basic realm="COGS Dashboard"' }); return res.end('Authentication required'); }
  let p = decodeURIComponent((req.url || '/').split('?')[0]);
  if (p === '/' || p === '') p = '/index.html';
  const file = path.normalize(path.join(root, p));
  if (!file.startsWith(root)) { res.writeHead(403); return res.end('Forbidden'); }
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404, { 'Content-Type': 'text/plain' }); return res.end('Not found: ' + p); }
    res.writeHead(200, { 'Content-Type': CT[path.extname(file).toLowerCase()] || 'application/octet-stream', 'Cache-Control': 'no-cache' });
    res.end(data);
  });
}).listen(port, '0.0.0.0', () => {
  console.log('COGS dashboard webserver listening on 0.0.0.0:' + port + '  (root: ' + root + ')');
});
