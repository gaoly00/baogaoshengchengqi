import { defineConfig } from 'vite';
import fs from 'fs';
import path from 'path';

export default defineConfig({
  base: './', // 使用相对路径，适用于 Electron 本地文件
  server: {
    port: 5000,
    host: '0.0.0.0',
    allowedHosts: true,
  },
  build: {
    assetsDir: 'assets',
    outDir: 'dist',
  },
  plugins: [
    {
      name: 'download-directory',
      configureServer(server) {
        server.middlewares.use('/download', (req, res) => {
          const url = req.url === '/' ? '' : req.url;
          const downloadDir = path.join(process.cwd(), 'public', 'download');
          const requestedPath = path.join(downloadDir, url.replace(/^\//, ''));
          
          if (!fs.existsSync(requestedPath)) {
            res.statusCode = 404;
            res.end('Not Found');
            return;
          }

          if (fs.statSync(requestedPath).isDirectory()) {
            const files = fs.readdirSync(requestedPath);
            const fileLinks = files.map(file => {
              const filePath = path.join(requestedPath, file);
              const stat = fs.statSync(filePath);
              const size = stat.isFile() ? formatSize(stat.size) : '-';
              const isDir = stat.isDirectory();
              return `<li><a href="/download/${url.replace(/^\/|\/$/g, '')}${url.endsWith('/') ? '' : '/'}${file}${isDir ? '/' : ''}">${file}</a> ${isDir ? '(dir)' : size}</li>`;
            }).join('');
            
            res.setHeader('Content-Type', 'text/html; charset=utf-8');
            res.end(`<!DOCTYPE html>
<html>
<head><title>Download Directory</title></head>
<body>
<h1>Download Directory</h1>
<ul>${fileLinks}</ul>
</body>
</html>`);
            return;
          }

          // Serve file for download
          if (fs.statSync(requestedPath).isFile()) {
            const stat = fs.statSync(requestedPath);
            res.setHeader('Content-Type', 'application/octet-stream');
            res.setHeader('Content-Length', stat.size);
            res.setHeader('Content-Disposition', `attachment; filename="${path.basename(requestedPath)}"`);
            fs.createReadStream(requestedPath).pipe(res);
            return;
          }
          
          res.statusCode = 404;
          res.end('Not Found');
        });
      }
    }
  ],
});

function formatSize(bytes: number): string {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}
