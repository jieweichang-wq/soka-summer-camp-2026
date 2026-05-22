const fs = require('fs');
const path = require('path');

const srcFile = path.join(__dirname, 'index.html');
const distDir = path.join(__dirname, 'dist');
const destFile = path.join(distDir, 'index.html');

console.log('Preparing production build...');

try {
  // Create dist directory if it doesn't exist
  if (!fs.existsSync(distDir)) {
    fs.mkdirSync(distDir);
    console.log('Created dist/ directory.');
  }

  // Copy index.html to dist/index.html
  fs.copyFileSync(srcFile, destFile);
  
  console.log('Build completed successfully!');
  console.log(`Production file generated at: ${destFile}`);
} catch (error) {
  console.error('Build failed:', error.message);
  process.exit(1);
}
