#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const https = require('https');
const { execSync } = require('child_process');

const REPO = 'yen0304/xlex';
const BINARY_NAME = process.platform === 'win32' ? 'xlex.exe' : 'xlex';
const BIN_DIR = path.join(__dirname, 'bin');

// Platform mapping to release asset names
const PLATFORM_MAP = {
  'linux': 'linux',
  'darwin': 'macos',
  'win32': 'windows'
};

const ARCH_MAP = {
  'x64': 'x86_64',
  'arm64': 'aarch64'
};

function getPlatformInfo() {
  const platform = PLATFORM_MAP[process.platform];
  const arch = ARCH_MAP[process.arch];
  
  if (!platform || !arch) {
    console.error(`Unsupported platform: ${process.platform}-${process.arch}`);
    process.exit(1);
  }
  
  return { platform, arch };
}

function getPackageVersion() {
  const pkg = require('./package.json');
  return pkg.version;
}

function downloadFile(url) {
  return new Promise((resolve, reject) => {
    https.get(url, (res) => {
      if (res.statusCode === 302 || res.statusCode === 301) {
        downloadFile(res.headers.location).then(resolve).catch(reject);
        return;
      }
      
      if (res.statusCode !== 200) {
        reject(new Error(`Failed to download: ${res.statusCode}`));
        return;
      }
      
      const chunks = [];
      res.on('data', chunk => chunks.push(chunk));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    }).on('error', reject);
  });
}

async function getLatestRelease() {
  return new Promise((resolve, reject) => {
    const options = {
      hostname: 'api.github.com',
      path: `/repos/${REPO}/releases/latest`,
      headers: { 'User-Agent': 'xlex-npm-installer' }
    };
    
    https.get(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          resolve(JSON.parse(data));
        } catch (e) {
          reject(e);
        }
      });
    }).on('error', reject);
  });
}

async function install() {
  const { platform, arch } = getPlatformInfo();
  const version = getPackageVersion();
  
  console.log(`Installing xlex for ${platform}-${arch}...`);
  
  // Construct download URL
  const ext = process.platform === 'win32' ? 'zip' : 'tar.gz';
  const filename = `xlex-${platform}-${arch}.${ext}`;
  
  let downloadUrl;
  
  // Try to get the exact version first
  try {
    downloadUrl = `https://github.com/${REPO}/releases/download/v${version}/${filename}`;
    console.log(`Trying version v${version}...`);
  } catch (e) {
    // Fallback to latest release
    console.log('Fetching latest release...');
    const release = await getLatestRelease();
    const asset = release.assets.find(a => a.name === filename);
    if (!asset) {
      throw new Error(`Binary not found for ${platform}-${arch}`);
    }
    downloadUrl = asset.browser_download_url;
  }
  
  console.log(`Downloading from ${downloadUrl}...`);
  
  // Download
  const buffer = await downloadFile(downloadUrl);
  
  // Create bin directory
  fs.mkdirSync(BIN_DIR, { recursive: true });
  
  // Extract
  const tmpFile = path.join(__dirname, `tmp-archive.${ext}`);
  fs.writeFileSync(tmpFile, buffer);
  
  try {
    if (process.platform === 'win32') {
      execSync(`powershell -command "Expand-Archive -Path '${tmpFile}' -DestinationPath '${BIN_DIR}' -Force"`, { stdio: 'ignore' });
    } else {
      execSync(`tar -xzf "${tmpFile}" -C "${BIN_DIR}"`, { stdio: 'ignore' });
    }
  } finally {
    fs.unlinkSync(tmpFile);
  }
  
  // Make executable on Unix
  if (process.platform !== 'win32') {
    const binaryPath = path.join(BIN_DIR, BINARY_NAME);
    fs.chmodSync(binaryPath, 0o755);
  }
  
  console.log('xlex installed successfully!');
}

install().catch((err) => {
  console.error('Installation failed:', err.message);
  console.error('');
  console.error('You can manually download from:');
  console.error(`https://github.com/${REPO}/releases`);
  process.exit(1);
});
