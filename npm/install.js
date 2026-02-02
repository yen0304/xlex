#!/usr/bin/env node
/**
 * XLEX npm installer
 * Downloads and installs the appropriate binary for the current platform
 */

const https = require('https');
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const os = require('os');
const zlib = require('zlib');

const REPO = 'user/xlex';
const VERSION = require('./package.json').version;

function getPlatform() {
  const platform = os.platform();
  const arch = os.arch();

  const platformMap = {
    darwin: 'macos',
    linux: 'linux',
    win32: 'windows',
  };

  const archMap = {
    x64: 'x86_64',
    arm64: 'aarch64',
  };

  const mappedPlatform = platformMap[platform];
  const mappedArch = archMap[arch];

  if (!mappedPlatform || !mappedArch) {
    throw new Error(`Unsupported platform: ${platform}/${arch}`);
  }

  return { platform: mappedPlatform, arch: mappedArch };
}

function getDownloadUrl() {
  const { platform, arch } = getPlatform();
  const ext = platform === 'windows' ? 'zip' : 'tar.gz';
  return `https://github.com/${REPO}/releases/download/v${VERSION}/xlex-${platform}-${arch}.${ext}`;
}

function downloadFile(url, dest) {
  return new Promise((resolve, reject) => {
    const file = fs.createWriteStream(dest);
    
    const request = (url) => {
      https.get(url, (response) => {
        if (response.statusCode === 302 || response.statusCode === 301) {
          request(response.headers.location);
          return;
        }
        
        if (response.statusCode !== 200) {
          reject(new Error(`Failed to download: ${response.statusCode}`));
          return;
        }

        response.pipe(file);
        file.on('finish', () => {
          file.close();
          resolve();
        });
      }).on('error', (err) => {
        fs.unlink(dest, () => {});
        reject(err);
      });
    };

    request(url);
  });
}

async function extractTarGz(src, dest) {
  const tar = require('tar');
  await tar.x({
    file: src,
    cwd: dest,
  });
}

async function install() {
  const binDir = path.join(__dirname, 'bin');
  const { platform } = getPlatform();
  const binaryName = platform === 'windows' ? 'xlex.exe' : 'xlex';
  const binaryPath = path.join(binDir, binaryName);

  // Create bin directory
  if (!fs.existsSync(binDir)) {
    fs.mkdirSync(binDir, { recursive: true });
  }

  // Download
  const url = getDownloadUrl();
  const tmpFile = path.join(os.tmpdir(), `xlex-${Date.now()}.tar.gz`);

  console.log(`Downloading xlex v${VERSION}...`);
  console.log(`URL: ${url}`);

  try {
    await downloadFile(url, tmpFile);
  } catch (err) {
    console.error('Failed to download xlex:', err.message);
    console.log('');
    console.log('You can manually install xlex from:');
    console.log(`  https://github.com/${REPO}/releases`);
    process.exit(1);
  }

  // Extract
  console.log('Extracting...');
  
  if (platform === 'windows') {
    execSync(`powershell -Command "Expand-Archive -Path '${tmpFile}' -DestinationPath '${binDir}' -Force"`);
  } else {
    execSync(`tar -xzf "${tmpFile}" -C "${binDir}"`);
  }

  // Clean up
  fs.unlinkSync(tmpFile);

  // Make executable (Unix)
  if (platform !== 'windows') {
    fs.chmodSync(binaryPath, 0o755);
  }

  console.log(`xlex v${VERSION} installed successfully!`);
}

install().catch((err) => {
  console.error('Installation failed:', err);
  process.exit(1);
});
