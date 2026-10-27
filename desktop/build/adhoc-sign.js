// electron-builder afterPack hook: without an Apple Developer ID certificate the macOS app
// is left unsigned, and Apple Silicon Macs refuse to start it. An ad-hoc signature fixes that.
// codesign exists only on macOS; builds made on Linux stay unsigned.
const { execFileSync } = require('child_process');
const path = require('path');

exports.default = async (ctx) => {
  if (ctx.electronPlatformName !== 'darwin' || process.platform !== 'darwin' || process.env.CSC_LINK || process.env.CSC_NAME) return;
  const app = path.join(ctx.appOutDir, `${ctx.packager.appInfo.productFilename}.app`);
  execFileSync('codesign', ['--force', '--deep', '--sign', '-', app], { stdio: 'inherit' });
};
