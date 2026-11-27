// Kertas Office - audio and video for the Convert app, through FFmpeg. FFmpeg is not shipped
// with Kertas: it is downloaded the first time someone converts audio or video, from one
// pinned release, and used only after its checksum matches the one recorded here.
// FFmpeg is free software under the GPL; see https://ffmpeg.org/legal.html.
const { app, net } = require('electron');
const path = require('path');
const fs = require('fs');
const fsp = require('fs/promises');
const zlib = require('zlib');
const crypto = require('crypto');
const { spawn } = require('child_process');

const RELEASE = 'b6.1.1';
const SOURCE = `https://github.com/eugeneware/ffmpeg-static/releases/download/${RELEASE}`;
// sha256 of each download (the .gz file)
const BUILDS = {
  'win32-x64': { sha256: '8883a3dffbd0a16cf4ef95206ea05283f78908dbfb118f73c83f4951dcc06d77', size: 29581307 },
  'darwin-arm64': { sha256: '8923876afa8db5585022d7860ec7e589af192f441c56793971276d450ed3bbfa', size: 19246198 },
  'darwin-x64': { sha256: '929b375c1182d956c51f7ac25e0b2b0411fb01f6f407aa15c9758efeb4242106', size: 25296431 },
  'linux-x64': { sha256: 'bfe8a8fc511530457b528c48d77b5737527b504a3797a9bc4866aeca69c2dffa', size: 29354986 },
  'linux-arm64': { sha256: '754a678672298bc68156adff58aa7385a592c2b30b1d0ae8750c45c915c4bac0', size: 25568691 },
};
// Windows on ARM runs the x64 build
const KEY = process.platform === 'win32' ? 'win32-x64' : `${process.platform}-${process.arch}`;
const EXE = process.platform === 'win32' ? 'ffmpeg.exe' : 'ffmpeg';
const dir = () => path.join(app.getPath('userData'), 'engines', `ffmpeg-${RELEASE}`);
const binary = () => process.env.KERTAS_FFMPEG || path.join(dir(), EXE);

const status = () => ({
  supported: !!BUILDS[KEY] || !!process.env.KERTAS_FFMPEG,
  installed: fs.existsSync(binary()),
  size: BUILDS[KEY]?.size || 0,
});

let installing = null;
/** Downloads, checks and unpacks FFmpeg; onProgress(0..1) */
function install(onProgress) {
  if (status().installed) return Promise.resolve(status());
  if (installing) return installing;
  const build = BUILDS[KEY];
  if (!build) return Promise.reject(new Error('Audio and video conversion is not available on this kind of computer'));
  installing = new Promise((resolve, reject) => {
    const target = dir();
    fs.mkdirSync(target, { recursive: true });
    const gz = path.join(target, `${EXE}.gz.part`);
    const out = fs.createWriteStream(gz);
    const hash = crypto.createHash('sha256');
    let got = 0;
    const req = net.request({ url: `${SOURCE}/ffmpeg-${KEY}.gz`, redirect: 'follow' });
    req.on('response', (res) => {
      if (res.statusCode !== 200) { out.destroy(); reject(new Error(`The download failed (${res.statusCode})`)); return; }
      res.on('data', (chunk) => {
        got += chunk.length;
        hash.update(chunk);
        out.write(chunk);
        onProgress?.(Math.min(1, got / build.size));
      });
      res.on('end', () => out.end(async () => {
        try {
          if (hash.digest('hex') !== build.sha256) throw new Error('The download did not match the expected file, so it was not used');
          const exe = path.join(target, EXE);
          const tmp = `${exe}.part`;
          await new Promise((ok, fail) => fs.createReadStream(gz).pipe(zlib.createGunzip()).pipe(fs.createWriteStream(tmp, { mode: 0o755 })).on('finish', ok).on('error', fail));
          await fsp.rename(tmp, exe);
          resolve(status());
        } catch (e) { reject(e); } finally { fsp.rm(gz, { force: true }).catch(() => {}); }
      }));
      res.on('error', (e) => { out.destroy(); reject(e); });
    });
    req.on('error', (e) => { out.destroy(); reject(new Error(`The download failed: ${e.message}`)); });
    req.end();
  }).finally(() => { installing = null; });
  return installing;
}

// What each output is made with. Video sizes are capped by `height` when one is chosen.
const scale = (h) => (h ? ['-vf', `scale=-2:'min(${h},ih)'`] : []);
const ARGS = {
  mp4: (o) => [...scale(o.height), '-c:v', 'libx264', '-preset', 'veryfast', '-crf', '23', '-pix_fmt', 'yuv420p', '-c:a', 'aac', '-b:a', '160k', '-movflags', '+faststart'],
  mov: (o) => [...scale(o.height), '-c:v', 'libx264', '-preset', 'veryfast', '-crf', '23', '-pix_fmt', 'yuv420p', '-c:a', 'aac', '-b:a', '160k'],
  mkv: (o) => [...scale(o.height), '-c:v', 'libx264', '-preset', 'veryfast', '-crf', '23', '-c:a', 'aac', '-b:a', '160k'],
  webm: (o) => [...scale(o.height), '-c:v', 'libvpx-vp9', '-crf', '33', '-b:v', '0', '-deadline', 'realtime', '-cpu-used', '8', '-row-mt', '1', '-c:a', 'libopus', '-b:a', '128k'],
  avi: (o) => [...scale(o.height), '-c:v', 'mpeg4', '-q:v', '4', '-c:a', 'libmp3lame', '-q:a', '4'],
  gif: (o) => ['-an', '-vf', `fps=12,scale='min(${o.height ? Math.round(o.height * 16 / 9) : 640},iw)':-2:flags=lanczos,split[a][b];[a]palettegen=stats_mode=diff[p];[b][p]paletteuse=dither=bayer`, '-loop', '0'],
  mp3: () => ['-vn', '-c:a', 'libmp3lame', '-q:a', '2'],
  m4a: () => ['-vn', '-c:a', 'aac', '-b:a', '192k'],
  wav: () => ['-vn', '-c:a', 'pcm_s16le'],
  flac: () => ['-vn', '-c:a', 'flac'],
  ogg: () => ['-vn', '-c:a', 'libvorbis', '-q:a', '5'],
  opus: () => ['-vn', '-c:a', 'libopus', '-b:a', '128k'],
};

const running = new Map();
const seconds = (h, m, s) => Number(h) * 3600 + Number(m) * 60 + Number(s);

/** input file -> output file in `to`; onProgress(0..1). Resolves with the output path */
function convert(input, output, to, opts = {}, onProgress, job) {
  if (!ARGS[to]) return Promise.reject(new Error(`Audio and video cannot be saved as ${to}`));
  if (!status().installed) return Promise.reject(new Error('Download the audio and video engine first'));
  return new Promise((resolve, reject) => {
    const tmp = `${output}.kertas-part${path.extname(output)}`;
    const child = spawn(binary(), ['-hide_banner', '-nostdin', '-y', '-i', input, ...ARGS[to](opts), tmp], { windowsHide: true });
    if (job) running.set(job, child);
    let duration = 0;
    let log = '';
    child.stderr.on('data', (d) => {
      const s = String(d);
      log = (log + s).slice(-4000);
      const dm = /Duration: (\d+):(\d+):([\d.]+)/.exec(s);
      if (dm && !duration) duration = seconds(dm[1], dm[2], dm[3]);
      const tm = /time=(\d+):(\d+):([\d.]+)/g;
      let m; let last = null;
      while ((m = tm.exec(s))) last = m;
      if (last && duration) onProgress?.(Math.min(0.99, seconds(last[1], last[2], last[3]) / duration));
    });
    child.on('error', (e) => { running.delete(job); reject(e); });
    child.on('close', async (code, signal) => {
      running.delete(job);
      if (code === 0) {
        try { await fsp.rename(tmp, output); resolve(output); } catch (e) { reject(e); }
        return;
      }
      fsp.rm(tmp, { force: true }).catch(() => {});
      if (signal || code === 255) { reject(new Error('Stopped')); return; }
      // the last meaningful line FFmpeg wrote
      const why = log.split(/\r?\n/).map((l) => l.trim()).filter((l) => l && !/^(frame|size)=/.test(l)).pop() || `code ${code}`;
      reject(new Error(/Invalid data|could not find codec|does not contain any stream/i.test(log) ? 'This file has no audio or video Kertas can read' : why));
    });
  });
}
function cancel(job) {
  const child = running.get(job);
  if (child) { try { child.kill(); } catch { /* gone */ } }
}

module.exports = { status, install, convert, cancel, RELEASE };
