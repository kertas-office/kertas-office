import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile, writeFile, mkdtemp } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { readCertificate, signPdf } from '../src/lib/pdf-sign.js';

const fixture = (n) => readFile(new URL(`./fixtures/${n}`, import.meta.url));
const hasPdfsig = (() => { try { execFileSync('pdfsig', ['-v'], { stdio: 'ignore' }); return true; } catch { return false; } })();

test('a .p12 certificate is read, and a wrong password is explained', async () => {
  const p12 = await fixture('signer.p12');
  const c = readCertificate(p12, 'kertas');
  assert.equal(c.name, 'Kertas Test Signer');
  assert.equal(c.organisation, 'Kertas Office');
  assert.ok(c.selfSigned);
  assert.throws(() => readCertificate(p12, 'nope'), /password/);
});

test('a signed PDF verifies, and any later change breaks it', { skip: !hasPdfsig && 'pdfsig not installed' }, async () => {
  const certificate = readCertificate(await fixture('signer.p12'), 'kertas');
  const out = await signPdf(await fixture('report.pdf'), { certificate, reason: 'Approved', location: 'Kuala Lumpur', visible: { page: 0, x: 340, y: 60, width: 200, height: 54 } });
  const dir = await mkdtemp(join(tmpdir(), 'kertas-sign-'));
  await writeFile(join(dir, 'signed.pdf'), out);
  let report = '';
  try { report = execFileSync('pdfsig', [join(dir, 'signed.pdf')], { encoding: 'utf8' }); } catch (e) { report = e.stdout; }
  assert.match(report, /Signer Certificate Common Name: Kertas Test Signer/);
  assert.match(report, /Signature Validation: Signature is Valid/);
  assert.match(report, /Signed Ranges: \[0 - \d+\], \[\d+ - \d+\]/);
  assert.match(report, /Total document signed/);
  assert.match(report, /Signature Type: adbe\.pkcs7\.detached/);

  const tampered = new Uint8Array(out);
  const at = Buffer.from(tampered).indexOf('Annual');
  if (at > 0) tampered[at] = 0x61; else tampered[100] ^= 1;
  await writeFile(join(dir, 'tampered.pdf'), tampered);
  try { report = execFileSync('pdfsig', [join(dir, 'tampered.pdf')], { encoding: 'utf8' }); } catch (e) { report = e.stdout; }
  assert.doesNotMatch(report, /Signature Validation: Signature is Valid/);
});
