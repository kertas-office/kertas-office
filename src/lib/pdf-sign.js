// Certificate-based digital signatures (PAdES-style detached PKCS#7, SHA-256).
// The signer's certificate and private key come from a .p12 / .pfx file.
// Signing is always the last step: any change to the bytes afterwards breaks it.
import forge from 'node-forge';
import { PDFDocument, PDFName, PDFNumber, PDFHexString, PDFString, StandardFonts, rgb } from '@cantoo/pdf-lib';

const SIG_BYTES = 16384; // room for the signature, certificates included
const PLACEHOLDER = 9999999999;

const toBinary = (bytes) => {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return s;
};
const ascii = (s) => Uint8Array.from(s, (c) => c.charCodeAt(0));
function indexOf(hay, needle, from = 0) {
  outer: for (let i = from; i <= hay.length - needle.length; i++) {
    for (let j = 0; j < needle.length; j++) if (hay[i + j] !== needle[j]) continue outer;
    return i;
  }
  return -1;
}
function lastIndexOf(hay, needle) {
  outer: for (let i = hay.length - needle.length; i >= 0; i--) {
    for (let j = 0; j < needle.length; j++) if (hay[i + j] !== needle[j]) continue outer;
    return i;
  }
  return -1;
}

/** Opens a .p12/.pfx: { key, cert, chain, name, organisation, issuer, validFrom, validTo } */
export function readCertificate(p12Bytes, password = '') {
  let p12;
  try {
    const asn1 = forge.asn1.fromDer(toBinary(new Uint8Array(p12Bytes)));
    p12 = forge.pkcs12.pkcs12FromAsn1(asn1, false, password);
  } catch (e) {
    const msg = String(e?.message || e);
    throw new Error(/mac|password|Invalid/i.test(msg) ? 'The password for this certificate is not right.' : `This certificate file could not be read (${msg}).`);
  }
  const keyBags = [
    ...(p12.getBags({ bagType: forge.pki.oids.pkcs8ShroudedKeyBag })[forge.pki.oids.pkcs8ShroudedKeyBag] || []),
    ...(p12.getBags({ bagType: forge.pki.oids.keyBag })[forge.pki.oids.keyBag] || []),
  ];
  const certs = (p12.getBags({ bagType: forge.pki.oids.certBag })[forge.pki.oids.certBag] || []).map((b) => b.cert).filter(Boolean);
  const key = keyBags[0]?.key;
  if (!key) throw new Error('This file has no private key, so it cannot be used to sign.');
  if (!certs.length) throw new Error('This file has no certificate.');
  // The signer's certificate is the one that matches the private key
  const cert = certs.find((c) => c.publicKey?.n && key.n && c.publicKey.n.equals(key.n)) || certs[0];
  const attr = (c, n) => c.subject.getField(n)?.value || '';
  return {
    key,
    cert,
    chain: certs.filter((c) => c !== cert),
    name: attr(cert, 'CN') || attr(cert, 'O') || 'Unknown signer',
    organisation: attr(cert, 'O'),
    email: attr(cert, 'E') || cert.subject.getField({ name: 'emailAddress' })?.value || '',
    issuer: cert.issuer.getField('CN')?.value || cert.issuer.getField('O')?.value || '',
    validFrom: cert.validity.notBefore,
    validTo: cert.validity.notAfter,
    selfSigned: cert.isIssuer(cert),
  };
}

/**
 * Signs a PDF. opts: { certificate (from readCertificate), reason, location, contact,
 *   visible: { page: index, x, y, width, height } | null, date }
 * Returns the signed bytes.
 */
export async function signPdf(pdfBytes, opts) {
  const { certificate: c, reason = '', location = '', contact = '', visible = null } = opts;
  const date = opts.date || new Date();
  if (date < c.validFrom || date > c.validTo) throw new Error(`This certificate is only valid from ${c.validFrom.toLocaleDateString()} to ${c.validTo.toLocaleDateString()}.`);

  const doc = await PDFDocument.load(pdfBytes, { updateMetadata: false });
  if (doc.isEncrypted) throw new Error('Remove the password before signing. A password can be added by your PDF reader after signing.');
  const ctx = doc.context;
  const pages = doc.getPages();
  const page = pages[Math.max(0, Math.min(pages.length - 1, visible?.page ?? 0))];

  // A visible signature is drawn into the page itself, so it is covered by the signature
  let rect = [0, 0, 0, 0];
  if (visible) {
    const { x, y, width: w, height: h } = visible;
    rect = [x, y, x + w, y + h];
    const font = await doc.embedFont(StandardFonts.Helvetica);
    const bold = await doc.embedFont(StandardFonts.HelveticaBold);
    const accent = rgb(0.31, 0.27, 0.9);
    page.drawRectangle({ x, y, width: w, height: h, color: rgb(0.97, 0.97, 1), borderColor: accent, borderWidth: 1, opacity: 0.95 });
    page.drawRectangle({ x, y, width: 3, height: h, color: accent });
    const lines = [
      [bold, `Digitally signed by ${c.name}`],
      ...(c.organisation && c.organisation !== c.name ? [[font, c.organisation]] : []),
      [font, `Date: ${date.toISOString().slice(0, 19).replace('T', ' ')} UTC`],
      ...(reason ? [[font, `Reason: ${reason}`]] : []),
      ...(location ? [[font, `Location: ${location}`]] : []),
    ];
    let size = Math.min(10, (h - 8) / (lines.length * 1.25));
    const widest = Math.max(...lines.map(([f, t]) => f.widthOfTextAtSize(t, 1)));
    size = Math.max(4, Math.min(size, (w - 14) / widest));
    let ty = y + h - 5 - size;
    for (const [f, t] of lines) { page.drawText(t, { x: x + 9, y: ty, size, font: f, color: rgb(0.1, 0.1, 0.2) }); ty -= size * 1.25; }
  }

  const sig = ctx.obj({
    Type: 'Sig',
    Filter: 'Adobe.PPKLite',
    SubFilter: 'adbe.pkcs7.detached',
    ByteRange: [0, PLACEHOLDER, PLACEHOLDER, PLACEHOLDER],
    Contents: PDFHexString.of('0'.repeat(SIG_BYTES * 2)),
    M: PDFString.fromDate(date),
    Name: PDFString.of(c.name),
    ...(reason ? { Reason: PDFString.of(reason) } : {}),
    ...(location ? { Location: PDFString.of(location) } : {}),
    ...(contact || c.email ? { ContactInfo: PDFString.of(contact || c.email) } : {}),
    Prop_Build: { App: { Name: 'Kertas_Office' } },
  });
  const sigRef = ctx.register(sig);

  const acro = doc.catalog.getOrCreateAcroForm();
  const existing = acro.getAllFields().length;
  const ap = ctx.formXObject([], { BBox: [0, 0, rect[2] - rect[0], rect[3] - rect[1]] });
  const widget = ctx.obj({
    Type: 'Annot',
    Subtype: 'Widget',
    FT: 'Sig',
    Rect: rect,
    V: sigRef,
    T: PDFString.of(`Signature${existing + 1}`),
    F: 132, // print + locked
    P: page.ref,
    AP: { N: ctx.register(ap) },
  });
  const widgetRef = ctx.register(widget);
  page.node.addAnnot(widgetRef);
  acro.addField(widgetRef);
  acro.dict.set(PDFName.of('SigFlags'), PDFNumber.of(3));

  const bytes = await doc.save({ useObjectStreams: false, updateFieldAppearances: false });

  // Fill in the byte range around the signature's hole
  const hole = ascii(`<${'0'.repeat(SIG_BYTES * 2)}>`);
  const start = lastIndexOf(bytes, hole);
  if (start < 0) throw new Error('Could not prepare the signature space.');
  const end = start + hole.length;
  const brText = `[ 0 ${PLACEHOLDER} ${PLACEHOLDER} ${PLACEHOLDER} ]`;
  const brAt = indexOf(bytes, ascii(brText));
  if (brAt < 0) throw new Error('Could not prepare the signature byte range.');
  const range = [0, start, end, bytes.length - end];
  const brNew = `[0 ${range[1]} ${range[2]} ${range[3]}]`.padEnd(brText.length, ' ');
  bytes.set(ascii(brNew), brAt);

  // PKCS#7 detached signature over everything except the hole
  const signed = new Uint8Array(range[1] + range[3]);
  signed.set(bytes.subarray(0, start), 0);
  signed.set(bytes.subarray(end), start);
  const p7 = forge.pkcs7.createSignedData();
  p7.content = forge.util.createBuffer(toBinary(signed));
  p7.addCertificate(c.cert);
  for (const extra of c.chain) p7.addCertificate(extra);
  p7.addSigner({
    key: c.key,
    certificate: c.cert,
    digestAlgorithm: forge.pki.oids.sha256,
    authenticatedAttributes: [
      { type: forge.pki.oids.contentType, value: forge.pki.oids.data },
      { type: forge.pki.oids.messageDigest },
      { type: forge.pki.oids.signingTime, value: date },
    ],
  });
  p7.sign({ detached: true });
  const der = forge.asn1.toDer(p7.toAsn1()).getBytes();
  if (der.length > SIG_BYTES) throw new Error('The certificate chain is too large to fit in the signature.');
  const hex = forge.util.bytesToHex(der).toUpperCase().padEnd(SIG_BYTES * 2, '0');
  bytes.set(ascii(hex), start + 1);
  return bytes;
}
