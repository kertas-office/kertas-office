// What the Convert app offers for each kind of file
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { groupOf, outputsFor, defaultOutput, canonical, needsLibreOffice } from '../src/lib/convert/formats.js';

test('each file goes to its group by extension, whatever the case', () => {
  assert.equal(groupOf('Photo.JPEG'), 'image');
  assert.equal(groupOf('scan.heic'), 'image');
  assert.equal(groupOf('report.pdf'), 'pdf');
  assert.equal(groupOf('letter.doc'), 'doc');
  assert.equal(groupOf('notes.markdown'), 'doc');
  assert.equal(groupOf('book.xlsb'), 'sheet');
  assert.equal(groupOf('clip.MOV'), 'video');
  assert.equal(groupOf('song.flac'), 'audio');
  assert.equal(groupOf('archive.zip'), null);
});

test('a file is never offered its own format, under any of its names', () => {
  assert.ok(!outputsFor('a.jpeg').includes('jpg'));
  assert.ok(!outputsFor('a.tif').includes('tiff'));
  assert.ok(!outputsFor('a.htm').includes('html'));
  assert.ok(!outputsFor('a.mpg').includes('mpeg'));
  assert.equal(canonical('tiff'), 'tiff');
  assert.equal(canonical('html'), 'html');
});

test('.odt and .rtf are offered only for word-processor files, and need LibreOffice', () => {
  assert.ok(outputsFor('a.docx').includes('odt'));
  assert.ok(outputsFor('a.doc').includes('rtf'));
  for (const f of ['a.md', 'a.html', 'a.txt']) assert.ok(!outputsFor(f).some(needsLibreOffice), f);
});

test('each file starts with the conversion people most often want', () => {
  assert.equal(defaultOutput('a.pdf'), 'docx');
  assert.equal(defaultOutput('a.docx'), 'pdf');
  assert.equal(defaultOutput('a.xlsx'), 'pdf');
  assert.equal(defaultOutput('a.csv'), 'xlsx');
  assert.equal(defaultOutput('a.heic'), 'jpg');
  assert.equal(defaultOutput('a.png'), 'jpg');
  assert.equal(defaultOutput('a.jpg'), 'png');
  assert.equal(defaultOutput('a.mp4'), 'mp3');
  assert.equal(defaultOutput('a.mov'), 'mp4');
  assert.equal(defaultOutput('a.wav'), 'mp3');
});
