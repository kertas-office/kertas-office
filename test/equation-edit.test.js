import { test } from 'node:test';
import assert from 'node:assert/strict';
import { latexToOmml, ommlToLatex, xmlToElements, equationText, EQUATION_PRESETS } from '../src/lib/omml.js';

const para = (tex) => xmlToElements(`<m:oMathPara>${latexToOmml(tex).omath}</m:oMathPara>`);

test('every example reads back as LaTeX that makes the same equation', () => {
  for (const p of EQUATION_PRESETS) {
    const back = ommlToLatex(para(p.tex));
    assert.equal(latexToOmml(back).omath, latexToOmml(p.tex).omath, `${p.label}: ${back}`);
  }
});

test('the LaTeX read back is the kind a person writes', () => {
  assert.equal(ommlToLatex(para('a^2+b^2=c^2')), 'a^2+b^2=c^2');
  assert.equal(ommlToLatex(para('A=\\pi r^2')), 'A=\\pi r^2');
  assert.equal(ommlToLatex(para('\\lim_{x\\to 0}\\frac{\\sin x}{x}=1')), '\\lim_{x\\to 0} \\frac{\\sin x}{x}=1');
  assert.equal(ommlToLatex(para('(x+a)^n')), '(x+a)^n');
});

test('an equation made in Word reads back too', () => {
  // as Word saves them: an integral with no chr, a function, a hidden root degree, an accent and a bar
  const word = `<m:oMathPara><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath>
    <m:nary><m:naryPr><m:limLoc m:val="subSup"/><m:ctrlPr/></m:naryPr><m:sub><m:r><m:t>0</m:t></m:r></m:sub><m:sup><m:r><m:t>∞</m:t></m:r></m:sup>
      <m:e><m:sSup><m:e><m:r><m:t>𝑒</m:t></m:r></m:e><m:sup><m:r><m:t>−</m:t></m:r><m:sSup><m:e><m:r><m:t>𝑥</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:sup></m:sSup><m:r><m:t>𝑑𝑥</m:t></m:r></m:e></m:nary>
    <m:r><m:t>=</m:t></m:r>
    <m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:r><m:t>π</m:t></m:r></m:e></m:rad>
    <m:r><m:t>+</m:t></m:r>
    <m:func><m:fName><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>cos</m:t></m:r></m:fName><m:e><m:acc><m:e><m:r><m:t>θ</m:t></m:r></m:e></m:acc></m:e></m:func>
    <m:r><m:t>+</m:t></m:r>
    <m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e><m:r><m:t>z</m:t></m:r></m:e></m:bar>
    <m:d><m:dPr><m:begChr m:val="|"/><m:endChr m:val="|"/></m:dPr><m:e><m:r><m:t>v</m:t></m:r></m:e></m:d>
  </m:oMath></m:oMathPara>`;
  const tex = ommlToLatex(xmlToElements(word));
  assert.equal(tex, '\\int_0^\\infty e^{-x^2}dx=\\sqrt{\\pi}+\\cos \\hat{\\theta}+\\overline{z}\\left|v\\right|');
  assert.doesNotThrow(() => latexToOmml(tex));
});

test('the element tree keeps prefixes, attributes and spaces in text', () => {
  const t = xmlToElements('<m:oMath><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t xml:space="preserve"> a </m:t></m:r></m:oMath>');
  assert.equal(t.name, 'm:oMath');
  assert.equal(t.elements[0].elements[0].elements[0].attributes['m:val'], 'p');
  assert.equal(equationText(t), ' a ');
});

test('Word\'s math italic letters, Greek ones included, read back as plain letters', () => {
  assert.equal(ommlToLatex(xmlToElements('<m:oMath><m:r><m:t>𝐴=𝜋𝑟</m:t></m:r><m:r><m:t>+𝛼𝛽+𝜃</m:t></m:r></m:oMath>')), 'A=\\pi r+\\alpha \\beta+\\theta');
});
