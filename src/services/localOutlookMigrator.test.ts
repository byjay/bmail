import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  inlineSignatureHtmlImages,
  readSignaturesFromDisk,
} from './localOutlookMigrator';

test('inlineSignatureHtmlImages and readSignaturesFromDisk embed relative business card assets as base64 data URIs', () => {
  const sigDir = fs.mkdtempSync(path.join(os.tmpdir(), 'bmail-sig-test-'));
  try {
    const assetsDir = path.join(sigDir, 'a (designsir@seastargo.com).files');
    fs.mkdirSync(assetsDir, { recursive: true });
    // Minimal 1x1 transparent PNG bytes
    const pngBytes = Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=',
      'base64'
    );
    fs.writeFileSync(path.join(assetsDir, 'image001.png'), pngBytes);

    const sampleHtml = [
      '<html><head><meta charset="utf-8"></head><body>',
      '<div class="WordSection1">',
      '<p>Bong Jeong Kim 김봉정 부장 / SEASTAR</p>',
      '<!--[if gte vml 1]><v:shape><v:imagedata src="a%20(designsir@seastargo.com).files/image001.png"/></v:shape><![endif]-->',
      '<![if !vml]><img width="120" height="40" src="a%20(designsir@seastargo.com).files/image001.png" alt="SEASTAR"><![endif]>',
      '<p style="color:#C00000">This email and any attachments are confidential.</p>',
      '</div></body></html>',
    ].join('\n');

    fs.writeFileSync(path.join(sigDir, 'a (designsir@seastargo.com).htm'), sampleHtml, 'utf8');
    fs.writeFileSync(
      path.join(sigDir, 'a (designsir@seastargo.com).txt'),
      'Bong Jeong Kim\nSEASTAR',
      'utf8'
    );

    const inlined = inlineSignatureHtmlImages(sampleHtml, sigDir);
    assert.equal(inlined.inlinedImages, 1);
    assert.ok(inlined.html.includes('data:image/png;base64,'));
    assert.ok(!inlined.html.includes('v:imagedata'));
    assert.ok(inlined.html.includes('#C00000'));

    const loaded = readSignaturesFromDisk('', '', sigDir, 'designsir@seastargo.com');
    assert.equal(loaded.length, 1);
    assert.equal(loaded[0].name, 'a (designsir@seastargo.com)');
    assert.equal(loaded[0].isDefaultNew, true);
    assert.equal(loaded[0].isDefaultReply, true);
    assert.ok(loaded[0].htmlContent.includes('data:image/png;base64,'));
  } finally {
    fs.rmSync(sigDir, { recursive: true, force: true });
  }
});
