import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import {
  MANUAL_CHANGELOG,
  MANUAL_ICONS,
  MANUAL_MOTION_SCENES,
  MANUAL_VERSION,
  MANUAL_ZONES,
} from './manualCatalog';

test('manualCatalog: version matches package.json and all icons/scenes/zones are synchronized', () => {
  const pkgPath = path.resolve(process.cwd(), 'package.json');
  const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8')) as { version?: string };
  assert.equal(
    MANUAL_VERSION,
    pkg.version,
    'MANUAL_VERSION in src/data/manualCatalog.ts must stay synchronized with package.json version'
  );

  assert.ok(MANUAL_ZONES.length >= 8);
  assert.ok(MANUAL_ICONS.length >= 24);
  assert.ok(MANUAL_MOTION_SCENES.length >= 8);
  assert.ok(MANUAL_CHANGELOG.length >= 1);
  assert.equal(MANUAL_CHANGELOG[0].version, MANUAL_VERSION);

  const zoneIds = new Set(MANUAL_ZONES.map((z) => z.id));
  const iconIds = new Set<string>();
  const iconNumbers = new Set<number>();

  for (const icon of MANUAL_ICONS) {
    assert.equal(iconIds.has(icon.id), false, `Duplicate manual icon id: ${icon.id}`);
    assert.equal(iconNumbers.has(icon.number), false, `Duplicate manual icon number: ${icon.number}`);
    iconIds.add(icon.id);
    iconNumbers.add(icon.number);
    assert.ok(zoneIds.has(icon.zone), `Unknown zone ${icon.zone} on icon ${icon.id}`);
    assert.ok(icon.label.trim().length > 0);
    assert.ok(icon.summary.trim().length > 0);
    assert.ok(icon.detail.trim().length > 0);
    assert.ok(icon.xPercent >= 0 && icon.xPercent <= 100);
    assert.ok(icon.yPercent >= 0 && icon.yPercent <= 100);
  }

  for (const scene of MANUAL_MOTION_SCENES) {
    assert.ok(iconIds.has(scene.targetIconId), `Scene ${scene.id} references missing icon ${scene.targetIconId}`);
    assert.ok(zoneIds.has(scene.zone), `Scene ${scene.id} references missing zone ${scene.zone}`);
    assert.ok(scene.durationSec >= 3 && scene.durationSec <= 20);
    assert.ok(scene.narration.trim().length >= 20);
  }
});
