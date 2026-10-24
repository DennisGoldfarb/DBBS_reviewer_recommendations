import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { test } from 'node:test';
import { applyReleaseVersion, versionFromTag } from './release-version.mjs';

test('accepts plain, v-prefixed, prerelease, and build-metadata SemVer tags', () => {
  for (const [tag, expected] of [
    ['1.2.3', '1.2.3'],
    ['v1.2.3', '1.2.3'],
    ['v1.2.3-rc.1', '1.2.3-rc.1'],
    ['1.2.3-beta.2+build.42', '1.2.3-beta.2+build.42'],
  ])
    assert.equal(versionFromTag(tag), expected);
});

test('rejects malformed tags and output/shell injection', () => {
  for (const tag of [
    undefined,
    '',
    'main',
    'v1.2',
    'v01.2.3',
    '1.2.3.4',
    '1.2.3-01',
    'v1.2.3\n',
    'v1.2.3\nversion=9.9.9',
    'v1.2.3; echo bad',
  ]) {
    assert.throws(() => versionFromTag(tag), /SemVer/);
  }
});

function fixture(t) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'dbbs-release-version-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  fs.mkdirSync(path.join(root, 'src-tauri'));
  const files = {
    'package.json': JSON.stringify({
      name: 'faculty-match',
      version: '0.1.0',
      scripts: { test: 'test' },
    }),
    'package-lock.json': JSON.stringify({
      version: '0.1.0',
      packages: {
        '': { name: 'faculty-match', version: '0.1.0' },
        'node_modules/dependency': { version: '0.1.0' },
      },
    }),
    'src-tauri/tauri.conf.json': JSON.stringify({
      version: '0.1.0',
      productName: 'Faculty Match',
    }),
    'src-tauri/Cargo.toml':
      '[package]\nname = "faculty-match"\nversion = "0.1.0"\n\n[dependencies]\nother = { version = "0.1.0" }\n',
    'src-tauri/Cargo.lock':
      'version = 4\n\n[[package]]\nname = "other"\nversion = "0.1.0"\n\n[[package]]\nname = "faculty-match"\nversion = "0.1.0"\n',
  };
  for (const [file, contents] of Object.entries(files))
    fs.writeFileSync(path.join(root, file), contents);
  return {
    root,
    files,
    read: (file) => fs.readFileSync(path.join(root, file), 'utf8'),
  };
}

test('stamps all app metadata and lockfiles while retaining dependency versions', (t) => {
  const { root, read } = fixture(t);
  assert.deepEqual(applyReleaseVersion(root, 'v1.2.3'), {
    version: '1.2.3',
    slug: 'faculty-match',
  });
  assert.equal(JSON.parse(read('package.json')).version, '1.2.3');
  assert.equal(JSON.parse(read('package-lock.json')).version, '1.2.3');
  assert.equal(
    JSON.parse(read('package-lock.json')).packages[''].version,
    '1.2.3',
  );
  assert.equal(
    JSON.parse(read('package-lock.json')).packages['node_modules/dependency']
      .version,
    '0.1.0',
  );
  assert.equal(JSON.parse(read('src-tauri/tauri.conf.json')).version, '1.2.3');
  assert.match(read('src-tauri/Cargo.toml'), /version = "1\.2\.3"/);
  assert.match(
    read('src-tauri/Cargo.toml'),
    /other = \{ version = "0\.1\.0" \}/,
  );
  assert.match(
    read('src-tauri/Cargo.lock'),
    /name = "other"\nversion = "0\.1\.0"/,
  );
  assert.match(
    read('src-tauri/Cargo.lock'),
    /name = "faculty-match"\nversion = "1\.2\.3"/,
  );
  const first = read('src-tauri/Cargo.lock');
  applyReleaseVersion(root, 'v1.2.3');
  assert.equal(read('src-tauri/Cargo.lock'), first);
});

test('invalid tags and missing app lock entries fail before changing files', (t) => {
  const { root, files, read } = fixture(t);
  assert.throws(() => applyReleaseVersion(root, 'not-a-version'), /SemVer/);
  for (const [file, contents] of Object.entries(files))
    assert.equal(read(file), contents);
  const cargoLock = files['src-tauri/Cargo.lock'].replace(
    'name = "faculty-match"',
    'name = "unexpected"',
  );
  fs.writeFileSync(path.join(root, 'src-tauri/Cargo.lock'), cargoLock);
  assert.throws(() => applyReleaseVersion(root, 'v1.2.3'), /Expected one/);
  for (const [file, contents] of Object.entries(files)) {
    assert.equal(
      read(file),
      file === 'src-tauri/Cargo.lock' ? cargoLock : contents,
    );
  }
});

test('validation CLI emits the original tag and normalized version for GitHub Actions', (t) => {
  const { root } = fixture(t);
  const output = path.join(root, 'github-output');
  const script = fileURLToPath(
    new URL('./release-version.mjs', import.meta.url),
  );
  const result = spawnSync(process.execPath, [script, '--check', 'v2.3.4'], {
    encoding: 'utf8',
    env: { ...process.env, GITHUB_OUTPUT: output },
  });
  assert.equal(result.status, 0, result.stderr);
  assert.equal(fs.readFileSync(output, 'utf8'), 'tag=v2.3.4\nversion=2.3.4\n');
});
