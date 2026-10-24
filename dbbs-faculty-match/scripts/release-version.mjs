import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const applicationRoot = fileURLToPath(new URL('../', import.meta.url));
const number = '(?:0|[1-9][0-9]*)';
const prerelease = '(?:0|[1-9][0-9]*|[0-9]*[a-zA-Z-][0-9a-zA-Z-]*)';
const releaseTag = new RegExp(
  `^v?(${number}\\.${number}\\.${number}(?:-${prerelease}(?:\\.${prerelease})*)?(?:\\+[0-9a-zA-Z-]+(?:\\.[0-9a-zA-Z-]+)*)?)$`,
);

export function versionFromTag(tag) {
  const match = typeof tag === 'string' && releaseTag.exec(tag);
  if (!match || match[0] !== tag) {
    throw new Error(
      'Release tag must be a SemVer version such as v1.2.3 or 1.2.3.',
    );
  }
  return match[1];
}

function replaceTomlVersion(source, section, packageName, version) {
  let replacements = 0;
  const updated = source
    .split(/(?=^\[)/m)
    .map((block) => {
      if (!block.startsWith(section)) return block;
      if (
        packageName &&
        block.match(/^name[ \t]*=[ \t]*"([^"]+)"/m)?.[1] !== packageName
      ) {
        return block;
      }
      return block.replace(/^(version[ \t]*=[ \t]*)"[^"]+"/m, (_, prefix) => {
        replacements += 1;
        return `${prefix}"${version}"`;
      });
    })
    .join('');
  if (replacements !== 1)
    throw new Error(
      `Expected one ${section} version for ${packageName || 'the app'}.`,
    );
  return updated;
}

export function applyReleaseVersion(root, tag) {
  const version = versionFromTag(tag);
  const read = (name) => fs.readFileSync(path.join(root, name), 'utf8');
  const manifest = JSON.parse(read('package.json'));
  const lock = JSON.parse(read('package-lock.json'));
  const config = JSON.parse(read('src-tauri/tauri.conf.json'));
  const cargo = read('src-tauri/Cargo.toml');
  const cargoLock = read('src-tauri/Cargo.lock');
  const packageSection = cargo
    .split(/(?=^\[)/m)
    .find((block) => block.startsWith('[package]'));
  const cargoName = packageSection?.match(/^name[ \t]*=[ \t]*"([^"]+)"/m)?.[1];
  if (!cargoName || !manifest.name || !lock.packages?.['']) {
    throw new Error(
      'The app manifests are missing their root package metadata.',
    );
  }
  manifest.version = version;
  lock.version = version;
  lock.packages[''].version = version;
  config.version = version;
  const json = (value) => `${JSON.stringify(value, null, 2)}\n`;
  // Validate every edit before writing any file. Other Cargo package versions stay intact.
  const changes = new Map([
    ['package.json', json(manifest)],
    ['package-lock.json', json(lock)],
    ['src-tauri/tauri.conf.json', json(config)],
    [
      'src-tauri/Cargo.toml',
      replaceTomlVersion(cargo, '[package]', null, version),
    ],
    [
      'src-tauri/Cargo.lock',
      replaceTomlVersion(cargoLock, '[[package]]', cargoName, version),
    ],
  ]);
  for (const [name, contents] of changes)
    fs.writeFileSync(path.join(root, name), contents);
  return { version, slug: manifest.name };
}

if (
  process.argv[1] &&
  pathToFileURL(path.resolve(process.argv[1])).href === import.meta.url
) {
  try {
    const args = process.argv.slice(2);
    const checkOnly = args[0] === '--check';
    const tag = args[checkOnly ? 1 : 0];
    if (args.length !== (checkOnly ? 2 : 1))
      throw new Error('Pass exactly one release tag.');
    const metadata = checkOnly
      ? { version: versionFromTag(tag) }
      : applyReleaseVersion(applicationRoot, tag);
    if (process.env.GITHUB_OUTPUT) {
      for (const [key, value] of Object.entries({ tag, ...metadata })) {
        fs.appendFileSync(process.env.GITHUB_OUTPUT, `${key}=${value}\n`);
      }
    }
    console.log(`Release version: ${metadata.version}`);
  } catch (error) {
    console.error(`Unable to set release version: ${error.message}`);
    process.exitCode = 1;
  }
}
