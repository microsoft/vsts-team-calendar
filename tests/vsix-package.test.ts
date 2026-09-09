import * as fs from 'fs';
import * as path from 'path';
import AdmZip = require('adm-zip');

const packageDirectory = path.resolve(__dirname, '..');
const vsixFiles = fs
  .readdirSync(packageDirectory)
  .filter((fileName) => fileName.endsWith('.vsix'))
  .sort();

describe('VSIX Package Verification', () => {
  test('includes both release and dev packages', () => {
    expect(vsixFiles).toEqual(
      expect.arrayContaining([
        expect.stringMatching(/\.vsix$/),
        expect.stringMatching(/-dev-.*\.vsix$/),
      ])
    );
  });
});

describe.each(vsixFiles)('%s', (vsixFile) => {
  const vsixPath = path.join(packageDirectory, vsixFile);
  let zip: AdmZip;
  let entryNames: Set<string>;

  beforeAll(() => {
    // Ensure file exists before testing
    expect(fs.existsSync(vsixPath)).toBe(true);
    zip = new AdmZip(vsixPath);
    entryNames = new Set(
      zip.getEntries().map((entry) => entry.entryName.replace(/\\/g, '/'))
    );
  });

  test('should contain core Azure DevOps extension files', () => {
    const expectedFiles = [
      'extension.vsomanifest',
      'dist/azure-devops-extension.json',
      'dist/calendar.html',
      'dist/Calendar.js',
    ];

    for (const expectedFile of expectedFiles) {
      expect(entryNames.has(expectedFile)).toBe(true);
    }
  });

  test('should contain a valid Azure DevOps manifest', () => {
    const manifestEntry = zip.getEntry('extension.vsomanifest');
    expect(manifestEntry).toBeDefined();

    const manifest = JSON.parse(manifestEntry!.getData().toString('utf-8'));
    expect(manifest.manifestVersion).toBe(1);
    expect(manifest.contributions).toEqual(expect.any(Array));
  });

  test('should contain the expected package configuration', () => {
    const configEntry = zip.getEntry('dist/azure-devops-extension.json');
    expect(configEntry).toBeDefined();

    const config = JSON.parse(configEntry!.getData().toString('utf-8'));
    const manifestEntry = zip.getEntry('extension.vsomanifest');
    expect(manifestEntry).toBeDefined();

    const manifest = JSON.parse(manifestEntry!.getData().toString('utf-8'));
    const isDevPackage = /-dev-.*\.vsix$/.test(vsixFile);

    if (isDevPackage) {
      expect(manifest.baseUri).toBe('https://localhost:8888');
    } else {
      expect(manifest.baseUri).toBeUndefined();
    }
    expect(config.public).toBe(true);
  });
});