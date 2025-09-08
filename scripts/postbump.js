const fs = require('fs');
const path = require('path');

// --- Read package.json ---
const packageJsonPath = path.resolve(__dirname, '..', 'package.json');
const pkg = JSON.parse(fs.readFileSync(packageJsonPath, 'utf-8'));

// The version in the file is the new 3-part version from standard-version.
// This is the correct way to get the version, not from an environment variable.
const newVersion3Part = pkg.version;

// We will append ".0" to create the 4-part version string.
const newVersion4Part = `${newVersion3Part}.0`;

console.log(`Read 3-part version ${newVersion3Part} from package.json. Converting to 4-part version: ${newVersion4Part}.`);

// --- Update package.json ---
console.log(`Updating package.json version to ${newVersion4Part}...`);
pkg.version = newVersion4Part;
fs.writeFileSync(packageJsonPath, JSON.stringify(pkg, null, 2) + '\n');
console.log('Successfully updated package.json.');

// --- Update manifest.xml ---
const manifestXmlPath = path.resolve(__dirname, '..', 'manifest.xml');
let manifest = fs.readFileSync(manifestXmlPath, 'utf-8');

console.log(`Updating manifest.xml version to ${newVersion4Part}...`);
const versionRegex = /(<Version>)([\d\.]+)(<\/Version>)/;
if (!versionRegex.test(manifest)) {
    console.error('Error: Could not find <Version> tag in manifest.xml.');
    process.exit(1);
}
manifest = manifest.replace(versionRegex, `$1${newVersion4Part}$3`);
fs.writeFileSync(manifestXmlPath, manifest);
console.log('Successfully updated manifest.xml.');
