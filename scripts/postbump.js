const fs = require('fs');
const path = require('path');

// When standard-version bumps the version in package.json, it sets an environment variable.
// We read that new version. It will be a 3-part semver string like "1.2.3".
const newVersion3Part = process.env.npm_package_version;

if (!newVersion3Part) {
  console.error('Error: Could not find the new version in the environment variable npm_package_version.');
  console.error('This script should be run by npm in a "postbump" lifecycle script.');
  process.exit(1);
}

// We will append ".0" to create the 4-part version string.
// This ensures that "major", "minor", and "patch" releases from conventional
// commits correctly increment the first three parts of the version.
const newVersion4Part = `${newVersion3Part}.0`;

console.log(`The new 3-part version is ${newVersion3Part}. Converting to 4-part version: ${newVersion4Part}.`);

// --- Update package.json ---
const packageJsonPath = path.resolve(__dirname, '..', 'package.json');
const pkg = JSON.parse(fs.readFileSync(packageJsonPath, 'utf-8'));

console.log(`Updating package.json version from ${pkg.version} to ${newVersion4Part}...`);
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
