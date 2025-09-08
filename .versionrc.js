module.exports = {
  // We need to tell standard-version which files to check for version updates.
  // It will use this list to decide which files to add to the release commit.
  "bumpFiles": [
    { "filename": "package.json", "type": "json" },
    { "filename": "manifest.xml", "type": "plain-text" }
  ],
  // The 'scripts' object allows us to hook into standard-version's process.
  "scripts": {
    // 'postbump' runs after standard-version has bumped the version in package.json,
    // but before it has created the commit. This is the perfect place for our script.
    "postbump": "node scripts/postbump.js"
  }
}
