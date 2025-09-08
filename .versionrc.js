module.exports = {
  // We need to tell standard-version which files to check for version updates.
  // It will use this list to decide which files to add to the release commit.
  "bumpFiles": [
    { "filename": "package.json", "type": "json" },
    { "filename": "manifest.xml", "type": "plain-text" }
  ]
}
