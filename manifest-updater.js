const regex = /<Version>(\d+\.\d+\.\d+(\.\d+)?)<\/Version>/;

module.exports.readVersion = function (contents) {
  const match = contents.match(regex);
  if (match === null) {
    throw new Error('Could not find version in manifest.xml');
  }
  return match[1];
};

module.exports.writeVersion = function (contents, version) {
  return contents.replace(regex, `<Version>${version}</Version>`);
};
