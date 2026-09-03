const VERSION_PATTERN = /<Version>([^<]+)<\/Version>/;
const UNKNOWN_VERSION = "unknown";

function readManifestVersion(xml) {
  const match = xml.match(VERSION_PATTERN);
  return match ? match[1] : UNKNOWN_VERSION;
}

function stampUrl(xml, urlPattern, version) {
  return xml.replace(urlPattern, (fullMatch, prefix, query, quote) => {
    const separator = query ? "&" : "?";
    return `${prefix}${query || ""}${separator}manifestVersion=${version}${quote}`;
  });
}

function stampManifestVersion(xml) {
  const version = readManifestVersion(xml);
  let stamped = xml;

  stamped = stampUrl(
    stamped,
    /(<SourceLocation DefaultValue="[^\"]*taskpane\.html)(\?[^\"]*)?(\")/,
    version,
  );
  stamped = stampUrl(
    stamped,
    /(<bt:Url id="[^"]+" DefaultValue="[^"]*taskpane\.html)(\?[^\"]*)?(\")/g,
    version,
  );

  return stamped;
}

module.exports = {stampManifestVersion};
