const MANIFEST_VERSION_PARAM = "manifestVersion";
const UNKNOWN_VERSION = "unknown";

function parseManifestVersion(search) {
  const value = new URLSearchParams(search).get(MANIFEST_VERSION_PARAM);
  return value && value.trim() ? value : UNKNOWN_VERSION;
}

function getManifestVersion() {
  return parseManifestVersion(
    typeof window !== "undefined" ? window.location.search : "",
  );
}

export {getManifestVersion, parseManifestVersion};
