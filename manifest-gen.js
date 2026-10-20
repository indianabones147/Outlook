/* Builds a manifest from manifest.template.xml for a given https host URL. Used by setup.html and scripts/set-host.js. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.ManifestGen = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";
  var PLACEHOLDER = "https://YOUR-HOST/";

  function normalizeHost(url) {
    var u = String(url || "").trim();
    if (!/^https:\/\//i.test(u)) throw new Error("The host URL must start with https://");
    u = u.replace(/[?#].*$/, "").replace(/\/[^/]*\.html?$/i, "/");
    if (!/\/$/.test(u)) u += "/";
    return u;
  }

  function originOf(url) {
    var m = /^(https:\/\/[^/]+)/i.exec(url);
    return m ? m[1] : url;
  }

  function buildManifest(template, hostUrl) {
    var base = normalizeHost(hostUrl);
    return String(template)
      .split("<AppDomain>https://YOUR-HOST</AppDomain>").join("<AppDomain>" + originOf(base) + "</AppDomain>")
      .split(PLACEHOLDER).join(base);
  }

  return { PLACEHOLDER: PLACEHOLDER, normalizeHost: normalizeHost, originOf: originOf, buildManifest: buildManifest };
});
