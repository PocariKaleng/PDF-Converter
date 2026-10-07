# PDF.js

Upstream: https://github.com/mozilla/pdf.js

Distribution: https://www.npmjs.com/package/pdfjs-dist

Version: 6.4.299. The build files pdf.mjs and pdf.worker.mjs are unmodified.
License: Apache-2.0, preserved in LICENSE alongside the files.

The renderer is served locally. It does not fetch the user's PDF or fonts from
external services. Standard-font data and optional image decoders are not
included; PDFs with unusual image codecs may require opening the downloaded PDF.
