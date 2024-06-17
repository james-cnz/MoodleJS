REM Build script
REM First go to https://unpkg.com/webextension-polyfill/dist/
REM and get browser-polyfill.js & browser-polyfill.js.map
REM TODO: continue on error
npx eslint . --ext .ts
tsc
web-ext lint
web-ext build --overwrite-dest --ignore-files=node_modules web-ext-artifacts package-lock.json
