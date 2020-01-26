REM Build script
REM First go to https://unpkg.com/webextension-polyfill/dist/
REM and get browser-polyfill.js & browser-polyfill.js.map
REM TODO: continue on error
tsc
tslint --project .
web-ext lint
web-ext build --overwrite-dest
