REM Build script
REM TODO: continue on error
tsc
tslint --project .
web-ext lint
web-ext build --overwrite-dest
