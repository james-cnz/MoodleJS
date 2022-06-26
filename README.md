# Moodle JS

A script for automating some Moodle tasks at Otago Polytechnic (New Zealand).

Automates adding an empty course, module, or topic, and rebuilding the module index, for a course using Otago Polytechnic's Blended Course template.

Works in Chrome and Firefox.

Only tried on Windows.


## Installing

For Chrome: Goto https://chrome.google.com/webstore/detail/moodlejs/ceabcncppbgolikembclgecoeabpciki

For Firefox: Open install file moodlejs-x.x-an+fx.xpi

It should add an icon "MJS" to the toolbar.

On Moodle server, install plugin Ad-hoc database queries
Add query "Course list" with contents as per course_list.sql

## Running

You can start the script by clicking on the icon, after you've logged in to OP Moodle.

##### Options are:

###### Create a new course:

1. Open the category that you want the new course created in.
2. Enter the course full name and short name.
3. Click "Create".

###### Create a new module in a course:

1. Open the course that you want the new module created in.
2. Enter the module full name, and a short name (to show on the tab).
3. Click "Create".

`Note: The image in the created module is linked to the blended template. This is not ideal, because if this image is removed from the blended template, or the blended template is removed (e.g. because a new version of it is created), then the image link will break. But perhaps not too much of an issue, because these images are supposed to be replaced?`

###### Create a new topic:

1. Open the module that you want the new topic created in.
2. Enter the topic name.
3. Click "Create".

`Note: If the module page differs from the template, the script may be confused, and and either put the new topic in the wrong place, or crash.`

###### Rebuild the module index page:

1. Open the module index page.
2. Click "Rebuild...".

`Note: If the module index page differs from the template, the script may be confused, and either wipe custom content on the page, or crash.`


## Known issues

* May need to turn editing on before use.
* If viewing as student, etc. return to normal role (or role with at least teacher privilages) before use.
* Take care not to type while the script is active and the content page is in focus.  (Input to page isn't blocked.)
* May be best not to have multiple tabs / windows open on the same Moodle site?
* In Firefox, sometimes displays error message "Invalid JSON string", but doesn't seem to cause any problems.
* Problems if OneTopic format has setting section 0 before tabs


## Development Notes

HTML content should only be updated based on a copy that was got immediately proir.
Otherwise image links will likely be broken.


## Development Tools Used

* VS Code
* nodejs (for installing other packages)
  * typescript
  * typescript-eslint
  * eslint
  * @typescript-eslint/eslint-plugin
  * web-ext
  * web-ext-types (possible alternatives: https://github.com/mozilla/webextension-polyfill/issues/78 )
  * webextension-polyfill
* validator.w3.org (for checking panel.htm)


## To Do

###### For Mozilla:
* Include detailed version notes.
* Include functions that can be tested at demo.moodle.net
* Include notes to reviewer, with website account (for demo.moodle.net).