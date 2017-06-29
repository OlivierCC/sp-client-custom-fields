![release](https://img.shields.io/badge/release-v1.3.7-blue.svg)
![npm](https://img.shields.io/badge/npm-sp--client--custom--fields-red.svg)
![status](https://img.shields.io/badge/status-stable-green.svg)
![mit](https://img.shields.io/badge/license-MIT-yellow.svg)

This library is a kit of 30+ components to customize SPFx web parts custom fields, to make the optimal experience to edit your Web Parts.


Official Web Site: [https://oliviercc.github.io/sp-client-custom-fields](https://oliviercc.github.io/sp-client-custom-fields)

Documentation: [https://oliviercc.github.io/sp-client-custom-fields](https://oliviercc.github.io/sp-client-custom-fields)

API: [https://oliviercc.github.io/sp-client-custom-fields/docs](https://oliviercc.github.io/sp-client-custom-fields/docs)


# How to install & use

To install this library is your project, open a command line and execute this command in your WebPart's folder:
```bash
npm i --save sp-client-custom-fields
```
Open your file `'config/config.json'`, and add the following lines in the **externals** and in the **localizedResources** sections:
```
"externals": {
   "sp-client-custom-fields": "node_modules/sp-client-custom-fields/dist/sp-client-custom-fields.bundle.js"
}
"localizedResources": {
    "sp-client-custom-fields/strings": "../node_modules/sp-client-custom-fields/lib/loc/{locale}.js"
}
```
Execute gulp in the command line
```bash
gulp
```

You are now ready to use a custom property field in your web part! It's really easy to add a custom property field in your project, you can read any property documentation to view how to do that.  [More information](https://oliviercc.github.io/sp-client-custom-fields)

# Compilation Process

```bash
git clone the repo
npm i
tsd install
gulp serve
```

# Updates Log

Date | Description |  Contributors
------------ | ----------- | -----------
06/29/2017  | 1.3.7 Fix bugs with PropertyFieldRichTextBox | @OlivierC - Thanks to @kmartindale for bugs reports
06/27/2017  | 1.3.6 Fix bugs with PropertyFieldRichTextBox | @OlivierC
06/12/2017  | 1.3.5 Upgrade to latest SPFx version. Fix refreshing web part issues + fix bugs with latest version of Office UI Fabric | @ytasyol @OlivierC
05/03/2017  | 1.3.4 Fix bugs with PropertyFieldSPListQuery | @TravisGilbert
04/12/2017  | 1.3.3 Upgrade to Office UI Fabric 2.6.3 + bugs fix + new TermSetPicker custom field added | @OlivierC
04/09/2017  | 1.3.2 Custom List layout improvements + bugs fix | @OlivierC
04/05/2017  | 1.3.1 AutoComplete fixs + new OfficeVideoPicker custom field added | @OlivierC
04/03/2017  | 1.3.0 SearchPropertiesPicker custom field added | @OlivierC
04/03/2017  | 1.2.9 AutoComplete custom field added | @OlivierC
03/29/2017  | 1.2.8 release - CheckBoxes & RadioButtons styles changed | @OlivierC
03/25/2017  | 1.2.7 release - NumericInput custom field added | @OlivierC
03/25/2017  | 1.2.6 release - Bundle package optimization added | @OlivierC
03/23/2017  | 1.2.5 release - Group Picker custom field added | @OlivierC
03/21/2017  | 1.2.4 release - Web Site Documentation update | @OlivierC
03/19/2017  | 1.2.3 release - Mini Color Picker custom field added | @OlivierC
03/19/2017  | 1.2.2 release - StarRating custom field added | @OlivierC
03/18/2017  | 1.2.1 release - TagPicker custom field added | @OlivierC
03/16/2017  | 1.2.0 release - DropDownTreeView custom field added | @OlivierC
03/15/2017  | 1.1.9 release - TreeView custom field added | @OlivierC
03/13/2017  | 1.1.8 release - SortableList custom field added | @OlivierC
03/04/2017  | 1.1.7 release - bugs fix | @OlivierC
03/03/2017  | 1.1.6 release - bugs fix | @OlivierC
02/26/2017  | 1.1.5 release - upgrade to SPFx GA (v1.0.0.0) | @OlivierC
02/21/2017  | 1.1.4 release - RichTextBox update with ckeditor 4.2 upgrade | @OlivierC
02/13/2017  | 1.1.3 release - DateTimePicker improvements to support seconds + 12-hours time convention | @OlivierC
02/05/2017  | 1.1.2 release - bugs fix | @OlivierC
02/05/2017  | 1.1.1 release - Dimension Picker added + new languages (Thai, Hungarian, Greek, Turkish) | @OlivierC
02/04/2017  | 1.1.0 release - Document + Image pickers improvements (manual edition + management of file extensions) + bugs fix with SPListQuery picker | @OlivierC
02/02/2017  | 1.0.9 release - onGetErrorMessage + deferredValidationTime properties added | @OlivierC
01/28/2017  | 1.0.8 release - disabled property added + bugs fix | @OlivierC
01/23/2017  | 1.0.7 release - Fix PropertyFieldRichTextBox errors | @OlivierC
01/21/2017  | 1.0.6 release - SharePoint Framework RC0 support | @OlivierC & Daniel Zeller (@Crash753)
12/06/2016  | 1.0.5 release - SharePoint Framework Drop 6 support + improve accessibility | @OlivierC
11/13/2016  | 1.0.4 release - SharePoint Framework Drop 5 support + upgrade to Office UI Fabric 0.52 | @OlivierC
10/07/2016  | 1.0.3 release - Slider Range custom field added | @OlivierC
10/07/2016  | 1.0.2 release - Bug fix | @OlivierC
10/07/2016  | 1.0.1 release - 16 localization files included | @OlivierC
10/06/2016  | 1.0.0 release - RichTextBox custom field added | @OlivierC
10/04/2016  | 0.9.9 release - Drop down select custom field added | @OlivierC
10/03/2016  | 0.9.8 npm package release - Align Picker custom field added | @OlivierC
10/02/2016  | 0.9.7 npm package release | @OlivierC
09/29/2016  | List Query custom field added | @OlivierC
09/27/2016  | Update to Drop 4 | @OlivierC
09/22/2016  | Custom List custom field added | @OlivierC
09/19/2016  | Date Time custom field added | @OlivierC
09/18/2016  | Display Mode custom field added | @OlivierC
09/18/2016  | Document Picker field added | @OlivierC
09/17/2016  | Icon Picker field added | @OlivierC

# The MIT License (MIT)

Copyright (c) 2016 Olivier Carpentier

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

