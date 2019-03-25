# SWGoH Territory Battles Sheets [![clasp](https://img.shields.io/badge/built%20with-clasp-4285f4.svg)](https://github.com/google/clasp)

Google Apps Script SWGoH Territory Battles Google Spreadsheet
- [Instructions Manual](https://docs.google.com/document/d/1InDnFdYdlRDpejeiFkGi6HxMEWmBbfX_wAvWRf1MCME/edit?usp=sharing)
- [Spreadsheet template](https://docs.google.com/spreadsheets/d/17aFVEsOPe-HzoOZm0UizCSEbkc7fS_-4uCETRxkUxLw/edit?usp=sharing)

## Getting Started

- Make sure you have installed a recent version of [Node.js](https://nodejs.org/en/)
- install Clasp and TypeScript `npm install -g @google/clasp typescript`
- clone this project and edit `.clasp.json` with the scriptId of your copy of the Spreadsheet template (see above)

The `src` directory contains the bundled code that is pushed as Google Apps Script into the Google speadsheet.

## Build your own copy

Prerequisite:
- Have [Google Clasp CLI](https://developers.google.com/apps-script/guides/clasp) installed
- (Recommended) setup your [TypeScript](https://developers.google.com/apps-script/guides/typescript) IDE

Steps:
1. Make a local copy of the GitHub repository.
1. Edit the file `.clasp.json` with the scriptId of your own copy of the Territory Battles spreadsheet
1. If needed, edit the source `.ts` files under the `src/` directory
1. Use Clasp CLI to push the transpiled TypeScript into your Google script
1. (Optional) issue a Pull Request for your changes to be added to the official release

## About the Developer

Reach him on [Discord](https://discord.gg/ywzJEaQ)

## Contributions

Contributions and feature requests are welcome.

## License

[MIT License](https://github.com/labnol/apps-script-starter/blob/master/LICENSE) (c) Guillaume Contesso a.k.a. PopGoesTheWza