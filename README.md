# OSLO365 Plugin

This repository holds the software for the OSLO plugin that was created for Microsoft Office Word. The goal of this project was to create a Word add-in, which uses the OSLO Vocabularies.

## Install it locally

In order to run this plugin locally, first this repository must be cloned `git clone https://github.com/Informatievlaanderen/OSLO365-plugin.git`. Then the following commands must be executed in order to be able to run the plugin:
```
> npm install
> npm update
> npn run build
> npm run start
```
The `npm run build` command is necessary because it will generate the manifest.xml file. This file is needed in order to install the plugin in your Word.

## License

This code is released under the [MIT License](http://opensource.org/licenses/MIT)
