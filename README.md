# BureauEditing
## Introduction
This repository is a collection of automated tools for accomplishing editorial tasks. These tools have been written to follow the editorial styling of the Bureau of Economic Geology.

Download the most recent version (Feb 2020) of the *Bureau of Economic Geology Stylebook* [here](https://github.com/emnharris/BureauEditing/blob/master/Bureau-style-guide.pdf).

## Tools
### Macros
The [macro collection](https://github.com/emnharris/BureauEditing/tree/master/macros) consists of several macros for use with Microsoft Word that have been programmed to complete basic, automatable editing tasks:
* __SpaceErase with EgieCommas.__ Removes multiple spaces, spaces before carriage returns, and spaces around dashes and hyphens. Adds commas after i.e. and e.g.
* __AcronymFlag.__ Flags acronyms consisting of at least two capital letters.
* __Unit Flag.__ Flags units of measure and math symbols.
* __Blacklist Flag.__ Flags commonly misused phrases.

See the collection read-me for more information.

### Citation Style
This style tool is based on the [2009 BibWord Tool](https://archive.codeplex.com/?p=bibword). Installing the `.XSL` will give users the option to format citations and bibliography references in Word according the the Bureau of Economic Geology reference style.

This tool is currently unable to handle alphabetical sorting of references and should only be used with a modified workflow to accommodate this shortcoming. See the program [read-me](https://github.com/emnharris/BureauEditing/tree/master/citations) for more information.

## Future Goals
* AutoCorrect and Find-Replace `.VBA` tools
* alphabetization for bibliographies
* year extensions for bibliographies
