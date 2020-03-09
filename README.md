![Bureau of Economic Geology logo](https://i.ibb.co/t2RWGkp/Screen-Shot-2016-07-11-at-3-44-03-PM.png)

# Bureau Editing Resources
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

See the collection [read-me](https://github.com/emnharris/BureauEditing/blob/master/macros/readme.md) for more information.

### Bureau Citation Style
With the help of [BibWord](https://archive.codeplex.com/?p=bibword), users can now install the Bureau of Economic Geology style to use with Word's Citation and Bibliography tools.

This tool is currently unable to handle alphabetical sorting of references and should only be used with a modified workflow to accommodate this shortcoming. See the program [read-me](https://github.com/emnharris/BureauEditing/tree/master/citations) for more information.

## Future Goals
* AutoCorrect and Find-Replace `.VBA` tools (Macros)
* year suffixes and author bars for bibliographies (Bureau Citation Style)
