# Microsoft Word Macro Collection

## Tool Descriptions
This collection consists of several macros written in Visual Basic for Applications (VBA) for use with Microsoft Office Word. These macros that have been programmed to complete basic, automatable editing tasks:
* __SpaceErase with EgieCommas.__ Removes multiple spaces, unnecessary spaces before carriage returns, and spaces around dashes. Adds commas after i.e. and e.g.
* __AcronymFlag.__ Flags acronyms consisting of at least two capital letters.
* __Unit Flag.__ Flags units of measure and math symbols.
* __Blacklist Flag.__ Flags commonly misused phrases.

## Installation Instructions
To use these macros with your own Microsoft Office user profile, you will first have to open Word and make sure you can see Developer options in your Word ribbon. If you cannot see Developer options, go to *File > Options > Customize Ribbon* and make sure the *Developer* box is checked under the options for the Main Tabs.

Once you've enabled the Developer menu, select it and find the *Visual Basic* button at the far left. Click on it, and it will take you to the Visual Basic for Applications editor. To enable macros for all of your Word documents instead of just the present document, select *Normal* from the project hierarchy view on the left sidebar.

Once you are in Normal, go to *Insert > Module*. A blank box will pop up: this is where you will paste the code for the macros of your choice. Simply navitage to the `.VBA` file for the macro you want from this GitHub and save the code exactly as it appears. Save when you are done, then exit the VBA editor.

Once you are back to your Word document, to run the new macro, simply click the *Macros* button, and a menu will pop up prompting you to select a macro. Once you've selected a macro, hit *Run*.

## Product Notes
The flagging macros work by adding highlighting to the actual formatting of the text in your document, unlike Word's *Find* function (CTRL + F), which applies a temporary yellow highlight to its search results. You will want to remember to remove this highlighting as you edit because it will continue to be visible to clients, other readers, etc. unless it is changed.

## Modification Instructions
### AcronymFlag
This macro works by finding and highlighting all instances of two or more consecutive capital letters. I use it to make sure authors are not using acronyms without defining them properly. If you'd prefer to search for acronyms of a different character length, edit the following line of code:

`.Text = "[A-Z]{2,}"`

In this code fragment, `{2,}` allows the macro to search for consecutive capital letters of two or more characters. Change the `2` to adjust the minimum character length for acronyms you'd like to flag, and add a number after the `,` if you'd like to only search for acronyms up to a certain maximum length.

If you'd like to change the highlight color to a specfic shade, add the following line before the `Ends With` line:

`Options.DefaultHighlightColorIndex = wdBrightGreen`

Replace `wdBrightGreen` with any color on this [list of default color constants](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa212740(v%3Doffice.11)).

### UnitFlag
This macro works by finding matching words in a document to a string list defined in the `.VBA` file and highlighting all matches. I use it to make sure numerals are being used with units of time, measurement, etc.

If you 'd like to change the list of words the macro screens for, edit the word list in the line `StrFind = "list"`. Separate words with a comma, no spaces.

If you'd like to change the highlight color to a specfic shade, change this line:

`Options.DefaultHighlightColorIndex = wdYellow`

Replace `wdYellow` with any color on this [list of default color constants](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa212740(v%3Doffice.11)).

This macro currently only flags units. If you would like to flag *and replace* units, remove the `' ` ahead of the line `' StrRepl = "list"`, which comments out the line and prevents the computer from reading that code. Populate the `StrRepl` list with the words you want to replace the original units with, and make sure they appear in the same list order as their counterparts.

For example:

```
StrFind = "minutes,seconds,hours"
StrRepl = StrFind
StrRepl = "min,s,hr"
```

### BlacklistFlag
This macro works just like the UnitFlag macro, except it screens for a list of phrases that are commonly misused by writers. I use this macro to call these words to my attention so I can further evaluate whether the word is being used correctly or not. 

If you know you have instances of words that you never want to use, you can uncomment out the `StrRepl` line (see UnitFlag directions, above) and add the value you would like the macro to replace. For example, Bureau publications do not use "et al." in citations, so if you are not already using the [Bureau Citation Style](https://github.com/emnharris/BureauEditing/tree/master/citations), you can screen for "et al." and replace with "and others".

If you'd like to change the highlight color to a specfic shade, change this line:

`Options.DefaultHighlightColorIndex = wdTurquoise`

Replace `wdTurquoise` with any color on this [list of default color constants](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa212740(v%3Doffice.11)).
