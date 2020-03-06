Sub UnitFlag()

Dim StrFind As String
Dim StrRepl As String
Dim i As Long

' In StrFind and StrRepl, add words between the quote marks, separate with a comma, NO spaces
' To only highlight the found words (i.e. not replace with other words), either use StrRepl = StrFind OR use the SAME words in the same order in the StrRepl list as for the StrFind list; comment/uncomment to reflect the one you're using
' To replace a word with another and highlight it, put the new word in the StrRepl list in the SAME position as the word in the StrFind list you want to replace; comment/uncomment to reflect the one you're using

StrFind = "minutes,seconds,hours,days,weeks,months,years,percent,inches"
StrRepl = StrFind
' StrRepl = "minutes,seconds,hours,days,weeks,months,years,percent,inches"
Set RngTxt = Selection.Range

' Set highlight color - options are listed here: https://docs.microsoft.com/en-us/office/vba/api/word.wdcolorindex
' main ones are wdYellow, wdTurquoise, wdBrightGreen, wdPink
Options.DefaultHighlightColorIndex = wdYellow

Selection.HomeKey wdStory

' Clear existing formatting and settings in Find and Replace fields
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting

With ActiveDocument.Content.Find
  .Format = True
  .MatchWholeWord = True
  .MatchAllWordForms = False
  .MatchWildcards = False
  .Wrap = wdFindContinue
  .Forward = True
  For i = 0 To UBound(Split(StrFind, ","))
    .Text = Split(StrFind, ",")(i)
    .Replacement.Highlight = True
    .Replacement.Text = Split(StrRepl, ",")(i)
    .Execute Replace:=wdReplaceAll
  Next i
Options.DefaultHighlightColorIndex = wdBrightGreen
End With
End Sub

