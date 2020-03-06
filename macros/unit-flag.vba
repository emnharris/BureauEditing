Sub UnitFlag()

Dim StrFind As String
Dim StrRepl As String
Dim i As Long

' In StrFind and StrRepl, add words between the quote marks. Separate with a comma, no spaces.
' To replace a word with another and highlight it, put the new word in the StrRepl list in the same position as the word in the StrFind list you want to replace. Make sure you uncomment the StrRepl line if you do this.

StrFind = "minutes,seconds,hours,days,weeks,months,years,percent,inches,>,<,=,+,±,−,×,≥,≤"
StrRepl = StrFind
' StrRepl = "minutes,seconds,hours,days,weeks,months,years,percent,inches,>,<,=,+,±,−,×,≥,≤"
Set RngTxt = Selection.Range

' Set highlight color
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

