Attribute VB_Name = "NewMacros"
Sub FormatCSharp()
'DO NOT reorder below
    ChgFontInAllStyles 'set to consolas 9 pt
    FormatBlue
    FormatTurquoise
    FormatViolet
    QuotedTextToRed 'anything in quotes will now be red
    FormatLineGreen 'must call this before FormatLineGrey
    FormatLineGrey
End Sub
Private Sub FormatLineGreen()
Application.ScreenUpdating = False
Dim startingWord As String
'the string to search for
startingWord = "//"

Dim myRange As Range
'Will change selection to the document start
Set myRange = ActiveDocument.Range(ActiveDocument.Range.Start, ActiveDocument.Range.Start)
myRange.Select

While Selection.End < ActiveDocument.Range.End
   If Left(Trim(Selection.Text), Len(startingWord)) = startingWord Then
        With Selection.Font
            .ColorIndex = wdGreen
        End With
    End If

    Selection.MoveDown Unit:=wdLine
    Selection.Expand wdLine

Wend
Application.ScreenUpdating = True
End Sub
Private Sub FormatLineGrey()
Application.ScreenUpdating = False
Dim startingWord As String
'the string to search for
startingWord = "///"

Dim myRange As Range
'Will change selection to the document start
Set myRange = ActiveDocument.Range(ActiveDocument.Range.Start, ActiveDocument.Range.Start)
myRange.Select

While Selection.End < ActiveDocument.Range.End
   If Left(Trim(Selection.Text), Len(startingWord)) = startingWord Then
        With Selection.Font
            .ColorIndex = wdGray50
        End With
    End If

    Selection.MoveDown Unit:=wdLine
    Selection.Expand wdLine

Wend
Application.ScreenUpdating = True
End Sub
Private Sub FormatBlue()
    Dim TextStrng As String
    Dim Result() As String
    TextStrng = "using int bool null public private protected internal void base false true null string using jpeg override model class namespace clone get; set; this"
    Result() = Split(TextStrng)
    
    For i = LBound(Result) To UBound(Result)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
         .Text = Result(i)
         .MatchCase = True
         .MatchWholeWord = True
         .Replacement.Text = ""
         .Replacement.Font.ColorIndex = wdBlue
         .Forward = True
         .Wrap = wdFindContinue
         .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
End Sub
Private Sub FormatTurquoise()
    Dim TextStrng As String
    Dim Result() As String
    TextStrng = "ActionResult HttpStatusCodeResult Directory BindingFlags HttpStatusCode ApplicationDbContext Controller File. Path. Exception MagickImage IDisposable ExifTag IptcTag MagickGeometry Enumerable FileName"
    Result() = Split(TextStrng)
    
    For i = LBound(Result) To UBound(Result)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
         .Text = Result(i)
         .MatchCase = True
         .MatchWholeWord = True
         .Replacement.Text = ""
         .Replacement.Font.ColorIndex = wdTurquoise
         .Forward = True
         .Wrap = wdFindContinue
         .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
End Sub
Private Sub FormatViolet()
    Dim TextStrng As String
    Dim Result() As String
    TextStrng = "SaveChanges Add Save ResizeAndSave SaveChanges httpPostedFile RedirectToAction Dispose IsAjaxRequest Contains OrderByDescending Where Select PartialView HttpPostedFileBase Take HttpNotFound View return ToString GetFileName PhysicalPathFromRootPath"
    Result() = Split(TextStrng)
    
    For i = LBound(Result) To UBound(Result)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
         .Text = Result(i)
         .MatchCase = True
         .MatchWholeWord = True
         .Replacement.Text = ""
         .Replacement.Font.ColorIndex = wdViolet
         .Forward = True
         .Wrap = wdFindContinue
         .MatchWholeWord = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
End Sub
Private Sub QuotedTextToRed()

Dim SmartQtSetting As Boolean

'Make sure smartquotes are turned on
SmartQtSetting = Options.AutoFormatAsYouTypeReplaceQuotes
Options.AutoFormatAsYouTypeReplaceQuotes = True

With Selection.Find
    'Set parameters
    .ClearFormatting
    .Replacement.ClearFormatting
    .Forward = True
    .Wrap = wdFindContinue
    .MatchCase = False
    .MatchWholeWord = False
    .MatchAllWordForms = False
    .MatchSoundsLike = False

    'First do a replace to make sure quotes are all smartquotes
    .Format = False
    .MatchWildcards = False
    .Text = """"
    .Replacement.Text = """"
    .Execute Replace:=wdReplaceAll

    'Then do the wilcard replace
    .Replacement.Font.ColorIndex = wdRed
    .Format = True
    .Wrap = wdFindContinue
    .MatchWildcards = True
    'ChrW(8220) is the open quote character and ChrW(8221) is the close quote
    .Text = "(" & ChrW(8220) & ")(*)(" & ChrW(8221) & ")"
    .Replacement.Text = ChrW(8220) & "\2" & ChrW(8221)
    .Execute Replace:=wdReplaceAll


    'Clear dialog of all non-default settings
    .Text = ""
    .Execute

End With
'Reset options to the way they were
Options.AutoFormatAsYouTypeReplaceQuotes = SmartQtSetting

End Sub
Private Sub ChgFontInAllStyles()
Dim sty As Word.Style
For Each sty In ActiveDocument.Styles
    If sty.InUse And sty.Type = wdStyleTypeParagraph Then
        sty.Font.Name = "Consolas"
        sty.Font.Size = 9
        sty.NoSpaceBetweenParagraphsOfSameStyle = True
    End If
Next
End Sub
