Attribute VB_Name = "SetLang"
'Option Explicit

'------
' Subs used by Ribbon Buttons OnClick Callback
' OnClick property must be one single string - can not pass lang argument
Sub SetLangUS()
Call SetLang("US")
End Sub

Sub SetLangUK()
Call SetLang("UK")
End Sub

Sub SetLangDE()
Call SetLang("DE")
End Sub

Sub SetLangFR()
Call SetLang("FR")
End Sub
'--------


Sub SetLang(langStr As String)

MSG1 = MsgBox("Do you want to change language of whole presentation vs. only selected slides?", vbYesNoCancel, "All or Only selected Slides?")
If MSG1 = vbYes Then
    Call SetLangPres(ActivePresentation, langStr)
ElseIf MSG1 = vbNo Then
    Call SetLangSelectedSlides(langStr)
End If

End Sub


Private Function SetLangPres(oPres As Presentation, lang As String)
' Reference http://stackoverflow.com/questions/4735765/powerpoint-2007-set-language-on-tables-charts-etc-that-contains-text
' http://stackoverflow.com/questions/37653183/vba-powerpoint-2013-change-presentation-language-including-smartart-objects
' https://support.microsoft.com/en-us/kb/245468

Dim langID As Integer
langID = LangStr2Mso(lang)

    
On Error Resume Next

Dim oSlide As Slide
Dim oShape As Shape

'Set default language in application
oPres.DefaultLanguageID = langID

'Set language in each textbox in each slide
For Each oSlide In oPres.Slides
    Call SetLangSlide(oSlide, langID)
Next

' Update Masters
For Each oShape In oPres.SlideMaster.Shapes
    oShape.TextFrame.TextRange.LanguageID = langID
Next

For Each oShape In oPres.TitleMaster.Shapes
    oShape.TextFrame.TextRange.LanguageID = langID
Next

For Each oShape In oPres.NotesMaster.Shapes
    oShape.TextFrame.TextRange.LanguageID = langID
Next

' MsgBox
MsgBox "Presentation Language was changed to " & lang & ".", vbOKOnly, "SetLanguage"
    
End Function

' -------------------------------------------------------
Sub SetLangSelectedSlides(langStr As String)

Dim langID As Integer
langID = LangStr2Mso(langStr)
Dim oSl As Slide

For Each oSl In ActiveWindow.Selection.SlideRange
    Call SetLangSlide(oSl, langID)
Next oSl

MsgBox "Language of selected Slides (" & CStr(ActiveWindow.Selection.SlideRange.Count) & ") were changed to " & langStr & ".", vbOKOnly, "SetLanguage"
End Sub


Function SetLangSlide(oSlide As Slide, lang As Integer)
Dim oShape As Shape
Dim r, c As Integer
Dim oNode As SmartArtNode
    
On Error Resume Next

For Each oShape In oSlide.Shapes
    'Check first if it is a table
    If oShape.HasTable Then
        For r = 1 To oShape.Table.Rows.Count
            For c = 1 To oShape.Table.Columns.Count
                oShape.Table.Cell(r, c).Shape.TextFrame.TextRange.LanguageID = lang
            Next
        Next
    ElseIf oShape.HasSmartArt Then
        For Each oNode In oShape.SmartArt.AllNodes
            oNode.TextFrame2.TextRange.LanguageID = lang
        Next
    Else
        oShape.TextFrame.TextRange.LanguageID = lang
        For c = 0 To oShape.GroupItems.Count - 1
            oShape.GroupItems(c).TextFrame.TextRange.LanguageID = lang
        Next
    End If
Next

End Function

Function LangStr2Mso(langStr As String) As Integer
' Edit to extend languages supported

If langStr = "US" Then
    LangStr2Mso = msoLanguageIDEnglishUS
ElseIf langStr = "UK" Then
    LangStr2Mso = msoLanguageIDEnglishUK
ElseIf langStr = "DE" Then
    LangStr2Mso = msoLanguageIDGerman
ElseIf langStr = "FR" Then
    LangStr2Mso = msoLanguageIDFrench
End If
End Function
