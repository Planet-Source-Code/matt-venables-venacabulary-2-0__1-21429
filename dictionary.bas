Attribute VB_Name = "dictionary"
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Public Function Replace(strExpression As String, strFind As String, strReplace As String)
    Dim intX As Integer
    If (Len(strExpression) - Len(strFind)) >= 0 Then
        For intX = 1 To Len(strExpression)
            If Mid(strExpression, intX, Len(strFind)) = strFind Then
                strExpression = Left(strExpression, (intX - 1)) + strReplace + Mid(strExpression, intX + Len(strFind), Len(strExpression))
            End If
        Next
    End If
    Replace = strExpression
End Function
Public Function Splitter(SplitString As String, SplitLetter As String) As Variant
'Splits a string at a letter "," in this case

    ReDim SplitArray(1 To 1) As Variant
    Dim TempLetter As String
    Dim TempSplit As String
    Dim i As Integer
    Dim X As Integer
    Dim StartPos As Integer
    
    SplitString = SplitString & SplitLetter


    For i = 1 To Len(SplitString)
        TempLetter = Mid(SplitString, i, Len(SplitLetter))


        If TempLetter = SplitLetter Then
            TempSplit = Mid(SplitString, (StartPos + 1), (i - StartPos) - 1)


            If TempSplit <> "" Then
                X = X + 1
                ReDim Preserve SplitArray(1 To X) As Variant
                SplitArray(X) = TempSplit
            End If
            StartPos = i
        End If
    Next i
    Splitter = SplitArray
End Function
Public Sub FormDrag(Form As Form)
    ReleaseCapture
    Call SendMessage(Form.hwnd, &HA1, 2, 0&)
    If Form.Left < 400 Then Form.Left = 0
    If Form.Top < 400 Then Form.Top = 0
    If Val(Form.Width + Form.Left) > Val(Screen.Width - 400) Then Form.Left = Val(Screen.Width - Form.Width)
    If Val(Form.Height + Form.Top) > Val(Screen.Height - 400) Then Form.Top = Val(Screen.Height - Form.Height)
End Sub
Sub DefineWord(word As String, Inet1 As Inet, short As Boolean, message As Boolean)
     Dim definition As String
     Dim spot1, spot2 As Integer
     On Error GoTo err:
'Load Webpage
     Inet1.URL = "http://dictionary.msn.com/find/entry.asp?search=" & word
     def = Inet1.OpenURL(Inet1.URL)
     If InStr(1, def, "No matches found for") Then GoTo nomatch:
     beginspot = InStr(1, def, "<div class='dictionary'>") + 24
     EndSpot = InStr(beginspot, def, "Encarta® World English Dictionary")
'Get Definition
     spot1 = beginspot
     spot2 = beginspot + 2
     spot1 = InStr(spot1, def, "-1")
     spot1 = InStr(spot1, def, "<b>") + 3
     spot2 = InStr(spot1, def, "</b>")
     If InStr(1, Mid$(def, spot1, spot2 - spot1), "<") Then GoTo nomatch:
     If short = True Then
           definition = Replace(Mid$(def, spot1, (spot2 - spot1)), ":", " ")
           GoTo output
     End If
     spot1 = spot2 + 1
     spot1 = InStr(spot1, def, "</font>") + 7
     spot2 = InStr(spot1, def, "<br />")
     If InStr(spot1, def, "<img") < InStr(spot1, def, "<br />") Then
          spot2 = InStr(spot1, def, "<img")
     End If
     If InStr(spot1, def, "<a") < spot2 Then
          spot2 = InStr(spot1, def, "<a") - 1
     End If
     definition = Mid$(def, spot1, spot2 - spot1)
'Output
output:
    MakeHTML word, definition, dict.Label1
    Exit Sub
err:
nomatch:
    definition = "  "
    If message = True Then MsgBox "No matches found for: " & word & "."
    MakeHTML word, definition, dict.Label1
    Exit Sub
End Sub



Function WordoftheDay(Inet1 As Inet)
Dim word As String
Dim spot1, spot2 As Integer
word = Inet1.OpenURL("http://features.learningkingdom.com/word/")
spot1 = InStr(1, word, "<b><font size=") + 18
spot2 = InStr(spot1, word, "[")
WordoftheDay = Mid$(word, spot1, spot2 - spot1)
End Function
Sub MakeHTML(word As String, definition As String, Label1 As Label)
    Label1.Caption = Label1.Caption & "<b>" & word & "</b> - " & definition & "<br>"
End Sub
