VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   11640
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Command3 
      Caption         =   "Find Else"
      Height          =   330
      Left            =   2880
      TabIndex        =   5
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Find End"
      Height          =   330
      Left            =   1200
      TabIndex        =   4
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox txtOutput 
      Height          =   2775
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   4800
      Width           =   11535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run"
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox txtSource 
      Height          =   4095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form1.frx":000C
      Top             =   360
      Width           =   11535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Output:"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   4560
      Width           =   525
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------'
' Author    : sysdzw                             '
' E-mail    : sysdzw@163.com                     '
' Bolg      : http://blog.csdn.net/sysdzw         '
' QQ        : 171977759                          '
' Date      : 2010-4-6                           '
'------------------------------------------------'
Option Explicit
Dim reg As Object
Dim matchs As Object, match As Object
Dim reg2 As Object
Dim matchs2 As Object, match2 As Object
Dim regForTest As Object
Dim dic As Object
Dim sc As Object
Dim strSource$, vSource, i%
Dim rVar$, rValuePattern$, rVarKeyWord$, rString$, rNumber$, rLoop$
Dim lngCurrentLine As Integer
'the code line's type
Private Enum ECodeType
    eCTSetValue
    eCTOutPut
    eCTOutPutEx
    eCTIF
    eCTELSE
    eCTEND
    eCTFOR
    eCTGOTO
    eCTGOTOLine
    eCTUnknow
End Enum
'the variable's type
Private Enum EVariableType
    eVTString
    eVTNumber
    eVTBool
    eVTDate
    eVTTime
End Enum
Private Sub Form_Load()
    Set sc = CreateObject("ScriptControl")
    sc.Language = "VBScript"
    
    Set reg = CreateObject("vbscript.regexp")
    reg.Global = True
    reg.IgnoreCase = True
    
    Set reg2 = CreateObject("vbscript.regexp")
    reg2.Global = True
    reg2.IgnoreCase = True
    
    Set regForTest = CreateObject("vbscript.regexp")
    regForTest.Global = True
    regForTest.IgnoreCase = True
    
    Set dic = CreateObject("Scripting.Dictionary")
    
    rVar = "[a-zA-Z_\u4e00-\u9fa5][\da-zA-Z_\u4e00-\u9fa5]*?"
    rString = """[\s\S]*?"""
    rNumber = "\d+\.?\d*"
    rValuePattern = "(" & rNumber & "|" & rString & "|" & rVar & ")"
    rLoop = "[a-zA-z\Z_]+?\d*"
    rVarKeyWord = "(puts|print|printf)"
End Sub
Private Sub Command1_Click()
    Dim strLineCode$
    Dim strVarName$, strVarValue$
    Dim sBefore$, sAfter$
    Dim intCodeLineType As Integer
    Dim strCondition As String
    Dim blnCondition As Boolean
    
    If txtOutput.Text <> "" And Right(txtOutput.Text, 2) <> vbCrLf Then txtOutput.Text = txtOutput.Text & vbCrLf
    txtOutput.SelStart = Len(txtOutput.Text)
    txtOutput.SelText = ">Start run program" & vbCrLf
    strSource = txtSource.Text
    vSource = Split(strSource, vbCrLf)
    dic.removeall
    lngCurrentLine = 0
    
    Do
        strLineCode = Trim(vSource(lngCurrentLine))
        If strLineCode <> "" Then
            intCodeLineType = getType(strLineCode)
            Select Case intCodeLineType
                Case eCTSetValue
                    reg.Pattern = "^(" & rVar & ")\s*?\=(.*)$"
                    Set matchs = reg.Execute(strLineCode)
                    sBefore = matchs(0).SubMatches(0)
                    sAfter = matchs(0).SubMatches(1)
                    
                    If regTest(sAfter, "^" & rNumber & "$") Then
                        dic(sBefore) = eVTNumber & sAfter
                    ElseIf regTest(sAfter, "^" & rString & "$") Then
                        dic(sBefore) = eVTString & Mid(sAfter, 2, Len(sAfter) - 2)
                    ElseIf regTest(sAfter, "^" & rVar & "$") Then
                        If Not dic.Exists(sAfter) Then
                            dic(sAfter) = eVTString
                            dic(sBefore) = eVTString
                        Else
                            dic(sBefore) = dic(sAfter)
                        End If
                    Else 'is a expression
                        dic(sBefore) = getExpressionResult(sAfter)
                    End If
                Case eCTOutPut
                    reg.Pattern = "^" & rVarKeyWord & "\s+?" & rValuePattern & "$"
                    Set matchs = reg.Execute(strLineCode)
                    txtOutput.SelText = getTrueValue(matchs(0).SubMatches(1)) & vbCrLf
                Case eCTOutPutEx
                    reg.Pattern = "^" & rVarKeyWord & "\s+?(.*?)$"
                    Set matchs = reg.Execute(strLineCode)
                    txtOutput.SelText = Mid(getExpressionResult(matchs(0).SubMatches(1)), 2) & vbCrLf
                Case eCTIF
                    reg.Pattern = "^\s*?if(.+)$"
                    Set matchs = reg.Execute(strLineCode)
                    strCondition = matchs(0).SubMatches(0)
                    
                    reg.Pattern = "^\s*?" & rValuePattern & "\s*?(<>|>\=|<\=|\=|>|<)\s*?" & rValuePattern & "\s*?$"
                    Set matchs = reg.Execute(strCondition)
                    sBefore = getTrueValue(matchs(0).SubMatches(0))
                    sAfter = getTrueValue(matchs(0).SubMatches(2))
                    
                    If Left(sBefore, 1) = eVTNumber Or Left(sAfter, 1) = eVTNumber Then
                        sBefore = Mid(sBefore, 2)
                        sAfter = Mid(sAfter, 2)
                        Select Case matchs(0).SubMatches(1)
                            Case "<>": blnCondition = (Val(sBefore) <> Val(sAfter))
                            Case ">=": blnCondition = (Val(sBefore) >= Val(sAfter))
                            Case "<=": blnCondition = (Val(sBefore) <= Val(sAfter))
                            Case ">": blnCondition = (Val(sBefore) > Val(sAfter))
                            Case "<": blnCondition = (Val(sBefore) < Val(sAfter))
                            Case "=":  blnCondition = (Val(sBefore) = Val(sAfter))
                        End Select
                    Else
                        Select Case matchs(0).SubMatches(1)
                            Case "<>": blnCondition = (sBefore <> sAfter)
                            Case ">=": blnCondition = (sBefore >= sAfter)
                            Case "<=": blnCondition = (sBefore <= sAfter)
                            Case ">": blnCondition = (sBefore > sAfter)
                            Case "<": blnCondition = (sBefore < sAfter)
                            Case "=":  blnCondition = (sBefore = sAfter)
                        End Select
                    End If
                    
                    If Not blnCondition Then
                        lngCurrentLine = getTheElseLine(lngCurrentLine)
                    End If
                Case eCTELSE
                    lngCurrentLine = getTheElseLine(lngCurrentLine, "end")
                Case eCTGOTO
                    reg.Pattern = "^\s*?goto\s*?(" & rLoop & ")\s*$"
                    Set matchs = reg.Execute(strLineCode)
                    lngCurrentLine = getTheGotoLine(matchs(0).SubMatches(0) & ":")
                Case eCTFOR
                    
                Case eCTUnknow
                    txtOutput.SelText = "Error at Line " & lngCurrentLine + 1 & ": """ & strLineCode & """" & vbCrLf
                    txtOutput.SelText = ">Exit code: -1" & vbCrLf
                    Exit Sub
            End Select
        End If
        lngCurrentLine = lngCurrentLine + 1
        If lngCurrentLine > UBound(vSource) Then Exit Do
    Loop
    txtOutput.SelText = ">Exit code: 0" & vbCrLf
End Sub
'get the line code's type
Private Function getType(ByVal strLineCode$) As ECodeType
    Dim vTmp
    If regTest(strLineCode, "^" & rVar & "\s*?\=\s*?.*$") Then
        getType = eCTSetValue
        Exit Function
    End If
    
    If regTest(strLineCode, "^\s*?if\s*?.*$") Then
        getType = eCTIF
        Exit Function
    End If
    
    If regTest(strLineCode, "^\s*?for\s*?.*$") Then
        getType = eCTFOR
        Exit Function
    End If
    
    If regTest(strLineCode, "^\s*?else\s*?.*$") Then
        getType = eCTELSE
        Exit Function
    End If
    
    If regTest(strLineCode, "^\s*?end\s*?.*$") Then
        getType = eCTEND
        Exit Function
    End If
    
    If regTest(strLineCode, "^\s*?goto\s*?.*$") Then
        getType = eCTGOTO
        Exit Function
    End If
    
    If regTest(strLineCode, "^\s*?" & rLoop & "\:\s*$") Then
        getType = eCTGOTOLine
        Exit Function
    End If
    If regTest(strLineCode, "^" & rVarKeyWord & "\s+?.+?$") Then
        If regTest(strLineCode & " + ", "^" & rVarKeyWord & "\s+?" & "(.+?)\s+?\+\s+?") Then
            getType = eCTOutPutEx
            Exit Function
        End If
    
        If regTest(strLineCode, "^" & rVarKeyWord & "\s+?" & rValuePattern & "$") Then
            getType = eCTOutPut
        End If
        
        getType = eCTOutPut
        Exit Function
    End If
    
    getType = eCTUnknow
End Function
'the para has 3 kinds
'1.num 2.string 3.var
Private Function getTrueValue(sValue$) As String
    If regTest(sValue, "^" & rNumber & "$") Then
        getTrueValue = eVTNumber & sValue
    ElseIf regTest(sValue, "^" & rString & "$") Then
        getTrueValue = eVTString & Mid(sValue, 2, Len(sValue) - 2)
    ElseIf regTest(sValue, "^" & rVar & "$") Then
        If Not dic.Exists(sValue) Then
            dic(sValue) = eVTString
            getTrueValue = eVTString
        Else
            getTrueValue = dic(sValue)
        End If
    End If
End Function
'the agrs is a expression
Private Function getExpressionResult(strExp As String) As String
    Dim strTmp$, i%, isNumExp As Boolean
    Dim strResult$, strTrueExp$, strExpTmp$
    
    strTmp = Replace(strExp, "(", "")
    strTmp = Replace(strTmp, ")", "")
    
    reg.Pattern = "(.+?)\s*?[\+\-\*/]\s*?"
    Set matchs = reg.Execute(strTmp & "+")
    
    reg2.Pattern = "(.+?)\s*?[\+\-\*/]\s*?"
    Set matchs2 = reg2.Execute(strExp & "+")
    isNumExp = True
    For i = 0 To matchs.Count - 1
        strExpTmp = Trim(matchs(i).SubMatches(0))
        If regTest(strExpTmp, "^" & rNumber & "$") Then
            strTrueExp = strTrueExp & matchs2(i)
        ElseIf regTest(strExpTmp, "^" & rString & "$") Then
            strTrueExp = strTrueExp & matchs2(i)
            isNumExp = False
        ElseIf regTest(strExpTmp, "^" & rVar & "$") Then
            If Not dic.Exists(strExpTmp) Then
                dic(strExpTmp) = eVTString
            Else
                If Left(dic(strExpTmp), 1) <> eVTNumber Then isNumExp = False
                strTrueExp = strTrueExp & Replace(matchs2(i), strExpTmp, Mid(dic(strExpTmp), 2))
            End If
        End If
    Next
    
    If isNumExp Then
        If Right(strTrueExp, 1) = "+" Then strTrueExp = Left(strTrueExp, Len(strTrueExp) - 1)
        getExpressionResult = eVTNumber & WZcalc(strTrueExp)
    Else
        reg.Pattern = "(.+?)\s*?\+\s*?"
        Set matchs = reg.Execute(strExp & "+")
        For i = 0 To matchs.Count - 1
            strExpTmp = Trim(matchs(i).SubMatches(0))
            If regTest(strExpTmp, "^" & rNumber & "$") Then
                strResult = strResult & strExpTmp
            ElseIf regTest(strExpTmp, "^" & rString & "$") Then
                strResult = strResult & Mid(strExpTmp, 2, Len(strExpTmp) - 2)
            ElseIf regTest(strExpTmp, "^" & rVar & "$") Then
                If Not dic.Exists(strExpTmp) Then
                    dic(strExpTmp) = eVTString
                Else
                    strResult = strResult & Mid(dic(strExpTmp), 2)
                End If
            End If
        Next
        getExpressionResult = eVTString & strResult
    End If
End Function
'get the "else" line,if it hadn't "else" it'll return the "end" line.
'if the para "strKey" is "end",this means it'll find the "end" line.
Private Function getTheElseLine(intCurIfLine As Integer, Optional strKey = "else") As Integer
    Dim v, i%
    Dim isIFClosed() As Boolean
    Dim intCurrentIF As Integer
    
    v = Split(txtSource.Text, vbCrLf)
    
    i = intCurIfLine + 1
    intCurrentIF = 0
    ReDim Preserve isIFClosed(intCurrentIF)
    isIFClosed(intCurrentIF) = False
    
    Do
        If regTest(v(i), "^\s*?if\s*?.*$") Then
            ReDim Preserve isIFClosed(UBound(isIFClosed) + 1)
            intCurrentIF = UBound(isIFClosed)
            isIFClosed(intCurrentIF) = False
        ElseIf regTest(v(i), "^\s*?else\s*$") Then
            If intCurrentIF = 0 Then
                getTheElseLine = i
                If strKey = "else" Then Exit Function
            End If
        ElseIf regTest(v(i), "^\s*?end\s*$") Then
            If intCurrentIF = 0 Then
                getTheElseLine = i
                Exit Function
            Else
                isIFClosed(intCurrentIF) = True
                Do
                    intCurrentIF = intCurrentIF - 1
                    If isIFClosed(intCurrentIF) = False Then Exit Do
                Loop
            End If
        End If
        i = i + 1
        If i > UBound(v) Then Exit Do
    Loop
End Function
'get the "goto" line
Private Function getTheGotoLine(strGotoTag As String) As Integer
    Dim v, i%
    Dim intLen As Integer
    intLen = Len(strGotoTag)
    
    v = Split(txtSource.Text, vbCrLf)
    For i = 0 To UBound(v)
        If Left(v(i), intLen) = strGotoTag Then
            getTheGotoLine = i
            Exit Function
        End If
    Next
End Function
'test the string is mathed the pattern
Private Function regTest(ByVal sData$, sPattern$) As Boolean
    regForTest.Pattern = sPattern
    regTest = regForTest.Test(sData)
End Function
Public Function WZcalc(Tmpstr$) As Double
   WZcalc = sc.Eval(Tmpstr)
End Function
Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    txtSource.Height = (Me.ScaleHeight - txtSource.Top) * 0.65
    txtSource.Width = Me.ScaleWidth - 45
    Label1.Top = txtSource.Top + txtSource.Height + 45
    txtOutput.Move 0, Label1.Top + Label1.Height + 45, Me.ScaleWidth - 45, Me.ScaleHeight - Label1.Top - Label1.Height - 90
End Sub
Private Sub txtSource_GotFocus()
    Dim C As Object
    On Error Resume Next
    For Each C In Me.Controls
        C.TabStop = False
    Next
End Sub
Private Sub txtSource_LostFocus()
    Dim C As Object
    On Error Resume Next
    For Each C In Me.Controls
        C.TabStop = True
    Next
End Sub
Private Sub Command2_Click()
    MsgBox getTheElseLine(1)
End Sub
Private Sub Command3_Click()
    Dim v, i%, t$
    Dim isIFClosed() As Boolean
    Dim intCurrentIF As Integer
    
    v = Split(txtSource.Text, vbCrLf)
    
    i = 2
    intCurrentIF = 0
    ReDim Preserve isIFClosed(intCurrentIF)
    isIFClosed(intCurrentIF) = False
    
    Do
        If InStr(v(i), "if") > 0 Then
            ReDim Preserve isIFClosed(UBound(isIFClosed) + 1)
            intCurrentIF = UBound(isIFClosed)
            isIFClosed(intCurrentIF) = False
        ElseIf InStr(v(i), "else") > 0 Then
            If intCurrentIF = 0 Then
                MsgBox "else at line: " & i + 1
            End If
        ElseIf InStr(v(i), "end") > 0 Then
            If intCurrentIF = 0 Then
                MsgBox "end at line: " & i + 1
            Else
                isIFClosed(intCurrentIF) = True
                Do
                    intCurrentIF = intCurrentIF - 1
                    If isIFClosed(intCurrentIF) = False Then Exit Do
                Loop
            End If
        End If
        i = i + 1
        If i > UBound(v) Then Exit Do
    Loop
End Sub
