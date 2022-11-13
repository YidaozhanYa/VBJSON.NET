' VBJSON is a VB6 adaptation of the VBA JSON project at http://code.google.com/p/vba-json/
' Some bugs fixed, speed improvements added for VB6 by Michael Glaser (vbjson@ediy.co.nz)
' BSD Licensed

' Ported to dotnet by YidaozhanYa

Option Explicit On

Public Class JsonParser
    Const INVALID_JSON As Long = 1
    Const INVALID_OBJECT As Long = 2
    Const INVALID_ARRAY As Long = 3
    Const INVALID_BOOLEAN As Long = 4
    Const INVALID_NULL As Long = 5
    Const INVALID_KEY As Long = 6
    Const INVALID_RPC_CALL As Long = 7

    Private psErrors As String = ""


    Public Function GetParserErrors() As String
        Return psErrors
    End Function

    Public Sub ClearParserErrors()
        psErrors = ""
    End Sub

    '
    '   parse string and create JSON object
    '

    Public Function Parse(ByRef str As String) As Object
        Dim index As Long
        index = 1
        psErrors = ""
        On Error Resume Next
        Call SkipChar(str, index)
        Select Case Mid(str, index, 1)
            Case "{"
                Parse = ParseObject(str, index)
            Case "["
                Parse = ParseArray(str, index) 'Collection
            Case Else
                psErrors = "Invalid JSON"
                Parse = New Dictionary(Of String, Object)
        End Select
        Return Parse
    End Function
    '
    '   parse collection of key/value
    '
    Private Function ParseObject(ByRef str As String, ByRef index As Long) As Dictionary(Of String, Object)

        ParseObject = New Dictionary(Of String, Object)
        Dim sKey As String

        ' "{"
        Call skipChar(str, index)
        If Mid(str, index, 1) <> "{" Then
            psErrors = psErrors & "Invalid Object at position " & index & " : " & Mid(str, index) & vbCrLf
            Exit Function
        End If

        index += 1

        Do
            Call skipChar(str, index)
            If "}" = Mid(str, index, 1) Then
                index += 1
                Exit Do
            ElseIf "," = Mid(str, index, 1) Then
                index += 1
                Call skipChar(str, index)
            ElseIf index > Len(str) Then
                psErrors = psErrors & "Missing '}': " & Right(str, 20) & vbCrLf
                Exit Do
            End If


            ' add key/value pair
            sKey = parseKey(str, index)
            On Error Resume Next

            ParseObject.Add(sKey, parseValue(str, index))
            If Err.Number <> 0 Then
                psErrors = psErrors & Err.Description & ": " & sKey & vbCrLf
                Exit Do
            End If
        Loop
    End Function
    '
    '   parse list
    '
    Private Function ParseArray(ByRef str As String, ByRef index As Long) As Collection

        Dim RetVal As New Collection

        ' "["
        Call skipChar(str, index)
        If Mid(str, index, 1) <> "[" Then
            psErrors = psErrors & "Invalid Array at position " & index & " : " + Mid(str, index, 20) & vbCrLf
            Exit Function
        End If

        index += 1

        Do

            Call skipChar(str, index)
            If "]" = Mid(str, index, 1) Then
                index += 1
                Exit Do
            ElseIf "," = Mid(str, index, 1) Then
                index += 1
                Call skipChar(str, index)
            ElseIf index > Len(str) Then
                psErrors = psErrors & "Missing ']': " & Right(str, 20) & vbCrLf
                Exit Do
            End If

            ' add value
            On Error Resume Next
            RetVal.Add(parseValue(str, index))
            If Err.Number <> 0 Then
                psErrors = psErrors & Err.Description & ": " & Mid(str, index, 20) & vbCrLf
                Exit Do
            End If
        Loop
        Return RetVal

    End Function

    '
    '   parse string / number / object / array / true / false / null
    '
    Private Function ParseValue(ByRef str As String, ByRef index As Long)
        Dim RetVal
        Call SkipChar(str, index)

        Select Case Mid(str, index, 1)
            Case "{"
                RetVal = ParseObject(str, index)
            Case "["
                RetVal = ParseArray(str, index)
            Case """", "'"
                RetVal = ParseString(str, index)
            Case "t", "f"
                RetVal = ParseBoolean(str, index)
            Case "n"
                RetVal = ParseNull(str, index)
            Case Else
                RetVal = ParseNumber(str, index)
        End Select
        Return RetVal
    End Function

    '
    '   parse string
    '
    Private Function ParseString(ByRef str As String, ByRef index As Long) As String

        Dim quote As String
        Dim Character As String = ""
        Dim Code As String

        Dim SB As String = ""

        Call skipChar(str, index)
        quote = Mid(str, index, 1)
        index += 1

        Do While index > 0 And index <= Len(str)
            Character = Mid(str, index, 1)
            Select Case Character
                Case "\"
                    index += 1
                    Character = Mid(str, index, 1)
                    Select Case Character
                        Case """", "\", "/", "'"
                            SB &= Character
                            index += 1
                        Case "b"
                            SB &= vbBack
                            index += 1
                        Case "f"
                            SB &= vbFormFeed
                            index += 1
                        Case "n"
                            SB &= vbLf
                            index += 1
                        Case "r"
                            SB &= vbCr
                            index += 1
                        Case "t"
                            SB &= vbTab
                            index += 1
                        Case "u"
                            index += 1
                            Code = Mid(str, index, 4)
                            SB &= ChrW(Val("&h" + Code))
                            index += 4
                    End Select
                Case quote
                    index += 1
                    Return SB
                Case Else
                    SB &= Character
                    index += 1
            End Select
        Loop

        Return SB

    End Function

    '
    '   parse number
    '
    Private Function ParseNumber(ByRef str As String, ByRef index As Long) As Decimal

        Dim Value As String = ""
        Dim Character As String = ""

        Call skipChar(str, index)
        Do While index > 0 And index <= Len(str)
            Character = Mid(str, index, 1)
            If InStr("+-0123456789.eE", Character) Then
                Value &= Character
                index += 1
            Else
                Return CDec(Value)
            End If
        Loop
    End Function

    '
    '   parse true / false
    '
    Private Function ParseBoolean(ByRef str As String, ByRef index As Long) As Boolean
        Dim RetVal As Boolean
        Call skipChar(str, index)
        If Mid(str, index, 4) = "true" Then
            RetVal = True
            index += 4
        ElseIf Mid(str, index, 5) = "false" Then
            RetVal = False
            index += 5
        Else
            psErrors = psErrors & "Invalid Boolean at position " & index & " : " & Mid(str, index) & vbCrLf
        End If
        Return RetVal

    End Function

    '
    '   parse null
    '
    Private Function ParseNull(ByRef str As String, ByRef index As Long)
        Dim RetVal As String = ""
        Call skipChar(str, index)
        If Mid(str, index, 4) = "null" Then
            RetVal = Nothing
            index += 4
        Else
            psErrors = psErrors & "Invalid null value at position " & index & " : " & Mid(str, index) & vbCrLf
        End If
        Return RetVal

    End Function

    Private Function ParseKey(ByRef str As String, ByRef index As Long) As String

        Dim dquote As Boolean
        Dim squote As Boolean
        Dim Character As String
        Dim RetVal As String = ""

        Call skipChar(str, index)
        Do While index > 0 And index <= Len(str)
            Character = Mid(str, index, 1)
            Select Case Character
                Case """"
                    dquote = Not dquote
                    index += 1
                    If Not dquote Then
                        Call skipChar(str, index)
                        If Mid(str, index, 1) <> ":" Then
                            psErrors = psErrors & "Invalid Key at position " & index & " : " & RetVal & vbCrLf
                            Exit Do
                        End If
                    End If
                Case "'"
                    squote = Not squote
                    index += 1
                    If Not squote Then
                        Call skipChar(str, index)
                        If Mid(str, index, 1) <> ":" Then
                            psErrors = psErrors & "Invalid Key at position " & index & " : " & RetVal & vbCrLf
                            Exit Do
                        End If
                    End If
                Case ":"
                    index += 1
                    If Not dquote And Not squote Then
                        Exit Do
                    Else
                        RetVal &= Character
                    End If
                Case Else
                    If InStr(vbCrLf & vbCr & vbLf & vbTab & " ", Character) Then
                    Else
                        RetVal &= Character
                    End If
                    index += 1
            End Select
        Loop
        Return RetVal

    End Function

    '
    '   skip special character
    '
    Private Sub SkipChar(ByRef str As String, ByRef index As Long)
        Dim bComment As Boolean
        Dim bStartComment As Boolean
        Dim bLongComment As Boolean
        Do While index > 0 And index <= Len(str)
            Select Case Mid(str, index, 1)
                Case vbCr, vbLf
                    If Not bLongComment Then
                        bStartComment = False
                        bComment = False
                    End If

                Case vbTab, " ", "(", ")"

                Case "/"
                    If Not bLongComment Then
                        If bStartComment Then
                            bStartComment = False
                            bComment = True
                        Else
                            bStartComment = True
                            bComment = False
                            bLongComment = False
                        End If
                    Else
                        If bStartComment Then
                            bLongComment = False
                            bStartComment = False
                            bComment = False
                        End If
                    End If

                Case "*"
                    If bStartComment Then
                        bStartComment = False
                        bComment = True
                        bLongComment = True
                    Else
                        bStartComment = True
                    End If

                Case Else
                    If Not bComment Then
                        Exit Do
                    End If
            End Select

            index += 1
        Loop

    End Sub

    Public Function Stringify(ByRef obj As Object) As String
        Dim SB As String = ""
        Select Case VarType(obj)
            Case VariantType.Null
                SB &= "null"
            Case VariantType.Date
                SB &= """" & CStr(obj) & """"
            Case VariantType.String
                SB &= """" & Encode(obj) & """"
            Case VariantType.Boolean
                If obj Then SB &= "true" Else SB &= "false"
            Case vbVariant, vbArray, vbArray + vbVariant
                Dim sEB
                SB &= MultiArray(obj, 1, "", sEB)
            Case Else

                Dim bFI As Boolean
                Dim i As Long

                bFI = True
                If TypeName(obj).StartsWith("Dictionary") Then

                    SB &= "{"
                    Dim keys As Dictionary(Of String, Object).KeyCollection
                    keys = obj.Keys
                    For i = 0 To obj.Count - 1
                        If bFI Then bFI = False Else SB &= ","
                        Dim key As String
                        key = keys(i)
                        SB &= """" & key & """:" & Stringify(obj.Item(key))
                    Next i
                    SB &= "}"

                ElseIf TypeName(obj).StartsWith("Collection") Then
                    SB &= "["
                    Dim Value
                    For Each Value In obj
                        If bFI Then bFI = False Else SB &= ","
                        SB &= Stringify(Value)
                    Next Value
                    SB &= "]"

                End If
        End Select
        Debug.Print(SB)
        Return SB
    End Function

    Private Function Encode(str) As String

        Dim SB As String = ""
        Dim i As Long
        Dim j As Long
        Dim aL1 As Array = {&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9}
        Dim aL2 As Array = {&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74}
        Dim c As String
        Dim p As Boolean

        For i = 1 To Len(str)
            p = True
            c = Mid(str, i, 1)
            For j = 0 To 7
                If c = Chr(aL1(j)) Then
                    SB &= "\" & Chr(aL2(j))
                    p = False
                    Exit For
                End If
            Next

            If p Then
                Dim a
                a = AscW(c)
                If a > 31 And a < 127 Then
                    SB &= c
                ElseIf a > -1 Or a < 65535 Then
                    SB &= "\u" & StrDup(4 - Len(Hex(a)), "0") & Hex(a)
                End If
            End If
        Next

        Return SB

    End Function

    Private Function MultiArray(aBD, iBC, sPS, ByRef sPT)   ' Array BoDy, Integer BaseCount, String PoSition

        Dim iDU As Long
        Dim iDL As Long
        Dim i As Long
        Dim SB As String = ""
        Dim sPB1 As String = "", sPB2 As String = ""  ' String PointBuffer1, String PointBuffer2
        Dim OutOfRange As Boolean = False

        Try
            iDL = LBound(aBD, iBC)
            iDU = UBound(aBD, iBC)

        Catch ex As IndexOutOfRangeException
            sPB1 = sPT & sPS
            For i = 1 To Len(sPB1)
                If i <> 1 Then sPB2 &= ","
                sPB2 &= Mid(sPB1, i, 1)
            Next
            '        multiArray = multiArray & JSONToString(Eval("aBD(" & sPB2 & ")"))
            SB &= Stringify(aBD(sPB2))
            OutOfRange = True
        End Try
        If OutOfRange Then
            sPT = sPT & sPS
            SB &= "["
            For i = iDL To iDU
                SB &= MultiArray(aBD, iBC + 1, i, sPT)
                If i < iDU Then SB &= ","
            Next
            SB &= "]"
            sPT = Left(sPT, iBC - 2)
        End If
        Return SB
    End Function

    ' Miscellaneous JSON functions

    Public Function StringToJSON(st As String) As String

        Const FIELD_SEP = "~"
        Const RECORD_SEP = "|"

        Dim sFlds As String
        Dim sRecs As String = ""
        Dim lRecCnt As Long
        Dim lFld As Long
        Dim fld As Object
        Dim rows As Object

        lRecCnt = 0
        If st = "" Then
            StringToJSON = "null"
        Else
            rows = Split(st, RECORD_SEP)
            For lRecCnt = LBound(rows) To UBound(rows)
                sFlds = ""
                fld = Split(rows(lRecCnt), FIELD_SEP)
                For lFld = LBound(fld) To UBound(fld) Step 2
                    sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld(lFld) & """:""" & toUnicode(fld(lFld + 1) & "") & """")
                Next 'fld
                sRecs &= IIf((Trim(sRecs) <> ""), "," & vbCrLf, "") & "{" & sFlds & "}"
            Next 'rec
            Return ("( {""Records"": [" & vbCrLf & sRecs & vbCrLf & "], " & """RecordCount"":""" & lRecCnt & """ } )")
        End If
    End Function

    Public Function ToUnicode(str As String) As String
        Dim x As Long
        Dim uStr As String = ""
        Dim uChrCode As Integer

        For x = 1 To Len(str)
            uChrCode = Asc(Mid(str, x, 1))
            Select Case uChrCode
                Case 8   ' backspace
                    uStr &= "\b"
                Case 9 ' tab
                    uStr &= "\t"
                Case 10  ' line feed
                    uStr &= "\n"
                Case 12  ' formfeed
                    uStr &= "\f"
                Case 13 ' carriage return
                    uStr &= "\r"
                Case 34 ' quote
                    uStr &= "\"""
                Case 39  ' apostrophe
                    uStr &= "\'"
                Case 92 ' backslash
                    uStr &= "\\"
                Case 123, 125  ' "{" and "}"
                    uStr &= ("\u" & Right("0000" & Hex(uChrCode), 4))
                Case Is < 32, Is > 127 ' non-ascii characters
                    uStr &= ("\u" & Right("0000" & Hex(uChrCode), 4))
                Case Else
                    uStr &= Chr(uChrCode)
            End Select
        Next
        Return uStr
        Exit Function
    End Function

End Class
