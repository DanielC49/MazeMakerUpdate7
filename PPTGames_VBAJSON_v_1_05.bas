Attribute VB_Name = "JSONVBA"
' ==============================================================
'
' ##############################################################
' #                                                            #
' #                     PPTGames VBA JSON                      #
' #                       VBAJSON Module                       #
' #                                                            #
' ##############################################################
'
' » version 1.05
'
' » https://pptgamespt.wixsite.com/pptg-coding/json-vba
'
' ===============================================================


Option Explicit
Option Compare Binary

Private ind As Long, ws As Integer, lb As String

Private Type json_options
    AllowUnquotedKeys As Boolean
    UseEscapeChars As Boolean
    IgnoreUndefinedExpr As Boolean
    DefinedExpr As New List
End Type

Public JSONOptions As json_options

Function ParseJSON(ByVal jsonString As String) As Object

    jsonString = Replace(jsonString, vbNewLine, "")
    jsonString = Replace(jsonString, vbLf, "")
    jsonString = Replace(jsonString, vbTab, "")

    Dim i As Long, i2 As Long, state_exp As String, curw As String, args As New List
    
    For i = 1 To Len(jsonString)
    
        If state_exp <> "" Then
        
            If Mid(jsonString, i, 1) = state_exp Then
                state_exp = ""
                curw = curw & Mid(jsonString, i, 1)
            ElseIf Mid(jsonString, i, 1) = "\" Then
                If InStr(1, "\""'bfnrtu", Mid(jsonString, i + 1, 1)) <> 0 Then
                    Select Case Mid(jsonString, i + 1, 1)
                        Case "b": curw = curw & vbBack
                        Case "f": curw = curw & vbFormFeed
                        Case "n": curw = curw & vbNewLine
                        Case "r": curw = curw & vbCr
                        Case "t": curw = curw & vbTab
                        Case "u": curw = curw & ChrW(val("&h" & Mid(jsonString, i + 2, 4))): i = i + 4
                        Case Else: curw = curw & Mid(jsonString, i + 1, 1)
                    End Select
                    i = i + 1
                Else
                    curw = curw & Mid(jsonString, i, 1)
                End If
            Else
                curw = curw & Mid(jsonString, i, 1)
            End If
            
        Else
        
            If Mid(jsonString, i, 1) = """" Or Mid(jsonString, i, 1) = "'" Then
                state_exp = Mid(jsonString, i, 1)
                curw = curw & Mid(jsonString, i, 1)
            ElseIf InStr(1, "[]{}:,", Mid(jsonString, i, 1)) <> 0 Then
                For i2 = i To Len(jsonString)
                    If Mid(jsonString, i2, 1) <> " " And Mid(jsonString, i2, 1) <> ":" Then i2 = -1: Exit For
                    If Mid(jsonString, i2, 1) = ":" Then Exit For
                Next
                If curw <> "" Then args.AddItem HandleExpression(curw, i2, jsonString)
                curw = ""
                args.AddItem Mid(jsonString, i, 1)
            ElseIf Mid(jsonString, i, 1) <> " " Then
                curw = curw & Mid(jsonString, i, 1)
            End If
        
        End If
    
    Next
    
    If args.Length > 0 Then
    
        Select Case args.Items(0)
        Case "["
            If args(1) = "]" Then
                Set ParseJSON = New List
            Else
                Set ParseJSON = ParseJSON_Array(args.Slice(1, args.Length - 2))
            End If
        Case "{"
            If args(1) = "}" Then
                Set ParseJSON = New Dictionary
                ParseJSON.RemoveAll
            Else
                Set ParseJSON = ParseJSON_Object(args.Slice(1, args.Length - 2))
            End If
        End Select
    
    Else
    
        Set ParseJSON = Nothing
    
    End If

End Function

Private Function HandleExpression(e As String, i As Long, c As String) As Variant
    If i > -1 Then
        If JSONOptions.AllowUnquotedKeys = False And InStr(1, """'", Left(e, 1)) = 0 Then
            Err.Raise 1, "JSONVBA", "Unquoted key: " & vbNewLine & vbNewLine & Left(c, 15) & IIf(Len(c) > 15, " ...", "") & vbNewLine & "^" & vbNewLine & "Expected: "" or '"
        Else
            If InStr(1, """'", Left(e, 1)) = 0 Then
                HandleExpression = e
            Else
                HandleExpression = Mid(e, 2, Len(e) - 2)
            End If
        End If
    Else
        If InStr(1, """'", Left(e, 1)) = 0 Then
            If e = "true" Or e = "false" Then
                HandleExpression = e = "true"
            ElseIf e = "undefined" Then
                HandleExpression = e
            ElseIf IsNumeric(Replace(e, ".", ",")) Then
                HandleExpression = CDbl(Replace(e, ".", ","))
            ElseIf JSONOptions.DefinedExpr.IndexOf(e) > -1 Then
                HandleExpression = "?UDExpr:" & JSONOptions.DefinedExpr.Items(JSONOptions.DefinedExpr.IndexOf(e))
            Else
                If JSONOptions.IgnoreUndefinedExpr = True Then
                    HandleExpression = "?UDExpr:" & e
                Else
                    Err.Raise 1, "JSONVBA", "'" & e & "' is not defined."
                End If
            End If
        Else
            HandleExpression = Mid(e, 2, Len(e) - 2)
        End If
    End If
End Function

Private Function ParseJSON_Array(e) As List

    Set ParseJSON_Array = New List

    Dim args As New List, arr As New List
    args.Items = e

    Dim i As Long, i2 As Long, s As Long
    
    For i = 0 To args.Length - 1

        If args.Items(i) = "[" Then
            If args(i + 1) = "]" Then
                arr.AddItem New List
                i = i + 1
            Else
                s = 0
                For i2 = i To args.Length - 1
                    If args(i2) = "[" Then s = s + 1
                    If args(i2) = "]" And s > 0 Then s = s - 1
                    If args(i2) = "]" And s = 0 Then Exit For
                Next
                arr.AddItem ParseJSON_Array(args.Slice(i + 1, i2 - 1))
                i = i2
            End If
        ElseIf args.Items(i) = "{" Then
            If args(i + 1) = "}" Then
                arr.AddItem New Dictionary
                i = i + 1
            Else
                s = 0
                For i2 = i To args.Length - 1
                    If args(i2) = "{" Then s = s + 1
                    If args(i2) = "}" And s > 0 Then s = s - 1
                    If args(i2) = "}" And s = 0 Then Exit For
                Next
                arr.AddItem ParseJSON_Object(args.Slice(i + 1, i2 - 1))
                i = i2
            End If
        ElseIf args.Items(i) <> "]" And args.Items(i) <> "}" And args.Items(i) <> "," Then
            arr.AddItem args.Items(i)
        End If
    
    Next
    ParseJSON_Array.Items = arr.Items

End Function

Private Function ParseJSON_Object(e) As Dictionary

    Set ParseJSON_Object = New Dictionary

    Dim args As New List
    args.Items = e

    Dim i As Long, i2 As Long, s As Long, key As String
    
    For i = 0 To args.Length - 1

        If i > 0 Then
            If IsObject(args.Items(i - 1)) = False Then
                If args.Items(i - 1) = ":" Then
                    If IsObject(args.Items(i)) = False Then
                        If args.Items(i) = "[" Then
                            If args(i + 1) = "]" Then
                                ParseJSON_Object.Add key, New List
                                i = i + 1
                            Else
                                s = 0
                                For i2 = i To args.Length - 1
                                    If args(i2) = "[" Then s = s + 1
                                    If args(i2) = "]" And s > 0 Then s = s - 1
                                    If args(i2) = "]" And s = 0 Then Exit For
                                Next
                                ParseJSON_Object.Add key, ParseJSON_Array(args.Slice(i + 1, i2 - 1))
                                i = i2
                            End If
                        ElseIf args.Items(i) = "{" Then
                            If args(i + 1) = "}" Then
                                ParseJSON_Object.Add key, New Dictionary
                                i = i + 1
                            Else
                                s = 0
                                For i2 = i To args.Length - 1
                                    If args(i2) = "{" Then s = s + 1
                                    If args(i2) = "}" And s > 0 Then s = s - 1
                                    If args(i2) = "}" And s = 0 Then Exit For
                                Next
                                ParseJSON_Object.Add key, ParseJSON_Object(args.Slice(i + 1, i2 - 1))
                                i = i2
                            End If
                        Else
                            ParseJSON_Object.Add key, args.Items(i)
                        End If
                    End If
                End If
            End If
        End If
        
        If i < args.Length - 1 Then
            If IsObject(args.Items(i + 1)) = False Then
                If args.Items(i + 1) = ":" Then key = args.Items(i)
            End If
        End If
    
    Next

End Function

Function StringJSON(ByVal jsonObject As Object, Optional WhiteSpace As Integer = 4, Optional UseLineBreaks As Boolean = True) As String

    ind = 0
    ws = WhiteSpace
    lb = IIf(UseLineBreaks, vbNewLine, "")
    
    If VarType(jsonObject) = vbObject And ListOrDic(jsonObject) = "dic" Then
    
        If jsonObject.Count > 0 Then
            StringJSON = StringJSON & "{" & String(ws, " ") & StringJSON_Object(jsonObject) & "}"
        Else
            StringJSON = StringJSON & "{ }"
        End If
        
    ElseIf VarType(jsonObject) = 8204 Or ListOrDic(jsonObject) = "list" Then
    
        If jsonObject.Length > 0 Then
            StringJSON = StringJSON & "[" & String(ws, " ") & StringJSON_Array(jsonObject) & "]"
        Else
            StringJSON = StringJSON & "[ ]"
        End If
        
    End If
    
End Function

Private Function StringJSON_Object(ByVal jsonObject As Dictionary) As String

    ind = ind + 1

    Dim i As Long, arr As New List
    
    arr.Items = jsonObject.Items
    
    For i = LBound(jsonObject.Items) To UBound(jsonObject.Items)
    
        If VarType(arr.Items(i)) = vbObject And ListOrDic(arr.Items(i)) = "dic" Then

            If jsonObject.Items(i).Count > 0 Then
                StringJSON_Object = StringJSON_Object & lb & String(ind * ws, " ") & IIf(JSONOptions.AllowUnquotedKeys = True, "", """") & _
                    jsonObject.Keys(i) & IIf(JSONOptions.AllowUnquotedKeys = True, "", """") & ": {" & StringJSON_Object(jsonObject.Items(i)) & _
                    IIf(i = UBound(jsonObject.Items), String(ind * ws, " ") & "}" & lb, String(ind * ws, " ") & "},")
            Else
                StringJSON_Object = StringJSON_Object & lb & String(ind * ws, " ") & IIf(JSONOptions.AllowUnquotedKeys = True, "", """") & _
                    jsonObject.Keys(i) & IIf(JSONOptions.AllowUnquotedKeys = True, "", """") & ": {" & _
                    IIf(i = UBound(jsonObject.Items), "}" & lb, "},")
            End If
            
        ElseIf VarType(jsonObject.Items(i)) = 8204 Or ListOrDic(arr.Items(i)) = "list" Then

            If jsonObject.Items(i).Length > 0 Then
                StringJSON_Object = StringJSON_Object & lb & String(ind * ws, " ") & IIf(JSONOptions.AllowUnquotedKeys = True, "", """") & _
                    jsonObject.Keys(i) & IIf(JSONOptions.AllowUnquotedKeys = True, "", """") & ": [" & StringJSON_Array(jsonObject.Items(i)) & _
                    IIf(i = UBound(jsonObject.Items), String(ind * ws, " ") & "]" & lb, String(ind * ws, " ") & "],")
            Else
                StringJSON_Object = StringJSON_Object & lb & String(ind * ws, " ") & IIf(JSONOptions.AllowUnquotedKeys = True, "", """") & _
                    jsonObject.Keys(i) & IIf(JSONOptions.AllowUnquotedKeys = True, "", """") & ": [ " & _
                    IIf(i = UBound(jsonObject.Items), "]" & lb, "],")
            End If
                
        ElseIf VarType(jsonObject.Items(i)) = vbBoolean Then
        
            StringJSON_Object = StringJSON_Object & lb & String(ind * ws, " ") & IIf(JSONOptions.AllowUnquotedKeys = True, "", """") & _
            jsonObject.Keys(i) & IIf(JSONOptions.AllowUnquotedKeys = True, "", """") & ": " & IIf(jsonObject.Items(i), "true", "false") & IIf(i = UBound(jsonObject.Items), lb, ",")
            
        ElseIf Left(jsonObject.Items(i), 8) = "?UDExpr:" And JSONOptions.DefinedExpr.IndexOf(Mid(jsonObject.Items(i), 9)) > -1 Then
        
            StringJSON_Object = StringJSON_Object & lb & String(ind * ws, " ") & IIf(JSONOptions.AllowUnquotedKeys = True, "", """") & _
                jsonObject.Keys(i) & IIf(JSONOptions.AllowUnquotedKeys = True, "", """") & ": " & JSONOptions.DefinedExpr.Items(JSONOptions.DefinedExpr.IndexOf(Mid(jsonObject.Items(i), 9))) & _
                IIf(i = UBound(jsonObject.Items), lb, ",")
                
        ElseIf Left(jsonObject.Items(i), 8) = "?UDExpr:" Then
        
            StringJSON_Object = StringJSON_Object & lb & String(ind * ws, " ") & IIf(JSONOptions.AllowUnquotedKeys = True, "", """") & _
                jsonObject.Keys(i) & IIf(JSONOptions.AllowUnquotedKeys = True, "", """") & ": " & Mid(jsonObject.Items(i), 9) & _
                IIf(i = UBound(jsonObject.Items), lb, ",")
            
        ElseIf IsNumeric(jsonObject.Items(i)) = False Then
            
            StringJSON_Object = StringJSON_Object & lb & String(ind * ws, " ") & IIf(JSONOptions.AllowUnquotedKeys = True, "", """") & _
                jsonObject.Keys(i) & IIf(JSONOptions.AllowUnquotedKeys = True, "", """") & ": """ & StrJSON_UE(jsonObject.Items(i)) & _
                IIf(i = UBound(jsonObject.Items), """" & lb, """,")
            
        Else
        
            StringJSON_Object = StringJSON_Object & lb & String(ind * ws, " ") & IIf(JSONOptions.AllowUnquotedKeys = True, "", """") & _
            jsonObject.Keys(i) & IIf(JSONOptions.AllowUnquotedKeys = True, "", """") & ": " & Replace(jsonObject.Items(i), ",", ".") & IIf(i = UBound(jsonObject.Items), lb, ",")
        
        End If
    
    Next
    
    ind = ind - 1

End Function

Private Function StringJSON_Array(ByVal jsonObject As List) As String

    ind = ind + 1

    Dim i As Long
    
    For i = 0 To jsonObject.Length - 1
    
        If VarType(jsonObject(i)) = vbObject And ListOrDic(jsonObject(i)) = "dic" Then
            
            If jsonObject(i).Count > 0 Then
                StringJSON_Array = StringJSON_Array & lb & String(ind * ws, " ") & "{" & StringJSON_Object(jsonObject(i)) & String(ind * ws, " ") & IIf(i = jsonObject.Length - 1, "}" & lb, "},")
            Else
                StringJSON_Array = StringJSON_Array & lb & String(ind * ws, " ") & "{ " & IIf(i = jsonObject.Length - 1, "}" & lb, "},")
            End If
            
        ElseIf VarType(jsonObject(i)) = 8204 Or ListOrDic(jsonObject(i)) = "list" Then
        
            If jsonObject(i).Length > 0 Then
                StringJSON_Array = StringJSON_Array & lb & String(ind * ws, " ") & "[" & StringJSON_Array(jsonObject(i)) & String(ind * ws, " ") & IIf(i = jsonObject.Length - 1, "]", "],") & lb
            Else
                StringJSON_Array = StringJSON_Array & lb & String(ind * ws, " ") & "[ " & IIf(i = jsonObject.Length - 1, "]" & lb, "],")
            End If
            
        ElseIf VarType(jsonObject(i)) = vbBoolean Then
            
            StringJSON_Array = StringJSON_Array & lb & String(ind * ws, " ") & IIf(jsonObject.Items(i), "true", "false") & IIf(i = jsonObject.Length - 1, "" & lb, ", ")
            
        ElseIf Left(jsonObject(i), 8) = "?UDExpr:" And JSONOptions.DefinedExpr.IndexOf(Mid(jsonObject(i), 9)) > -1 Then
        
            StringJSON_Array = StringJSON_Array & lb & String(ind * ws, " ") & JSONOptions.DefinedExpr.Items(JSONOptions.DefinedExpr.IndexOf(Mid(jsonObject(i), 9))) & IIf(i = jsonObject.Length - 1, lb, ", ")
            
        ElseIf Left(jsonObject(i), 8) = "?UDExpr:" Then
        
            StringJSON_Array = StringJSON_Array & lb & String(ind * ws, " ") & Mid(jsonObject(i), 9) & IIf(i = jsonObject.Length - 1, lb, ", ")
            
        ElseIf IsNumeric(jsonObject(i)) = False Then
            
            StringJSON_Array = StringJSON_Array & lb & String(ind * ws, " ") & """" & StrJSON_UE(jsonObject(i)) & IIf(i = jsonObject.Length - 1, """" & lb, """, ")
        
        Else
            
            StringJSON_Array = StringJSON_Array & lb & String(ind * ws, " ") & Replace(jsonObject(i), ",", ".") & IIf(i = jsonObject.Length - 1, "" & lb, ", ")
        
        End If
    
    Next
    
    ind = ind - 1

End Function

Private Function StrJSON_UE(ByVal e As String) As String
    Dim i As Long
    For i = 1 To Len(e)
        Select Case Mid(e, i, 1)
        Case vbBack: StrJSON_UE = StrJSON_UE & IIf(JSONOptions.UseEscapeChars = True, "\b", "")
        Case vbFormFeed: StrJSON_UE = StrJSON_UE & IIf(JSONOptions.UseEscapeChars = True, "\f", "")
        Case vbLf: StrJSON_UE = StrJSON_UE & IIf(JSONOptions.UseEscapeChars = True, "\n", "")
        Case vbTab: StrJSON_UE = StrJSON_UE & IIf(JSONOptions.UseEscapeChars = True, "\t", "")
        Case "\": StrJSON_UE = StrJSON_UE & IIf(JSONOptions.UseEscapeChars = True, "\\", "")
        Case """": StrJSON_UE = StrJSON_UE & IIf(JSONOptions.UseEscapeChars = True, "\""", "")
        Case "'": StrJSON_UE = StrJSON_UE & IIf(JSONOptions.UseEscapeChars = True, "\'", "")
        Case Else: If Mid(e, i + 1, 1) <> vbLf Then StrJSON_UE = StrJSON_UE & Mid(e, i, 1)
        End Select
    Next
End Function

Private Function ListOrDic(obj) As String
    Err.Clear
    On Error GoTo handler
    If IsObject(obj) Then
        If obj.Count > -1 Then ListOrDic = "dic"
    Else
        ListOrDic = "none"
    End If
    Exit Function
handler: If ListOrDic <> "dic" Then ListOrDic = "list"
End Function
