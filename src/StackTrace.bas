Attribute VB_Name = "StackTrace"
#Const DEBUG_MODE = 1
#Const DEBUG_PRINT_MODE = 0
#Const NO_TRACE = 1

Private StackLevel As Long
Private DebugTrace As Collection
Private Counter As Object

Sub WriteStackTrace()
    Dim ws As Worksheet
    Set ws = Sheet_DebugTrace
    If Not DebugTrace Is Nothing Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Dim OutputRow As Long
        OutputRow = 1
        Do While Len(ws.Cells(OutputRow, 1).FormulaR1C1) > 0
            OutputRow = OutputRow + 1
        Loop
        Dim v As Variant
        For Each v In DebugTrace
            Dim l As StackTraceLog
            Set l = v
            With ws
                .Cells(OutputRow, 1) = l.Level
                .Cells(OutputRow, 2) = l.modName
                .Cells(OutputRow, 3) = l.procName
                .Cells(OutputRow, 4) = l.argList
                .Cells(OutputRow, 5) = l.retValue
                ' ����
                .Cells(OutputRow, 6).FormulaR1C1 = "=CONCAT(REPT(""|"",RC1-1),""+"",RC2,""."",RC3,""("",RC4,IF(RC5="""","")"","")=""),RC5)"
                ' �n�C�p�[�����N
                .Hyperlinks.Add .Cells(OutputRow, 6), "", .Cells(OutputRow, 6).AddressLocal
                OutputRow = OutputRow + 1
            End With
        Next
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        ws.Calculate
        Set DebugTrace = Nothing
        Set Counter = Nothing
        StackLevel = 0
    End If
End Sub
Sub PushStackTrace(modName As String, procName As String, ParamArray args() As Variant)
#If DEBUG_MODE Then
    
    If DebugTrace Is Nothing Then Set DebugTrace = New Collection
    If Counter Is Nothing Then Set Counter = CreateObject("Scripting.Dictionary")
    
    StackLevel = StackLevel + 1
    
    Dim argList As String
    argList = ""
    Dim i As Long
    For i = LBound(args) To UBound(args) Step 2
        If i > LBound(args) Then argList = argList & ", "
        If (i + 1) <= UBound(args) Then
            argList = argList & args(i) & ":=" & ArgsToString(args(i + 1))
        Else
            argList = argList & args(i)
        End If
    Next
    
    Dim l As New StackTraceLog
    With l
        .Level = StackLevel
        .modName = modName
        .procName = procName
        .argList = argList
    End With
    
    DebugTrace.Add l
#If DEBUG_PRINT_MODE Then
    Debug.Print String(StackLevel, "|") & "+" & modName & "." & procName & "(" & argList & ")"
#End If
    
    Dim keyName As String
    keyName = modName & "." & procName
    If Counter.Exists(keyName) Then
        Counter.Item(keyName) = Counter.Item(keyName) + 1
        If Counter.Item(keyName) = 10000 Then
            Debug.Print "[StackTrace]" & keyName & "�̌ďo�񐔂�10000��𒴂��܂����B"
            MsgBox "[StackTrace]" & keyName & "�̌ďo�񐔂�10000��𒴂��܂����B�p�t�H�[�}���X�̉e�����傫�����ߓ��Y�v���V�[�W���̃f�o�b�O�R�[�h�폜���������߂��܂��B"
            Stop
        End If
    Else
        Counter.Add keyName, 1
    End If

#End If
End Sub
Sub PopStackTrace(modName As String, procName As String, Optional returnValue As Variant)
#If DEBUG_MODE Then
    
    Dim retValue As String
    If Not IsMissing(returnValue) Then
        retValue = ArgsToString(returnValue)
        Dim iRow As Long
        iRow = DebugTrace.Count
        Dim l As StackTraceLog
        Set l = DebugTrace(iRow)
        Do While l.Level > StackLevel
            iRow = iRow - 1
            Set l = DebugTrace(iRow)
        Loop
        Dim canWriteResult As Boolean
        canWriteResult = True
        If l.Level <> StackLevel Then canWriteResult = False
        If l.modName <> modName Then canWriteResult = False
        If l.procName <> procName Then canWriteResult = False
        If canWriteResult Then
            l.retValue = retValue
        End If
    End If
    
    StackLevel = StackLevel - 1
    
#End If
End Sub
Private Function ArgsToString(arg As Variant) As String
    On Error Resume Next
    Dim result As String
    Dim argType As String
    argType = TypeName(arg)
    Select Case argType
        Case "String"
            result = """" & arg & """"
        Case Else
            result = CStr(arg)
    End Select
    If Err.Number <> 0 Then
        Err.Clear
        If Right(argType, 2) = "()" Then
            argType = Replace(argType, "()", PrintArrayBounds(arg))
        End If
        result = Replace("[%]", "%", argType)
    End If
    On Error GoTo 0
    ArgsToString = result
End Function
Private Function PrintArrayBounds(arg As Variant) As String
    Dim i As Long
    Dim res As String
    Dim d As String
    Err.Clear
    On Error Resume Next
    i = 0
    d = ""
    res = "("
    Do While Err.Number = 0
        i = i + 1
        res = res & d & LBound(arg, i) & ".." & UBound(arg, i)
        d = ","
    Loop
    res = res & ")"
    Err.Clear
    On Error GoTo 0
    PrintArrayBounds = res
End Function


Sub EnableDEBUGMODE()
    Dim vbEnv As Object
    Dim vbComp As Variant
    Dim vbMod As Object
    Dim modName As String
    ' VBE�̎Q��
    Set vbEnv = Application.VBE
    ' ���ׂĂ�VBComponents�𒀎���������
    For Each vbComp In vbEnv.ActiveVBProject.VBComponents
        ' Module�����擾
        Set vbMod = vbComp.CodeModule
        modName = vbMod.Name
        Dim doAddST As Boolean
        doAddST = True
        If modName = "StackTraceLog" Then doAddST = False
        If modName = "Sheet_DebugTrace" Then doAddST = False
        If doAddST Then
            If InStr(1, vbMod.Lines(1, 1), "#Const DEBUG_MODE") > 0 Then
                vbMod.DeleteLines 1
            End If
            vbMod.InsertLines 1, "#Const DEBUG_MODE = 1"
        End If
    Next
End Sub
Sub DisableDEBUGMODE()
    Dim vbEnv As Object
    Dim vbComp As Variant
    Dim vbMod As Object
    Dim modName As String
    ' VBE�̎Q��
    Set vbEnv = Application.VBE
    ' ���ׂĂ�VBComponents�𒀎���������
    For Each vbComp In vbEnv.ActiveVBProject.VBComponents
        ' Module�����擾
        Set vbMod = vbComp.CodeModule
        modName = vbMod.Name
        Dim doAddST As Boolean
        doAddST = True
        If modName = "StackTraceLog" Then doAddST = False
        If modName = "Sheet_DebugTrace" Then doAddST = False
        If doAddST Then
            If InStr(1, vbMod.Lines(1, 1), "#Const DEBUG_MODE") > 0 Then
                vbMod.DeleteLines 1
            End If
        End If
    Next
End Sub

Sub AddStackTrace()
    Dim vbEnv As Object
    Dim vbComp As Variant
    Dim vbMod As Object
    Dim modName As String
    Dim procName As String
    Dim lineNum As Long
    Dim numLines As Long
    Dim defineLineNum As Long
    Dim startLineNum As Long
    Dim endLineNum As Long
    Dim procKind As Long
    Dim exitSub As Collection
    Dim i As Long
    Dim shCodeName As String
    shCodeName = "Sheet_DebugTrace"
    ' VBE�̎Q��
    Set vbEnv = Application.VBE
    ' ���ׂĂ�VBComponents�𒀎���������
    For Each vbComp In vbEnv.ActiveVBProject.VBComponents
        ' Module�����擾
        Set vbMod = vbComp.CodeModule
        modName = vbMod.Name
        Dim doAddST As Boolean
        doAddST = True
        If modName = "StackTrace" Then doAddST = False
        If modName = "StackTraceLog" Then doAddST = False
        If modName = shCodeName Then doAddST = False
        For i = 1 To vbMod.CountOfLines
            If i > 10 Then Exit For
            If InStr(1, vbMod.Lines(i, 1), "#Const NO_TRACE = 1") > 0 Then doAddST = False
        Next
        If doAddST Then
            ' �s�����擾
            lineNum = vbMod.CountOfDeclarationLines + 1
            numLines = vbMod.CountOfLines
            ' ���ׂẴv���V�[�W���𒲍�
            Do While lineNum < numLines
                ' ���̃v���V�[�W���̎擾
                procKind = 0
                procName = vbMod.ProcOfLine(lineNum, procKind)
                If procName <> "" Then
                    Set exitSub = New Collection
                    ' �v���V�[�W����`�̊J�n�s�̎擾
                    lineNum = vbMod.ProcBodyLine(procName, procKind)
                    defineLineNum = lineNum
                    Do While Right(vbMod.Lines(lineNum, 1), 1) = "_"
                        lineNum = lineNum + 1
                    Loop
                    startLineNum = lineNum + 1
                    
                    ' �v���V�[�W����`�̏I���s�̎擾
                    lineNum = vbMod.procStartLine(procName, procKind) + vbMod.ProcCountLines(procName, procKind)
                    endLineNum = 0
                    Do While lineNum >= startLineNum
                        If Left(vbMod.Lines(lineNum, 1), 3) = "End" Then
                            endLineNum = lineNum
                            Exit Do
                        End If
                        lineNum = lineNum - 1
                    Loop
                    
                    ' �}���ۃ`�F�b�N�i��`�`End�܂ł�1�s�ŏ�����Ă���C���^�[�t�F�C�X�N���X�̉��z���\�b�h�ɂ͑}������̂��ʓ|�Ȃ̂ő}�����Ȃ����߁B�j
                    If endLineNum > startLineNum Then
                    
                        ' ���ǉ��s�̍폜
                        Dim cnt As Long
                        cnt = 0
                        For i = endLineNum To startLineNum Step -1
                            If Right(vbMod.Lines(i, 1), Len("'AddStackTrace")) = "'AddStackTrace" Then
                                Call vbMod.DeleteLines(i)
#If DEBUG_PRINT_MODE Then
                                Debug.Print "DELETE:" & CStr(i)
#End If
                                cnt = cnt + 1
                            End If
                            If (InStr(1, vbMod.Lines(i, 1), "Exit Sub", vbTextCompare) > 0) Or (InStr(1, vbMod.Lines(i, 1), "Exit Function", vbTextCompare) > 0) Then
                                Do While InStr(1, vbMod.Lines(i, 1), "Call PopStackTrace", vbTextCompare) > 0
                                    Dim pos1 As Long
                                    Dim pos2 As Long
                                    pos1 = InStr(1, vbMod.Lines(i, 1), "Call PopStackTrace", vbTextCompare) - 1
                                    pos2 = InStr(pos1, vbMod.Lines(i, 1), ":", vbTextCompare) + 1
                                    Dim repLine As String
                                    repLine = Left(vbMod.Lines(i, 1), pos1) & Mid(vbMod.Lines(i, 1), pos2)
#If DEBUG_PRINT_MODE Then
                                    Debug.Print repLine
#End If
                                    vbMod.ReplaceLine i, repLine
                                Loop
                            End If
                        Next
                        endLineNum = endLineNum - cnt
                        
                        ' Exit Sub/Function�̌���
                        For i = endLineNum To startLineNum Step -1
                            If InStr(1, vbMod.Lines(i, 1), "Exit Sub", vbTextCompare) > 0 Then exitSub.Add i
                            If InStr(1, vbMod.Lines(i, 1), "Exit Function", vbTextCompare) > 0 Then exitSub.Add i
                        Next
                        
                        ' ��̓R�[�h�̍쐬
                        Dim paramValue As String
                        paramValue = GetProcedureArguments(vbMod, procName, procKind)
                        Dim returnValue As String
                        returnValue = ""
                        If InStr(1, vbMod.Lines(defineLineNum, 1), "Function ") > 0 Then
                            returnValue = ", " & procName
                        ElseIf InStr(1, vbMod.Lines(defineLineNum, 1), "Property Get ") > 0 Then
                            returnValue = ", " & procName
                            If Len(paramValue) = 0 Then
                                paramValue = ", ""[Property-Get]"""
                            End If
                        End If
                        Dim insertLinePush As String
                        insertLinePush = _
                            "#If DEBUG_MODE Then 'AddStackTrace" & vbCrLf & _
                            "Call PushStackTrace(""" & modName & """, """ & procName & """" & paramValue & ") 'AddStackTrace" & vbCrLf & _
                            "#End If 'AddStackTrace"
                        Dim insertLinePop As String
                        insertLinePop = _
                            "#If DEBUG_MODE Then 'AddStackTrace" & vbCrLf & _
                            "Call PopStackTrace(""" & modName & """, """ & procName & """" & returnValue & ") 'AddStackTrace" & vbCrLf & _
                            "#End If 'AddStackTrace"
                        Dim insertProcPop As String
                        insertProcPop = _
                            "Call PopStackTrace(""" & modName & """, """ & procName & """" & returnValue & "): "
                        
                        ' ��̓R�[�h�̒ǉ�
                        vbMod.InsertLines endLineNum, insertLinePop
                        For i = 1 To exitSub.Count
                            ' Exit Sub/Function �� 1�s�`��If���Ŏg�p����Ă���ꍇ�� Exit Sub/Function �� insertProcPop : Exit Sub/Function �ɒu��������
                            ' �����łȂ��ꍇ��insertLinePop��ǉ�����
                            If InStr(1, vbMod.Lines(exitSub(i), 1), "If ", vbTextCompare) > 0 Then
                                If InStr(1, vbMod.Lines(exitSub(i), 1), " Then ", vbTextCompare) > 0 Then
                                    If InStr(1, vbMod.Lines(exitSub(i), 1), "Exit Sub", vbTextCompare) > 0 Then
                                        Call vbMod.ReplaceLine(exitSub(i), Replace(vbMod.Lines(exitSub(i), 1), "Exit Sub", insertProcPop & "Exit Sub", 1, -1, vbTextCompare))
                                    End If
                                    If InStr(1, vbMod.Lines(exitSub(i), 1), "Exit Function", vbTextCompare) > 0 Then
                                        Call vbMod.ReplaceLine(exitSub(i), Replace(vbMod.Lines(exitSub(i), 1), "Exit Function", insertProcPop & "Exit Function", 1, -1, vbTextCompare))
                                    End If
                                Else
                                    vbMod.InsertLines exitSub(i), insertLinePop
                                End If
                            ElseIf InStr(1, vbMod.Lines(exitSub(i), 1), "Else:", vbTextCompare) > 0 Then
                                If InStr(1, vbMod.Lines(exitSub(i), 1), "Exit Sub", vbTextCompare) > 0 Then
                                    Call vbMod.ReplaceLine(exitSub(i), Replace(vbMod.Lines(exitSub(i), 1), "Exit Sub", insertProcPop & "Exit Sub", 1, -1, vbTextCompare))
                                End If
                                If InStr(1, vbMod.Lines(exitSub(i), 1), "Exit Function", vbTextCompare) > 0 Then
                                    Call vbMod.ReplaceLine(exitSub(i), Replace(vbMod.Lines(exitSub(i), 1), "Exit Function", insertProcPop & "Exit Function", 1, -1, vbTextCompare))
                                End If
                            Else
                                vbMod.InsertLines exitSub(i), insertLinePop
                            End If
                        Next
                        vbMod.InsertLines startLineNum, insertLinePush
#If DEBUG_PRINT_MODE Then
                        Debug.Print "INSERT[" & CStr(endLineNum) & "]:" & insertLinePop
                        Debug.Print "INSERT[" & CStr(startLineNum) & "]:" & insertLinePush
#End If
                    End If
#If DEBUG_PRINT_MODE Then
                    Debug.Print modName & ":" & CStr(startLineNum); "-"; CStr(endLineNum) & " " & procName
#End If
                    lineNum = vbMod.procStartLine(procName, procKind) + vbMod.ProcCountLines(procName, procKind) + 1
                    numLines = vbMod.CountOfLines
                    Set exitSub = Nothing
                Else
                    Exit Do
                End If
            Loop
        End If
    Next
End Sub

Sub RemoveStackTrace()
    Dim vbEnv As Object
    Dim vbComp As Variant
    Dim vbMod As Object
    Dim modName As String
    Dim procName As String
    Dim lineNum As Long
    Dim numLines As Long
    Dim defineLineNum As Long
    Dim startLineNum As Long
    Dim endLineNum As Long
    Dim procKind As Long
    Dim exitSub As Collection
    Dim shCodeName As String
    shCodeName = "Sheet_DebugTrace"
    ' VBE�̎Q��
    Set vbEnv = Application.VBE
    ' ���ׂĂ�VBComponents�𒀎���������
    For Each vbComp In vbEnv.ActiveVBProject.VBComponents
        ' Module�����擾
        Set vbMod = vbComp.CodeModule
        modName = vbMod.Name
        Dim doAddST As Boolean
        doAddST = True
        If modName = "StackTrace" Then doAddST = False
        If modName = "StackTraceLog" Then doAddST = False
        If modName = shCodeName Then doAddST = False
        If doAddST Then
            ' �s�����擾
            lineNum = vbMod.CountOfDeclarationLines + 1
            numLines = vbMod.CountOfLines
            ' ���ׂẴv���V�[�W���𒲍�
            Do While lineNum < numLines
                ' ���̃v���V�[�W���̎擾
                procKind = 0
                procName = vbMod.ProcOfLine(lineNum, procKind)
                If procName <> "" Then
                    ' �v���V�[�W����`�̊J�n�s�̎擾
                    lineNum = vbMod.ProcBodyLine(procName, procKind)
                    defineLineNum = lineNum
                    Do While Right(vbMod.Lines(lineNum, 1), 1) = "_"
                        lineNum = lineNum + 1
                    Loop
                    startLineNum = lineNum + 1
                    
                    ' �v���V�[�W����`�̏I���s�̎擾
                    lineNum = vbMod.procStartLine(procName, procKind) + vbMod.ProcCountLines(procName, procKind)
                    endLineNum = 0
                    Do While lineNum >= startLineNum
                        If Left(vbMod.Lines(lineNum, 1), 3) = "End" Then
                            endLineNum = lineNum
                            Exit Do
                        End If
                        lineNum = lineNum - 1
                    Loop
                    
                    
                    ' �}���ۃ`�F�b�N�i��`�`End�܂ł�1�s�ŏ�����Ă���C���^�[�t�F�C�X�N���X�̉��z���\�b�h�ɂ͑}������̂��ʓ|�Ȃ̂ő}�����Ȃ����߁B�j
                    If endLineNum > startLineNum Then
                        Dim i As Long
                        Dim cnt As Long
                        cnt = 0
                        For i = endLineNum To startLineNum Step -1
                            If Right(vbMod.Lines(i, 1), Len("'AddStackTrace")) = "'AddStackTrace" Then
                                Call vbMod.DeleteLines(i)
#If DEBUG_PRINT_MODE Then
                                Debug.Print "DELETE:" & CStr(i)
#End If
                                cnt = cnt + 1
                            End If
                            If (InStr(1, vbMod.Lines(i, 1), "Exit Sub", vbTextCompare) > 0) Or (InStr(1, vbMod.Lines(i, 1), "Exit Function", vbTextCompare) > 0) Then
                                Do While InStr(1, vbMod.Lines(i, 1), "Call PopStackTrace", vbTextCompare) > 0
                                    Dim pos1 As Long
                                    Dim pos2 As Long
                                    pos1 = InStr(1, vbMod.Lines(i, 1), "Call PopStackTrace", vbTextCompare) - 1
                                    pos2 = InStr(pos1, vbMod.Lines(i, 1), ":", vbTextCompare) + 1
                                    Dim repLine As String
                                    repLine = Left(vbMod.Lines(i, 1), pos1) & Mid(vbMod.Lines(i, 1), pos2)
#If DEBUG_PRINT_MODE Then
                                    Debug.Print repLine
#End If
                                    vbMod.ReplaceLine i, repLine
                                Loop
                            End If
                        Next
                        endLineNum = endLineNum - cnt
                        
                    End If
#If DEBUG_PRINT_MODE Then
                    Debug.Print modName & ":" & CStr(startLineNum); "-"; CStr(endLineNum) & " " & procName
#End If
                    lineNum = vbMod.procStartLine(procName, procKind) + vbMod.ProcCountLines(procName, procKind) + 1
                    numLines = vbMod.CountOfLines
                Else
                    Exit Do
                End If
            Loop
        End If
    Next
End Sub

Private Function GetProcedureArguments(vbMod As Object, procName As String, procKind As Long) As String
    Dim procStartLine As Long
    Dim procdefLineCount As Long
    Dim procDef As String
    Dim Line As String
    Dim i As Long
    Dim startPos As Long, endPos As Long
    Dim result As String
    
    ' �v���V�[�W���̊J�n�s���擾
    procStartLine = vbMod.ProcBodyLine(procName, procKind)
    
    ' �v���V�[�W����`�̍s������
    procdefLineCount = 1
    Do
        Line = vbMod.Lines(procStartLine + procdefLineCount - 1, 1)
        If Right(Line, 1) = "_" Then
            procDef = procDef & " " & Left(Trim(Line), Len(Trim(Line)) - 1) ' ������ _ �̂ݍ폜
        Else
            procDef = procDef & " " & Trim(Line)
        End If
        procdefLineCount = procdefLineCount + 1
    Loop While Right(Line, 1) = "_"
    
    ' �������X�g�̊J�n�ʒu��T��
    procDef = Replace(procDef, "()", "")
    startPos = InStr(1, procDef, "(")
    endPos = InStrRev(procDef, ")")
    
    result = ""
    
    If startPos > 0 And endPos > startPos Then
        procDef = Mid(procDef, startPos + 1, endPos - startPos - 1)
        
        ' �������X�g�𕪊�
        Dim args As Variant
        args = Split(procDef, ",")
        For i = LBound(args) To UBound(args)
            Dim arg As String
            arg = Trim(args(i))
            
            ' �s�v�ȏC���q���폜
            arg = Replace(arg, "ByRef ", "")
            arg = Replace(arg, "ByVal ", "")
            arg = Replace(arg, "Optional ", "")
            arg = Replace(arg, "ParamArray ", "")
            
            ' �^����f�t�H���g�l���폜
            Dim argParts As Variant
            argParts = Split(arg, " As ")
            argParts = Split(argParts(0), "=")
            
            ' �ŏ��̗v�f���������Ƃ݂Ȃ�
            If Len(Trim(argParts(0))) > 0 Then
                Dim paramName As String
                paramName = Trim(argParts(0))
                result = result & ", """ & paramName & """, " & paramName
            End If
        Next i
    End If
    
    GetProcedureArguments = result
End Function

