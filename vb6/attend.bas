Attribute VB_Name = "attend"
Option Explicit

Public StrSql                    As String
Global GStrSysDate               As String
Global GStrYYYYMM                As String
Global GStrYY                    As String
Global GIntDeptCnt               As Integer

Public ResultHwnd                As Long

Public GStrDept(20)               As String
Public GStrDeptName(20)           As String
Public GStrGradeG(20)           As String

Public GStrUserID                As String
Public GStrGradeInsa             As String
Public GStrUserName              As String
Public GStrGrade                 As String * 1
Public GStrPassChk               As String * 1
Public adoDual                   As ADODB.Recordset
Public adoWork                   As ADODB.Recordset

'/ 각종 SPREAD 화면에서 쓰이는 취소플래그 FORECOLOR 색깔
Public Const gvarCancelColor1  As Variant = 0        '원래색깔
Public Const gvarCancelColor2  As Variant = 233      'SPREAD CANCEL색깔
Public Const varColorSunDay    As Variant = 16777201 '일요일
Public Const varColorDefault   As Variant = 16777215          '원래색깔
Public Const gvarSpreadRow     As Long = 2000        '평상시 SPREAD DISPLAY ROW

'/ 각종 출력폼에서 쓰이는 공통된 결재BOX ┳┫━┃┏┓┛┗┛┣┳┫┻╋
Public Const gsSignForm1 As String = "┏━━━━┳━━━━┳━━━━┳━━━━┳━━━━┓"
Public Const gsSignForm2 As String = "┃담    당┃파 트 장┃팀    장┃행정부장┃의료원장┃"
Public Const gsSignForm3 As String = "┣━━━━╋━━━━╋━━━━╋━━━━╋━━━━┫"
Public Const gsSignForm4 As String = "┃        ┃        ┃        ┃ (전결) ┃        ┃"
Public Const gsSignForm5 As String = "┃        ┃        ┃        ┃        ┃        ┃"
Public Const gsSignForm6 As String = "┗━━━━┻━━━━┻━━━━┻━━━━┻━━━━┛"

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Public Const IME_CMODE_NATIVE = &H1
Public Const IME_CMODE_HANGEUL = IME_CMODE_NATIVE
Public Const IME_CMODE_ALPHANUMERIC = &H0
Public Const IME_SMODE_NONE = &H0

Public Declare Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As Long) As Long
Public Declare Function ImmSetConversionStatus Lib "imm32.dll" (ByVal hIMC As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long

Global GstrPostCode1    As String
Global GstrPostCode2    As String
Global GstrAddress      As String

Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Type PointAPI
     X As Long
     Y As Long
End Type

Global ReturnPos           As PointAPI
Global GstrJikwi           As String


Public Function Date_Format(AnyDateFormat As String) As String
    
    Dim strDate             As String
    
    AnyDateFormat = Trim(AnyDateFormat)
    
    If IsNumeric(AnyDateFormat) Then
        Select Case Len(AnyDateFormat)
            Case 4
                strDate = Left(AnyDateFormat, 2)
                strDate = strDate & "-" & Mid(AnyDateFormat, 3, 2)
            Case 6
                strDate = Left(AnyDateFormat, 2)
                strDate = strDate & "-" & Mid(AnyDateFormat, 3, 2)
                strDate = strDate & "-" & Mid(AnyDateFormat, 5, 2)
            Case 8
                strDate = Left(AnyDateFormat, 4)
                strDate = strDate & "-" & Mid(AnyDateFormat, 5, 2)
                strDate = strDate & "-" & Mid(AnyDateFormat, 7, 2)
        End Select
    Else
        strDate = AnyDateFormat
    End If

    If IsDate(strDate) Then
        Date_Format = Format(strDate, "YYYY-MM-DD")
    Else
        MsgBox "날짜를 잘못입력하셨습니다.", vbCritical, "날짜입력오류"
        Date_Format = ""
    End If

End Function

Public Sub Nrs_Dept(Agr As Object)
    
    Dim AdoDept       As ADODB.Recordset
        
    StrSql = ""
    StrSql = StrSql & "  SELECT B.DEPT, B.DEPTNAME  "
    StrSql = StrSql & "  FROM   TWINSA_WORKMEMBER A, TWINSA_DEPT B"
    StrSql = StrSql & "  WHERE  A.DEPT = B.DEPT"
    StrSql = StrSql & "  and  (b.code1 = '500000' OR A.DEPT = '312610') "
    StrSql = StrSql & "  AND   (DELMARK IS NULL OR DELMARK <> '*')"
    StrSql = StrSql & "  GROUP  BY B.DEPT, B.DEPTNAME  "
    StrSql = StrSql & "  ORDER  BY B.DEPTNAME"

    If adoSetOpen(StrSql, AdoDept) = False Then Exit Sub

    Do Until AdoDept.EOF
        Agr.AddItem AdoDept.Fields("DEPTNAME").Value & Space(25) & AdoDept.Fields("DEPT").Value
        AdoDept.MoveNext
    Loop

    AdoDept.Close
    Set AdoDept = Nothing
    
End Sub

Public Sub Data_Dept(Agr As Object)
    
    Dim AdoDept       As ADODB.Recordset
        
    StrSql = ""
    StrSql = StrSql & "  SELECT B.DEPT, B.DEPTNAME  "
    StrSql = StrSql & "  FROM   TWINSA_WORKMEMBER A, TWINSA_DEPT B"
    StrSql = StrSql & "  WHERE  A.DEPT = B.DEPT"
    StrSql = StrSql & "  AND   (DELMARK IS NULL OR DELMARK <> '*')"
    StrSql = StrSql & "  GROUP  BY B.DEPT, B.DEPTNAME  "
    StrSql = StrSql & "  ORDER  BY B.DEPTNAME"

    If adoSetOpen(StrSql, AdoDept) = False Then Exit Sub

    Do Until AdoDept.EOF
        Agr.AddItem AdoDept.Fields("DEPTNAME").Value & Space(25) & AdoDept.Fields("DEPT").Value
        AdoDept.MoveNext
    Loop

    AdoDept.Close
    Set AdoDept = Nothing
    
End Sub

Public Sub Data_Part(Agr As Object, StrDept As String)
Dim AdoDept       As ADODB.Recordset
Dim AdoPart       As ADODB.Recordset
    
    Agr.Clear
    
    StrSql = ""
    StrSql = StrSql & "  SELECT DEPTNAME, DEPT "
    StrSql = StrSql & "  FROM   TWINSA_DEPT"
    StrSql = StrSql & "  WHERE  ( CODE2 = '" & StrDept & "'"
    StrSql = StrSql & "         or  (CODE1 = '300000' and CODE3 = '" & StrDept & "')) "
    StrSql = StrSql & "  AND    (DELMARK IS NULL OR DELMARK <> '*')"
    StrSql = StrSql & "  ORDER  BY DEPTNAME"

    If adoSetOpen(StrSql, AdoDept) = False Then Exit Sub

    Do Until AdoDept.EOF
        
'2017-09-07 이전데이타 조회가 안됨
'        StrSql = ""
'        StrSql = StrSql & "   SELECT * FROM TWINSA_MASTER"
'        StrSql = StrSql & "   WHERE  DEPT = '" & AdoDept.Fields("DEPT").Value & "'"
'        StrSql = StrSql & "   AND    STATUS <> '3'  "
'
'        If adoSetOpen(StrSql, AdoPart) = True Then
'            Agr.AddItem AdoDept.Fields("DEPTNAME").Value & Space(25) & AdoDept.Fields("DEPT").Value
'        End If
'
'        AdoPart.Close
'        Set AdoPart = Nothing
'
        Agr.AddItem AdoDept.Fields("DEPTNAME").Value & Space(25) & AdoDept.Fields("DEPT").Value
        
        AdoDept.MoveNext
    Loop

    If Agr.ListCount > 0 Then Agr.ListIndex = 0

    AdoDept.Close
    Set AdoDept = Nothing
    
End Sub
Public Sub Data_Dept_NRS(Agr As Object)
Dim AdoDept       As ADODB.Recordset
        
    StrSql = ""
    StrSql = StrSql & "  SELECT DEPTCODE, DEPTNAMEK "
    StrSql = StrSql & "  FROM   TWBAS_DEPT  "
    StrSql = StrSql & "  ORDER  BY PRINTRANKING "

    If adoSetOpen(StrSql, AdoDept) = False Then Exit Sub

    Do Until AdoDept.EOF
        Agr.AddItem AdoDept.Fields("DEPTCODE").Value ' & Space(25) & AdoDept.Fields("DEPTNAMEK").Value
        AdoDept.MoveNext
    Loop

End Sub
Public Sub Data_Bun_Setting(Agr As Object, IntRow As Integer, StrGrade As String)
Dim AdoBun             As ADODB.Recordset
Dim strTemp            As String

    StrSql = ""
    StrSql = StrSql & "  SELECT BUN, NAME "
    StrSql = StrSql & "  FROM   TWNRS_BUN"
        
    Select Case StrGrade
        Case "H":    StrSql = StrSql & "  WHERE  GUBUN IN (1, 2)"
        Case "N", "A":   StrSql = StrSql & "  WHERE  GUBUN IN (0, 1, 2)"
        Case "D":   StrSql = StrSql & "  WHERE  GUBUN IN (1, 2, 3)"
        Case "H2":   StrSql = StrSql & "  WHERE  bun in ('off', 'A1') "
    End Select
    
    StrSql = StrSql & "  AND    GBUSE = '1' "
    StrSql = StrSql & "  ORDER BY NAME"
    
    If adoSetOpen(StrSql, AdoBun) = False Then Exit Sub
    
    Agr.BlockMode = True
    Agr.Col = IntRow:    Agr.Col2 = IntRow
    Agr.Row = -1
    Agr.CellType = CellTypeComboBox
    Agr.TypeMaxEditLen = 10
    Agr.TypeHAlign = TypeHAlignLeft
    Agr.TypeVAlign = TypeVAlignCenter
        
    strTemp = ""
    
    Do Until AdoBun.EOF
        strTemp = strTemp & Chr$(9) & AdoBun.Fields("NAME").Value & Space(20) & AdoBun.Fields("BUN").Value
        AdoBun.MoveNext
    Loop
    
    Agr.TypeComboBoxList = strTemp
    Agr.BlockMode = False

End Sub

Public Function RPadH(ByVal strString As String, ByVal lngLength As Long) As String

    RPadH = LeftH(strString & Space(lngLength), lngLength)

End Function
Public Function LeftH(ByVal strString As String, ByVal lngLength As Long) As String

    LeftH = StrConv(LeftB(StrConv(strString, vbFromUnicode), lngLength), vbUnicode)

End Function
'필드의 알파벳 리턴
Public Function AlphaCount(acDeg As Long) As String
    Dim acH As Long
    Dim acF As Long
    Dim Tmp As String
    If acDeg > 26 Then
        acH = acDeg / 26
        acF = acDeg Mod 26
        Tmp = Chr(acH + 64) & Chr(acF + 64)
    Else
        Tmp = Chr(acDeg + 64)
    End If
    AlphaCount = Tmp
    
    
End Function

Public Function SpreadToExcell(StrTitle As String, ByVal sSpr As Object, Optional sPathFile As String = "") As Integer

    On Error GoTo ErrMsg
    
    Dim xlApp       As Excel.Application    ''이렇게 선언하구요..
    Dim xlBook      As Excel.Workbook
    Dim xlSheet     As Excel.Worksheet
    
    Dim oData() ' As String
    Dim ri          As Long
    Dim ci          As Long
    Dim rMax        As Long
    Dim cMax        As Long
    Dim iCol        As Long
    Dim fLine       As Long
    Dim colS        As String
    
    
    fLine = 0
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add()
    Set xlSheet = xlBook.Worksheets(1)
    xlSheet.Activate
    xlSheet.Name = StrTitle
    
    If sSpr.DisplayColHeaders = True Then fLine = 1     'Header Header 가 True 일때
    For iCol = 1 To sSpr.MaxCols
        xlSheet.Cells(iCol).ColumnWidth = sSpr.ColWidth(iCol)
    Next
    rMax = sSpr.MaxRows + fLine
    cMax = sSpr.MaxCols
    
    ReDim oData(1 To rMax, 1 To cMax)
    
    For ri = 1 To rMax
        For ci = 1 To cMax
            sSpr.Row = ri - fLine   'Header Line 을 뺀다...
            sSpr.Col = ci
            oData(ri, ci) = sSpr.text
        Next
    Next
    
'    '필드의 알파벳 리턴
    colS = AlphaCount(cMax)

    xlSheet.Range("A1", colS & CStr(rMax)).Font.Name = "굴림체"
    xlSheet.Range("A1", colS & CStr(rMax)).Font.Size = 9
    xlSheet.Range("A1", colS & CStr(rMax)).VerticalAlignment = 2      'Center
    xlSheet.Range("A1", colS & CStr(rMax)).HorizontalAlignment = 3    'Center
    xlSheet.Range("A1", colS & CStr(rMax)).Value = oData
'    xlSheet.Range("A1", colS & CStr(rMax)).Borders.LineStyle = 1
    Erase oData
    
   'Header 만 Font를 변화준다........
'    For iCol = 1 To sSpr.DataColCnt
'        xlSheet.Cells(1, iCol).Font.Bold = True
'        xlSheet.Cells(1, iCol).Font.Size = 11
'        xlSheet.Cells(1, iCol).VerticalAlignment = 2   'Center
'        xlSheet.Cells(1, iCol).HorizontalAlignment = 3 'Center
'    Next
    
    
    '파일저장이냐 보내기냐
    If sPathFile = "" Then                              'Path 와 File 의 Argument 가 없으면... eXcell 을 Display
        xlApp.Visible = True                            '화면Display(Excell)
    Else                                                'Path 와 File 의 Argument 가 있으면 저장
        If Dir(sPathFile) <> "" Then Kill sPathFile
        xlSheet.SaveAs sPathFile  '저장
        xlBook.Close
        xlApp.Quit
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set xlApp = Nothing
    End If
    Exit Function



ErrMsg:
    MsgBox Error$, vbOKOnly


End Function

Public Function Get_DeptName(argDeptCode As String) As String   '부서(진료과)명 가져오기
    
    Dim RsTemp  As ADODB.Recordset
    Dim strTemp As String
    
    strTemp = "          SELECT DEPTNAMEK FROM TWBAS_DEPT "
    strTemp = strTemp & " WHERE  DEPTCODE = '" & argDeptCode & "' "
    
    If adoSetOpen(strTemp, RsTemp) Then
        If RsTemp.RecordCount > 0 Then
            Get_DeptName = Trim(RsTemp.Fields("DEPTNAMEK").Value & "")
        Else
            Get_DeptName = ""
        End If
        RsTemp.Close
        Set RsTemp = Nothing
    End If
    
End Function


Public Sub HanOn(Src As Object)

'// 오브젝트의 입력모드를 한글로 고정

  Dim hIME As Long

  hIME = ImmGetContext(Src.hwnd)

  ImmSetConversionStatus hIME, IME_CMODE_HANGEUL, IME_SMODE_NONE

End Sub

Public Sub Form_Clear(ByVal frm As Object, Optional ByVal CtlStart, Optional ByVal CtlEnd)
    
    Dim Ctl             As Variant
    Dim nMinTabIndex    As Integer
    Dim nMaxTabIndex    As Integer
    
    On Error Resume Next        ' TabIndex...
    
    If IsMissing(CtlStart) Then nMinTabIndex = 0 Else nMinTabIndex = CtlStart.TabIndex
    If IsMissing(CtlEnd) Then nMaxTabIndex = 32767 Else nMaxTabIndex = CtlEnd.TabIndex
    
    For Each Ctl In frm.Controls
        If Ctl.TabIndex >= nMinTabIndex And Ctl.TabIndex <= nMaxTabIndex Then
            If TypeOf Ctl Is TextBox Then
                Ctl.text = ""
                Ctl.Enabled = True
            ElseIf TypeOf Ctl Is ListBox Then
                Ctl.Clear
            ElseIf TypeOf Ctl Is ComboBox Then
                If Ctl.Tag = "" Then
                    If Ctl.Style = vbComboDropdownList Then
                        Ctl.ListIndex = -1
                    Else
                        Ctl.text = ""
                    End If
                    Ctl.Enabled = True
                End If
            ElseIf TypeOf Ctl Is CheckBox Then
                Ctl.Value = vbUnchecked
            ElseIf TypeOf Ctl Is OptionButton Then
                If Ctl.Tag = "" Then Ctl.Value = False
                Ctl.Enabled = True
            ElseIf TypeOf Ctl Is vaSpread Then
                Call Spread_Clear(Ctl)
                Ctl.Enabled = True
            ElseIf TypeOf Ctl Is Label Then
                If Ctl.BorderStyle = vbFixedSingle Then
                    Ctl.Caption = ""
                End If
                Ctl.Enabled = True
            ElseIf TypeOf Ctl Is SSPanel Then
                If Ctl.Tag <> "" Then Ctl.Caption = ""
            ElseIf TypeOf Ctl Is Frame Then
                If Ctl.Name = "fraMenu" Then Ctl.Enabled = False
            End If
        End If
    Next Ctl
End Sub


Public Sub Spread_Clear(ByVal ArgSpread As vaSpread)
    If ArgSpread.Tag <> "" Then Exit Sub
    With ArgSpread
        .BlockMode = True
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .ForeColor = gvarCancelColor1
        .Action = ActionClearText
        .BackColor = varColorDefault
        .BlockMode = False
        If .MaxRows > 0 And .MaxRows > 0 Then
            .Col = 1: .Row = 1: .Action = ActionActiveCell
        End If
        .MaxRows = gvarSpreadRow
    End With
    
End Sub
Public Function Get_GradeName(argGradeCode As String) As String     '등급명 가져오기
    
    Dim RsTemp  As ADODB.Recordset
    Dim strTemp As String
    
    strTemp = "          SELECT NAME FROM TWNRS_GRADE "
    strTemp = strTemp & " WHERE  GRADE = '" & argGradeCode & "' "
    strTemp = strTemp & "   AND  GBUSE = '1'"
    
    If adoSetOpen(strTemp, RsTemp) Then
        
        If RsTemp.RecordCount > 0 Then
            Get_GradeName = Trim(RsTemp("NAME") & "")
        Else
            Get_GradeName = ""
        End If
        RsTemp.Close
        Set RsTemp = Nothing
    End If
    
End Function

Public Function Get_BunName(argBunCode As String) As String       '분류명 가져오기
    
    Dim RsTemp  As ADODB.Recordset
    Dim strTemp As String
    
    strTemp = "          SELECT NAME FROM TWNRS_BUN "
    strTemp = strTemp & "  WHERE BUN = '" & argBunCode & "' "
    
    If adoSetOpen(strTemp, RsTemp) Then
        If RsTemp.RecordCount > 0 Then
            Get_BunName = RsTemp("NAME") & ""
        Else
            Get_BunName = ""
        End If
        RsTemp.Close
        Set RsTemp = Nothing
    End If
    
End Function

'반복되는 컬럼을 그룹화 시킨다
Public Sub Set_Spread_Group(ByRef ArgSpread As vaSpread, Argcol As Long, Optional argCol1 As Long = 0, Optional ArgCol2 As Long = 0)
    Dim strBackUp   As String
    Dim i           As Integer
    Dim j           As Integer
    
    strBackUp = ""
    With ArgSpread
        If argCol1 = 0 Then
            For i = 1 To .MaxRows
                .Row = i: .Col = Argcol
                If .text = strBackUp Then
                    strBackUp = .text
                    .text = ""
                Else
                    strBackUp = .text
                End If
            Next i
        Else
            For i = 1 To .MaxRows
                .Row = i: .Col = Argcol
                If .text = strBackUp Then
                    strBackUp = .text
                    For j = argCol1 To ArgCol2
                        .Col = j
                        .text = ""
                    Next j
                Else
                    strBackUp = .text
                End If
            Next i
        End If
    End With
    
End Sub

Public Sub Spread_Sort(ByVal SS As Object, ByVal nSortKey1 As Integer, Optional ByVal nSortKey2 As Variant, Optional ByVal nSortKey3 As Variant)

    If IsMissing(nSortKey2) Then nSortKey2 = nSortKey1
    If IsMissing(nSortKey3) Then nSortKey3 = nSortKey2

    SS.Row = 1: SS.Row2 = SS.DataRowCnt
    SS.Col = 1: SS.Col2 = SS.DataColCnt
    SS.SortBy = SS_SORT_BY_ROW
    SS.SortKey(1) = Abs(nSortKey1)
    SS.SortKeyOrder(1) = IIf(nSortKey1 > 0, SS_SORT_ORDER_ASCENDING, SS_SORT_ORDER_DESCENDING)
    SS.SortKey(2) = Abs(nSortKey2)
    SS.SortKeyOrder(2) = IIf(nSortKey2 > 0, SS_SORT_ORDER_ASCENDING, SS_SORT_ORDER_DESCENDING)
    SS.SortKey(3) = Abs(nSortKey3)
    SS.SortKeyOrder(3) = IIf(nSortKey3 > 0, SS_SORT_ORDER_ASCENDING, SS_SORT_ORDER_DESCENDING)
    SS.Action = SS_ACTION_SORT

End Sub


Public Sub cvtToHan(ByRef ArgObject As Object)
   Dim hIMC                 As Long
   
   hIMC = ImmGetContext(ArgObject.hwnd)
   ImmSetConversionStatus hIMC, IME_CMODE_HANGEUL, IME_SMODE_NONE
   
End Sub

Public Function Get_HealthCheck(argGubun As String, ArgSabun As String, ArgYear As String) As String       '일반검진,특수검진 내역 가져오기
    
    Dim RsTemp  As ADODB.Recordset
    Dim strTemp As String
    
    strTemp = "  "
    strTemp = strTemp & " SELECT NVL2 ( "
    strTemp = strTemp & "           MAX (jubsfdate), "
    strTemp = strTemp & "              SUBSTR (MAX (jubsfdate), 3, 2) "
    strTemp = strTemp & "           || '/' "
    strTemp = strTemp & "           || SUBSTR (MAX (jubsfdate), 5, 2) "
    strTemp = strTemp & "           || '/' "
    strTemp = strTemp & "           || SUBSTR (MAX (jubsfdate), 7, 2), "
    strTemp = strTemp & "           NULL) "
    strTemp = strTemp & "           jubsfdate "
    strTemp = strTemp & "   FROM twinsa_master A, "
    strTemp = strTemp & "        twbas_jumin b, "
    strTemp = strTemp & "        kuh_me.jubsutable c, "
    strTemp = strTemp & "        kuh_me.jubgktable d "
    strTemp = strTemp & "  WHERE     A.status <> '3' "
    strTemp = strTemp & "        AND c.jubsfdate >= '" & ArgYear & "0101' "
    strTemp = strTemp & "        AND c.jubsfdate <= '" & ArgYear & "1231' "
    strTemp = strTemp & "        AND c.jubsjubgn = 'A' "
    strTemp = strTemp & "        AND c.jubshplce = '11' "
    strTemp = strTemp & "        AND NVL (c.jubsptcno, '') <> ' ' "
    strTemp = strTemp & "        AND A.jumin = b.jumin "
    strTemp = strTemp & "        AND b.ptno = c.jubsptcno "
    strTemp = strTemp & "        AND c.jubsjubno = d.jubgjubno "
    strTemp = strTemp & "        AND A.sabun = '" & ArgSabun & "' "
    If argGubun = "G" Then
        strTemp = strTemp & "  AND d.jubggumkd <> 'S1' "
    Else
        strTemp = strTemp & "  and d.jubggumkd = 'S1' "
    End If
    
    If adoSetOpen(strTemp, RsTemp) Then
        If RsTemp.RecordCount > 0 Then
            Get_HealthCheck = RsTemp("jubsfdate") & ""
        Else
            Get_HealthCheck = ""
        End If
        RsTemp.Close
        Set RsTemp = Nothing
    End If
    
End Function


Public Function DrPlan_Save(StrWDate As String, strDrCode As String, strbun As String, strRemark As String) As String


    DrPlan_Save = "Y"
    
    'DRPLAN 삭제 시작---------------------------------------------------------------------------------

        StrSql = ""
        StrSql = StrSql & "DELETE FROM twbas_drplan "
        StrSql = StrSql & "      WHERE drcode = '" & strDrCode & "'  "
        StrSql = StrSql & "     AND bdate = TO_DATE('" & StrWDate & "','YYYY-MM-DD')"
        If adoExecute(StrSql) = False Then DrPlan_Save = "N"


        StrSql = ""
        StrSql = StrSql & "DELETE FROM twbas_drplan_status "
        StrSql = StrSql & "      WHERE drcode = '" & strDrCode & "'  "
        StrSql = StrSql & "     AND entdate = TO_DATE('" & StrWDate & "','YYYY-MM-DD')"
        If adoExecute(StrSql) = False Then DrPlan_Save = "N"


'DRPLAN 삭제 끝---------------------------------------------------------------------------------


'DRPLAN 입력 시작---------------------------------------------------------------------------------
    
        If Trim(strbun) = "A3" Or Trim(strbun) = "A3A" Or Trim(strbun) = "A3P" Or Trim(strbun) = "L" Or Trim(strbun) = "Hi" Or Trim(strbun) = "Hx" Or Trim(strbun) = "Ho" Or Trim(strbun) = "HoH" Or Trim(strbun) = "H" Or Trim(strbun) = "A" Or _
           Trim(strbun) = "off" Or Trim(strbun) = "H6" Or Trim(strbun) = "H7" Or Trim(strbun) = "H2" Or _
           Trim(strbun) = "A11" Or Trim(strbun) = "A12" Or Trim(strbun) = "A21" Or Trim(strbun) = "A22" Or _
           Trim(strbun) = "A31" Or Trim(strbun) = "A32" Or Trim(strbun) = "A63" Or Trim(strbun) = "A64" Or Trim(strbun) = "A68" Or Trim(strbun) = "H3" Or _
           Trim(strbun) = "H3A" Or Trim(strbun) = "H3P" Or _
           Trim(strbun) = "Hc" Or Trim(strbun) = "H8" Or Trim(strbun) = "H5" Or Trim(strbun) = "Hb" Or Trim(strbun) = "HaP" Or _
           Trim(strbun) = "HhA" Or Trim(strbun) = "HhP" Or Trim(strbun) = "HiA" Or Trim(strbun) = "HiP" Or Trim(strbun) = "HE" Or Trim(strbun) = "A5" Or Trim(strbun) = "Hn" Or Trim(strbun) = "Dp" Then
        
            StrSql = ""
            StrSql = StrSql & "INSERT INTO twbas_drplan "
            StrSql = StrSql & "            (drcode, bdate, drstatus, jdate, part "
            StrSql = StrSql & "            ) "
            StrSql = StrSql & "     VALUES ('" & strDrCode & "', TO_DATE('" & StrWDate & "','YYYY-MM-DD'), "
            Select Case Trim(strbun)
                Case "A3":    StrSql = StrSql & "            '41', "
                Case "A3A":   StrSql = StrSql & "            '42', "
                Case "A3P":   StrSql = StrSql & "            '43', "
                Case "L":     StrSql = StrSql & "            '51', "
                Case "Hi", "HaP":    StrSql = StrSql & "            '61', "
                Case "Ho":    StrSql = StrSql & "            '62', "
                Case "H", "Hx", "H6", "H7", "H3", "H2", "H5", "Hb", "Hn", "HoH": StrSql = StrSql & "            '65', " 'Hx추가 2023-07-26
                Case "H3A":   StrSql = StrSql & "            '11', "
                Case "H3P":   StrSql = StrSql & "            '21', "
                Case "A":     StrSql = StrSql & "            '66', "
                Case "off":   StrSql = StrSql & "            '31', "
            
                Case "A11":   StrSql = StrSql & "            '11', "
                Case "A12":   StrSql = StrSql & "            '12', "
                Case "A21":   StrSql = StrSql & "            '21', "
                Case "A22":   StrSql = StrSql & "            '22', "
                Case "A31":   StrSql = StrSql & "            '31', "
                Case "A32":   StrSql = StrSql & "            '32', "                                ''당직(비번)을 당일마감으로 넘김
                
                Case "A63":   StrSql = StrSql & "            '63', "
                Case "A64":   StrSql = StrSql & "            '64', "
                Case "A68":   StrSql = StrSql & "            '68', "

                Case "Hc", "H8":  StrSql = StrSql & "            '69', "

                Case "HhA":   StrSql = StrSql & "            '52', "
                Case "HhP":   StrSql = StrSql & "            '53', "
                Case "HiA":   StrSql = StrSql & "            '54', "
                Case "HiP":   StrSql = StrSql & "            '55', "
                Case "HE":    StrSql = StrSql & "            '67', "
                
                Case "A5":    StrSql = StrSql & "            '70', "             '당직
                Case "Dp":    StrSql = StrSql & "            '82', "             '파견 -> 출장

'                Case "H3":    strSql = strSql & "            '67', "
            
            End Select
            
            StrSql = StrSql & "            SYSDATE, '" & GstrPassSabun & "'"
            StrSql = StrSql & "            ) "
        
        
            If adoExecute(StrSql) = False Then DrPlan_Save = "N"
        
            If Trim(strbun) = "A11" Or Trim(strbun) = "A12" Or Trim(strbun) = "A21" Or Trim(strbun) = "A22" Or _
               Trim(strbun) = "A31" Or Trim(strbun) = "A32" Or Trim(strbun) = "A63" Or Trim(strbun) = "A64" Or Trim(strbun) = "A68" Then
        
                StrSql = ""
                StrSql = StrSql & "INSERT INTO twbas_drplan_status "
                StrSql = StrSql & "            (drcode, entdate, drstatus, remark1 "
                StrSql = StrSql & "            ) "
                StrSql = StrSql & "     VALUES ('" & strDrCode & "', TO_DATE('" & StrWDate & "','YYYY-MM-DD'),"
                
                Select Case Trim(strbun)
                    Case "A11":   StrSql = StrSql & "            '11', "
                    Case "A12":   StrSql = StrSql & "            '12', "
                    Case "A21":   StrSql = StrSql & "            '21', "
                    Case "A22":   StrSql = StrSql & "            '22', "
                    Case "A31":   StrSql = StrSql & "            '31', "
                    Case "A32":   StrSql = StrSql & "            '32', "
                    Case "A63":   StrSql = StrSql & "            '63', "
                    Case "A64":   StrSql = StrSql & "            '64', "
                    Case "A68":   StrSql = StrSql & "            '68', "
                    Case "A5":    StrSql = StrSql & "            '70', "
                End Select
                
                StrSql = StrSql & "             '" & strRemark & "' "
                StrSql = StrSql & "            ) "
 
                If adoExecute(StrSql) = False Then DrPlan_Save = "N"
 
            End If
        
        End If

'DRPLAN 입력 끝 -------------------------------------------------------------------------------
    
End Function

Function PayRound(ArgNum As Double) As Double
    
    Dim sNum As String
    
    sNum = Int(CStr(ArgNum))
    
    If InStr(1, "56789", Right(sNum, 1)) > 0 Then
        ArgNum = (Int(sNum / 10) + 1) * 10
    Else
        ArgNum = Int(sNum / 10) * 10
    End If
    
    PayRound = ArgNum

End Function
