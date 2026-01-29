VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.1#0"; "codejock.controls.v13.1.0.ocx"
Begin VB.Form FrmAttend1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "근태체크"
   ClientHeight    =   13830
   ClientLeft      =   2670
   ClientTop       =   1605
   ClientWidth     =   16815
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13830
   ScaleWidth      =   16815
   WindowState     =   2  '최대화
   Begin TabDlg.SSTab SSTab1 
      Height          =   5370
      Left            =   225
      TabIndex        =   33
      Top             =   8325
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   9472
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   -2147483628
      TabCaption(0)   =   "주 근무시간"
      TabPicture(0)   =   "FrmAttend1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "SSWeek"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "단시간"
      TabPicture(1)   =   "FrmAttend1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ss2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "btnDetail"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin XtremeSuiteControls.PushButton btnDetail 
         Height          =   690
         Left            =   14355
         TabIndex        =   36
         Top             =   450
         Width           =   645
         _Version        =   851969
         _ExtentX        =   1138
         _ExtentY        =   1217
         _StockProps     =   79
         Caption         =   "근태대장"
         UseVisualStyle  =   -1  'True
      End
      Begin FPSpread.vaSpread SSWeek 
         Height          =   4875
         Left            =   -74775
         TabIndex        =   34
         Top             =   405
         Width           =   14010
         _Version        =   393216
         _ExtentX        =   24712
         _ExtentY        =   8599
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   2
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   27
         MaxRows         =   50
         RowsFrozen      =   3
         ScrollBars      =   2
         SpreadDesigner  =   "FrmAttend1.frx":0038
      End
      Begin FPSpreadADO.fpSpread ss2 
         Height          =   4770
         Left            =   90
         TabIndex        =   35
         Top             =   450
         Width           =   14190
         _Version        =   393216
         _ExtentX        =   25030
         _ExtentY        =   8414
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   16
         MaxRows         =   20
         ScrollBars      =   2
         ShadowColor     =   8417376
         ShadowText      =   16777215
         SpreadDesigner  =   "FrmAttend1.frx":18F9
      End
   End
   Begin VB.CheckBox chkWeek 
      Caption         =   "주 근무시간 조회"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6075
      TabIndex        =   32
      Top             =   540
      Width           =   1680
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1740
      Left            =   225
      TabIndex        =   7
      Top             =   6570
      Width           =   15180
      _Version        =   65536
      _ExtentX        =   26776
      _ExtentY        =   3069
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      FloodColor      =   255
      Alignment       =   6
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "※ 코로나19 의심 증상은 발현 시 반드시 기입 후 감염 관리팀 연락 부탁드립니다."
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   135
         TabIndex        =   31
         Top             =   1215
         Width           =   6540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "(의심증상:기침, 가래, 인후통, 호흡곤란, 콧물, 급성설사, 근육통, 권태감, 후각 저하 등)"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   330
         TabIndex        =   30
         Top             =   1455
         Width           =   7050
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "단, 증상이 없다면 기입하지 않으셔도 됩니다."
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   7515
         TabIndex        =   29
         Top             =   1440
         Width           =   3660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "▷평일 근무: 반근무자는 정상근무 대신 Half 근무로 체크해주십시오."
         ForeColor       =   &H000040C0&
         Height          =   180
         Index           =   3
         Left            =   4590
         TabIndex        =   19
         Top             =   930
         Width           =   5535
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "▷토요일 근무:  비근무자는 비번으로 체크해주십시오."
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   18
         Top             =   930
         Width           =   4380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   $"FrmAttend1.frx":2120
         ForeColor       =   &H00008000&
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   11415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "출장은 출장내역을 기입해 주십시오."
         ForeColor       =   &H00008000&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   630
         Width           =   2940
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "▶ 다음의 사항은 사유를 반드시 입력해야합니다."
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   135
         TabIndex        =   8
         Top             =   90
         Width           =   3960
      End
   End
   Begin VB.CheckBox chkTime 
      BackColor       =   &H00C0FFFF&
      Caption         =   "출근시간세팅(정상근무자 중 출근카드 미태그자를 현재시간으로 설정)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   450
      TabIndex        =   27
      Top             =   7440
      Visible         =   0   'False
      Width           =   6540
   End
   Begin VB.ComboBox CmbPart 
      Height          =   300
      Left            =   4080
      Style           =   2  '드롭다운 목록
      TabIndex        =   17
      Top             =   510
      Width           =   1935
   End
   Begin FPSpreadADO.fpSpread SS 
      Height          =   5625
      Left            =   225
      TabIndex        =   14
      Top             =   930
      Width           =   15180
      _Version        =   393216
      _ExtentX        =   26776
      _ExtentY        =   9922
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   17
      MaxRows         =   20
      ScrollBars      =   2
      ShadowColor     =   8417376
      ShadowText      =   16777215
      SpreadDesigner  =   "FrmAttend1.frx":21AD
   End
   Begin Threed.SSPanel SSCHK 
      Height          =   315
      Left            =   3090
      TabIndex        =   11
      Top             =   120
      Width           =   1995
      _Version        =   65536
      _ExtentX        =   3519
      _ExtentY        =   556
      _StockProps     =   15
      BackColor       =   16775408
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.01
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   1
      Outline         =   -1  'True
      Begin VB.CheckBox ChkOK 
         BackColor       =   &H00FFF8F0&
         Caption         =   "완료"
         Enabled         =   0   'False
         Height          =   225
         Left            =   1020
         TabIndex        =   13
         Top             =   30
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFF8F0&
         Caption         =   "출근체크"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.ComboBox CmbDept 
      Height          =   300
      Left            =   1200
      Style           =   2  '드롭다운 목록
      TabIndex        =   4
      Top             =   510
      Width           =   1845
   End
   Begin MSComCtl2.DTPicker DTDate 
      Height          =   315
      Left            =   1215
      TabIndex        =   1
      Top             =   120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   556
      _Version        =   393216
      Format          =   149684225
      CurrentDate     =   37540
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   315
      Index           =   0
      Left            =   210
      TabIndex        =   2
      Top             =   120
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "날 짜"
      ForeColor       =   8388608
      BackColor       =   16769216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.01
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   315
      Index           =   1
      Left            =   210
      TabIndex        =   3
      Top             =   510
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "부 서"
      ForeColor       =   8388608
      BackColor       =   16769216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.01
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   315
      Index           =   2
      Left            =   3090
      TabIndex        =   16
      Top             =   510
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "파 트"
      ForeColor       =   8388608
      BackColor       =   16769216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.01
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   435
      Left            =   180
      TabIndex        =   20
      Top             =   6675
      Visible         =   0   'False
      Width           =   10995
      _Version        =   65536
      _ExtentX        =   19394
      _ExtentY        =   767
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Outline         =   -1  'True
      Begin Threed.SSPanel SSPanel4 
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   21
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   397
         _StockProps     =   15
         BackColor       =   10502304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   225
         Index           =   1
         Left            =   3630
         TabIndex        =   22
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   397
         _StockProps     =   15
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   225
         Index           =   2
         Left            =   5520
         TabIndex        =   23
         Top             =   90
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   397
         _StockProps     =   15
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "지각"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   3990
         TabIndex        =   26
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "출근체크기에 출근체크 안함"
         ForeColor       =   &H00A040A0&
         Height          =   180
         Index           =   0
         Left            =   450
         TabIndex        =   25
         Top             =   120
         Width           =   2280
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "정상근무"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   5880
         TabIndex        =   24
         Top             =   120
         Width           =   720
      End
   End
   Begin Threed.SSCommand CmdAdd 
      Height          =   825
      Left            =   9555
      TabIndex        =   10
      Top             =   60
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1455
      _StockProps     =   78
      Caption         =   "추 가"
      ForeColor       =   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmAttend1.frx":2AA7
   End
   Begin Threed.SSCommand CmdSearch 
      Height          =   825
      Left            =   7920
      TabIndex        =   0
      Top             =   60
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1455
      _StockProps     =   78
      Caption         =   "조  회"
      ForeColor       =   16576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmAttend1.frx":2DC1
   End
   Begin Threed.SSCommand CmdExit 
      Height          =   825
      Left            =   10365
      TabIndex        =   5
      Top             =   60
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1455
      _StockProps     =   78
      Caption         =   "닫 기"
      ForeColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmAttend1.frx":30DB
   End
   Begin Threed.SSCommand CmdSave 
      Height          =   825
      Left            =   8730
      TabIndex        =   6
      Top             =   60
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1455
      _StockProps     =   78
      Caption         =   "저  장"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmAttend1.frx":33F5
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ctrl + F : 검색"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   11340
      TabIndex        =   28
      Top             =   675
      Width           =   1140
   End
End
Attribute VB_Name = "FrmAttend1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrDept                  As String
Dim i                        As Integer
Dim AdoDual                  As ADODB.Recordset
Dim AdoWork                  As ADODB.Recordset
Dim strYYYYMM                As String
Dim StrDay                   As String
Dim StrSysDate               As String
Dim StrMonDay                As String        '월요일 가져오기
Dim StrWeek                  As String        '요일가져오기


Private Sub btnDetail_Click()
    Dim strDeptCode         As String

    If CmbPart.ListCount > 0 Then
        StrDept = Right(CmbPart.text, 6)
    Else
        StrDept = Right(CmbDept.text, 6)
    End If
    
    strDeptCode = StrDept
    
    '진료운영팀 통합관리 : 신경과, 뇌파검사실, 소화기능검사실, 폐기능검사실, 청력검사실, 전정기능검사실, 언어치료실, 수면다원검사실,근전도검사실,신경심리검사실
    If Right(CmbDept.text, 6) = "313000" And CmbPart.text = "ALL" Then
        strDeptCode = "313000','320402','320407','320800','320401','320405','320406', '320412', '320404', '320403', '310300"
    End If
    
    '원무팀은 파트별로 관리 가능하도록 함. 2022-09 -> 파트를 선택해도 입력은 원무팀 코드로
    If Right(CmbDept.text, 6) = "620100" Or Right(CmbDept.text, 6) = "620200" Then
        strDeptCode = Right(CmbDept.text, 6)
    End If
    
    Call frmAttend24.Form_Init(strDeptCode)
End Sub

Private Sub chkTime_Click()
    Dim i As Integer

    If chkTime.Value = 1 Then
        For i = 1 To SS.MaxRows
            SS.Row = i
            SS.Col = 7
            If Left(SS.text, 4) = "정상근무" Then
                SS.Col = 6
                If SS.text = "" Then
                    SS.text = Get_ChkTime
                End If
            End If
        Next i
    Else
        For i = 1 To SS.MaxRows
            SS.Row = i
            SS.Col = 7
            If Left(SS.text, 4) = "정상근무" Then
                SS.Col = 6
                If SS.ForeColor = RGB(160, 64, 160) Then
                    SS.text = ""
                End If
            End If
        Next i
    End If
    
End Sub

Private Function Get_ChkTime() As String

Dim AdoTime            As ADODB.Recordset
    
    strSql = "SELECT TO_CHAR(SYSDATE, 'HH24:MI') TIME1  FROM  DUAL"
        
    If adoSetOpen(strSql, AdoTime, 1) Then
            Get_ChkTime = AdoTime.Fields("TIME1").Value & ""
    End If
End Function

Private Sub CmbDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub
Private Sub CmbDept_LostFocus()
    StrDept = Right(CmbDept.text, 6)
    Call Data_Part(CmbPart, StrDept)
    
    '진료운영팀 통합관리 : 신경과, 뇌파검사실, 소화기능검사실, 폐기능검사실, 청력검사실, 전정기능검사실, 언어치료실
    '원무팀 파트별 관리
    If StrDept = "313000" Or StrDept = "620100" Or StrDept = "620200" Then
        CmbPart.AddItem "ALL"
    End If
    
    If CmbPart.ListCount > 0 Then
        CmbPart.Visible = True:  SSPanel1(2).Visible = True
    Else
        CmbPart.Visible = False: SSPanel1(2).Visible = False
    End If

End Sub

'Private Sub CmbPart_GotFocus()
'    Call Clear_Spread
'End Sub

Private Sub CmbPart_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub
Private Sub CmdAdd_Click()
Dim StrSabun              As String
Dim adoSabun              As ADODB.Recordset
Dim strYYYYMM             As String
Dim strDeptCode         As String
    
    strDeptCode = StrDept

    '원무팀은 파트별로 관리 가능하도록 함. 2022-09 -> 파트를 선택해도 입력은 원무팀 코드로
    If Right(CmbDept.text, 6) = "620100" Or Right(CmbDept.text, 6) = "620200" Then
        strDeptCode = Right(CmbDept.text, 6)
    End If

    strYYYYMM = Format(DTDate.Value, "yyyymm")

    StrSabun = InputBox("추가하려는 교직원번호를 입력하세요", "교직원추가")
    If Len(StrSabun) <> 6 Then MsgBox "교직원번호 입력이 잘못되었습니다.", vbCritical, "교직원번호 오류":   Exit Sub
    
    strSql = ""
    strSql = strSql & " SELECT NAMEK, A.DEPT, B.DEPTNAME "
    strSql = strSql & " FROM   TWINSA_MASTER A, (select * from TWINSA_DEPT where delmark is null) B"
    strSql = strSql & " WHERE  SABUN  = '" & StrSabun & "'"
    strSql = strSql & " AND    A.DEPT = B.DEPT"
    strSql = strSql & "  AND (   status = 0 "
    strSql = strSql & "        OR (status = '2' AND TO_CHAR (statusdate, 'yyyymm') >= '" & strYYYYMM & "') "
    strSql = strSql & "       ) "
    strSql = strSql & "   AND    SUBSTR(a.SABUN,1,1) NOT IN ('2','4')"

        
    If adoSetOpen(strSql, adoSabun, 1) = False Then
        MsgBox "존재하지 않는 교직원번호입니다.", vbCritical, "교직원번호 오류"
        Exit Sub
    Else
       
       If GStrGrade <> 0 And strDeptCode <> adoSabun.Fields("DEPT").Value & "" Then
            MsgBox adoSabun.Fields("NAMEK").Value & "는 현재 " & adoSabun.Fields("DEPTNAME").Value & " 소속입니다. 추가할 수 없습니다.", vbCritical, "입력오류"
            Exit Sub
     '관리자는 현재부서가 달라도 추가할수있도록 함 2024-09-24
       ElseIf GStrGrade = 0 And strDeptCode <> adoSabun.Fields("DEPT").Value & "" Then
            If vbNo = MsgBox(adoSabun.Fields("NAMEK").Value & "는 현재 " & adoSabun.Fields("DEPTNAME").Value & _
                        " 소속입니다. 그래도 추가하시겠습니까?", vbYesNo + vbDefaultButton2 + vbQuestion, "소속확인") Then
                Exit Sub
            End If
       End If
       
'            strSql = ""
'            strSql = strSql & "  INSERT INTO  TWINSA_WORKDAILY (YYYYMM, SABUN, DEPTCODE)"
'            strSql = strSql & "  SELECT '" & strYYYYMM & "' YYYYMM, SABUN, DEPT  "
'            strSql = strSql & "  FROM   TWINSA_MASTER"
'            strSql = strSql & "  WHERE  DEPT    = '" & strDeptCode & "'"
'            strSql = strSql & "  AND    SABUN   = '" & StrSabun & "'"
        
        strSql = ""
        strSql = strSql & "  INSERT INTO  TWINSA_WORKDAILY (YYYYMM, SABUN, DEPTCODE)"
        strSql = strSql & "  values ('" & strYYYYMM & "', '" & StrSabun & "', '" & strDeptCode & "')"
                    
        adoConnect.BeginTrans
        
        If adoExecute(strSql) Then
            adoConnect.CommitTrans
            Call Data_Search
        Else
            adoConnect.RollbackTrans
            MsgBox "등록에 문제가 있습니다. 전산개발팀 연락요망", vbCritical, "Error"
            Exit Sub
        End If
        
    
    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
Dim strChk                As String * 1
Dim StrSabun              As String * 6
Dim strbun                As String * 3
Dim strDeptCode           As String
Dim StrName               As String
Dim strRemark             As String
Dim StrRemark2             As String
Dim StrRemark3             As String
Dim StrStrDate            As String
Dim StrEndDate            As String
Dim AdoMemo               As ADODB.Recordset
Dim AdoTime               As ADODB.Recordset
Dim AdoCnt                As ADODB.Recordset
Dim StrDD                 As String
Dim strMM                 As String
Dim StrChkDate            As String
Dim j                     As Integer
Dim StrOverTime           As String
Dim strTime               As String

strChk = "Y"
    
'    If GStrGrade <> 0 Then     '관리자는 전부서
'        If Save_Time_Chk = "N" Then CmdSave.Enabled = False: Exit Sub
'    End If
    
    If GStrGrade <> 0 Then     '관리자는 전부서
        If StrSysDate < GStrSysDate Then             '오늘날짜와 조회일자가 다르면 저장 불가능
            CmdSave.Enabled = False
        ElseIf StrSysDate > GStrSysDate Then
            CmdSave.Enabled = True
        Else
            CmdSave.Enabled = True
            If Save_Time_Chk = "N" Then CmdSave.Enabled = False
        End If
    End If
    
    
    For i = 1 To SS.MaxRows
                            
        SS.Row = i
        SS.Col = 1: strDeptCode = SS.text
        SS.Col = 3: StrSabun = SS.text
        SS.Col = 4: StrName = SS.text
        SS.Col = 7: strbun = Trim(Right(SS.text, 3))
        SS.Col = 10: strRemark = Trim(SS.text)
        
        '2025-03-24 (Hs 반차 + Half 퇴근) 제외 인사팀 조인선 확인
        If (Trim(strbun) = "A" Or Trim(strbun) = "A3" Or Trim(strbun) = "A3A" Or Trim(strbun) = "A3P" Or Trim(strbun) = "A4" Or Trim(strbun) = "H2" Or Trim(strbun) = "H3" Or Trim(strbun) = "H3A" Or Trim(strbun) = "H3P" Or _
            Trim(strbun) = "H6" Or Trim(strbun) = "H7" Or Trim(strbun) = "H9" Or Trim(strbun) = "HL" Or Trim(strbun) = "L" Or Trim(strbun) = "Hd") _
        And strRemark = "" Then
            MsgBox "[" + StrSabun + " " + StrName + "] 사유를 입력하세요", vbCritical, "입력오류"
            SS.Row = i:   SS.Col = 8: SS.Action = ActionActiveCell
            Exit Sub
        End If
        
        If (Trim(strbun) = "H1" Or Trim(strbun) = "H0") Then
            strSql = ""
            strSql = strSql & " SELECT COUNT(*) CNT FROM TWINSA_WORKMEMO "
            strSql = strSql & " WHERE  SABUN = '" & StrSabun & "'"
            strSql = strSql & " AND    WORKDATE >= TO_DATE('" & Format(StrSysDate, "YYYY-MM-01") & "','YYYY-MM-DD')"
            strSql = strSql & " AND    WORKDATE <  TO_DATE('" & StrSysDate & "','YYYY-MM-DD')"
            strSql = strSql & " AND    BUN   = '" & strbun & "'"
                                                        
            If adoSetOpen(strSql, AdoCnt, 1) = True Then
                If AdoCnt.Fields("CNT").Value > 0 Then
                    MsgBox StrName & "는 이미 이번달에 같은 휴가(월차, 보건휴가)를 사용하였습니다.", vbCritical, "입력오류"
                    SS.Row = i:   SS.Col = 7: SS.Action = ActionActiveCell
                    Exit Sub
                End If
            End If
                    
        End If
                    
'        If Trim(strBun) = "Sh" Then
'            For j = 0 To 6
'                StrChkDate = DateAdd("d", j, StrMonDay)
'                StrDD = "DAY" & Format(StrChkDate, "DD")
'                StrMM = Format(StrChkDate, "YYYYMM")
'
'                strSql = ""
'                strSql = strSql & " SELECT *                           "
'                strSql = strSql & " FROM   TWINSA_WORKDAILY            "
'                strSql = strSql & " WHERE  YYYYMM = '" & StrMM & "'    "
'                strSql = strSql & " AND    SABUN  = '" & strSabun & "' "
'                strSql = strSql & " AND " & StrDD & "  = 'Sh'          "
'
'                If adoSetOpen(strSql, AdoCnt, 1) = True Then
'                    If AdoCnt.RecordCount = 1 Then
'                        MsgBox strName & "는 이미 이번주에 Half근무를 하였습니다.", vbCritical, "입력오류"
'                        SS.Row = i:   SS.Col = 4: SS.Action = ActionActiveCell
'                    End If
'                End If
'            Next j
        
'        End If
        
    Next i
    
'-----------------------------------------------------------------------------
    
    For i = 1 To SS.MaxRows
        
        SS.Row = i
        SS.Col = 1: strDeptCode = SS.text
        SS.Col = 3: StrSabun = SS.text
        SS.Col = 7: strbun = Trim(Right(SS.text, 3))
        SS.Col = 8: StrStrDate = Trim(SS.text)
        SS.Col = 9: StrEndDate = Trim(SS.text)
        SS.Col = 10: strRemark = Trim(SS.text)
        SS.Col = 11: StrRemark2 = Trim(SS.text)
        SS.Col = 12: StrRemark3 = Trim(SS.text)
        SS.Col = 6: strTime = Trim(SS.text)


'2012-11-28
'정상근무와 Half근무를 정상근무로 하고 오버Time 가져가자.
'지각만 오버Time 가져가기
'        If Trim(strBun) = "A1" Or Trim(strBun) = "Sh" Then
'지각 시간 계산하기 막음 2019-05-15
'        StrOverTime = ""
'
'        If Trim(StrBun) = "HL" Then
'            If strTime = "" Then
'                StrOverTime = Format(Time, "hh:mm")
'            Else
'                StrOverTime = Format(strTime, "hh:mm")
'            End If
'            StrOverTime = DateDiff("n", "08:30", StrOverTime)
'            If StrOverTime > 0 Then
'                StrOverTime = Format(Fix(StrOverTime / 60), "00") & ":" & Format(StrOverTime Mod 60, "00")
'            Else
'                StrOverTime = ""
'            End If
'        End If


'2017-09-08 혹시 몰라서 막는다 길은희
''연차면 출근체크시간 입력안되도록 인사팀 이은희 2013-01-31
'        If Trim(StrBun) = "H" Then
'            strTime = ""
'            StrOverTime = ""
'        End If
        
        
        strSql = ""
        strSql = strSql & " UPDATE TWINSA_WORKDAILY "
        strSql = strSql & " SET    " & StrDay & " = '" & strbun & "'"
        strSql = strSql & " WHERE  SABUN    = '" & StrSabun & "'"
        strSql = strSql & " AND    YYYYMM   = '" & strYYYYMM & "'"
        strSql = strSql & " AND    DEPTCODE = '" & strDeptCode & "'"
        
        If adoExecute(strSql) = False Then strChk = "N"
    
    
'MEMO 삭제 시작---------------------------------------------------------------------------
        
        strSql = ""
        strSql = strSql & " SELECT *  FROM TWINSA_WORKMEMO "
        strSql = strSql & " WHERE  WORKDATE = TO_DATE('" & StrSysDate & "','YYYY-MM-DD')"
        'StrSql = StrSql & " AND    DEPTCODE = '" & strDeptCode & "'"
        strSql = strSql & " AND    SABUN    = '" & StrSabun & "'            "
        
        If adoSetOpen(strSql, AdoMemo) = True Then
            
            strSql = ""
            strSql = strSql & " DELETE FROM TWINSA_WORKMEMO "
            strSql = strSql & " WHERE  WORKDATE = TO_DATE('" & StrSysDate & "','YYYY-MM-DD')"
            'StrSql = StrSql & " AND    DEPTCODE = '" & strDeptCode & "'"
            strSql = strSql & " AND    SABUN    = '" & StrSabun & "'      "
            
            If adoExecute(strSql) = False Then strChk = "N"
        
        End If
        
'MEMO 삭제 끝-----------------------------------------------------------------------------
'MEMO 입력 시작---------------------------------------------------------------------------
        If Trim(strRemark) <> "" Or Trim(StrRemark2) <> "" Or Trim(StrRemark3) <> "" Or Trim(StrStrDate) <> "" Or (Left(strbun, 1) = "H" And strbun <> "off" And strbun <> "A1") Then
        
            strSql = ""
            strSql = strSql & " INSERT INTO TWINSA_WORKMEMO (WORKDATE, SABUN, DEPTCODE,BUN, REMARK, REMARK2, REMARK3, STARTDATE,ENDDATE)"
            strSql = strSql & " VALUES (TO_DATE('" & StrSysDate & "','YYYY-MM-DD') "
            strSql = strSql & ",         '" & StrSabun & "'"
            strSql = strSql & ",         '" & strDeptCode & "' "
            strSql = strSql & ",         '" & strbun & "' "
            strSql = strSql & ",         '" & strRemark & "'     "
            strSql = strSql & ",         '" & StrRemark2 & "'     "
            strSql = strSql & ",         '" & StrRemark3 & "'     "
            strSql = strSql & ",        TO_DATE('" & StrStrDate & "','YYYY-MM-DD') "
            strSql = strSql & ",        TO_DATE('" & StrEndDate & "','YYYY-MM-DD')) "
            If adoExecute(strSql) = False Then strChk = "N"
        
        End If

'MEMO 입력 끝-----------------------------------------------------------------------------
'Time 입력 시작 -----------------------------------------------------------------------------
    
        strSql = ""
        strSql = strSql & " SELECT *  FROM TWINSA_WORKTIME "
        strSql = strSql & " WHERE  WORKDATE = ? "
        strSql = strSql & " AND    SABUN    = ? "
        
        Call adoCmd(strSql, 1)
        With adoCommand
            .Parameters.Append .CreateParameter("workdate", adDate, adParamInput, 10, Format(StrSysDate, "yyyy-mm-dd"))
            .Parameters.Append .CreateParameter("sabun", adChar, adParamInput, 6, StrSabun)
        End With
        
        Set AdoMemo = adoCommand.Execute
        If AdoMemo.RecordCount > 0 Then
            
            strSql = ""
            strSql = strSql & " DELETE FROM TWINSA_WORKTIME "
            strSql = strSql & " WHERE  WORKDATE = ? "
            strSql = strSql & " AND    SABUN    = ? "
            
            Call adoCmd(strSql, 1)
            With adoCommand
                .Parameters.Append .CreateParameter("workdate", adDate, adParamInput, 10, Format(StrSysDate, "yyyy-mm-dd"))
                .Parameters.Append .CreateParameter("sabun", adChar, adParamInput, 6, StrSabun)
            End With
                    
            Set AdoTime = adoCommand.Execute
        
        End If
    
        If strTime <> "" Or StrOverTime <> "" Then
        
            strSql = ""
            strSql = strSql & " INSERT INTO TWINSA_WORKTIME (WORKDATE, "
            strSql = strSql & "                             SABUN, "
            strSql = strSql & "                             WORKTIME, "
            strSql = strSql & "                             OVERTIME) "
            strSql = strSql & "     VALUES (?,"
            strSql = strSql & "             ?, "
            strSql = strSql & "             ?, "
            strSql = strSql & "             ?) "
        
            Call adoCmd(strSql, 1)
            With adoCommand
                .Parameters.Append .CreateParameter("workdate", adDate, adParamInput, 10, Format(StrSysDate, "yyyy-mm-dd"))
                .Parameters.Append .CreateParameter("sabun", adChar, adParamInput, 6, StrSabun)
                .Parameters.Append .CreateParameter("WORKTIME", adChar, adParamInput, 5, strTime)
                .Parameters.Append .CreateParameter("OVERTIME", adChar, adParamInput, 5, StrOverTime)
            End With
                    
            Set AdoTime = adoCommand.Execute
        End If
    
'Time 입력 끝 -----------------------------------------------------------------------------
    
    
    Next i
    
'------------------------단시간 근무자 근태체크 별도 구성 2024-03-07'-------------------
'-----------------------------------------------------------------------------
    Dim StrRemark1           As String
    Dim StrTime1             As String
    Dim StrTime2             As String
    Dim StrTime3             As String
    
    For i = 1 To ss2.MaxRows
        
        ss2.Row = i
        ss2.Col = 1: strDeptCode = ss2.text
        ss2.Col = 3: StrSabun = ss2.text
        ss2.Col = 5: strbun = Trim(Right(ss2.text, 3))
        ss2.Col = 6: StrTime1 = LeftH(ss2.text & Space(5), 5)
        ss2.Col = 7: StrTime2 = LeftH(ss2.text & Space(5), 5)
        ss2.Col = 8: StrTime3 = LeftH(ss2.text & Space(5), 5)
        StrRemark1 = "s" & StrTime1 & "e" & StrTime2 & "t" & StrTime3
        ss2.Col = 9: strRemark = Trim(ss2.text)
        ss2.Col = 10: StrRemark2 = Trim(ss2.text)
        ss2.Col = 11: StrRemark3 = Trim(ss2.text)
        
        strSql = ""
        strSql = strSql & " UPDATE TWINSA_WORKDAILY "
        strSql = strSql & " SET    " & StrDay & " = '" & strbun & "'"
        strSql = strSql & " WHERE  SABUN    = '" & StrSabun & "'"
        strSql = strSql & " AND    YYYYMM   = '" & strYYYYMM & "'"
        strSql = strSql & " AND    DEPTCODE = '" & strDeptCode & "'"
        
        If adoExecute(strSql) = False Then strChk = "N"
    
    
'MEMO 삭제 시작---------------------------------------------------------------------------
        
        strSql = ""
        strSql = strSql & " SELECT *  FROM TWINSA_WORKMEMO "
        strSql = strSql & " WHERE  WORKDATE = TO_DATE('" & StrSysDate & "','YYYY-MM-DD')"
        'StrSql = StrSql & " AND    DEPTCODE = '" & strDeptCode & "'"
        strSql = strSql & " AND    SABUN    = '" & StrSabun & "'            "
        
        If adoSetOpen(strSql, AdoMemo) = True Then
            
            strSql = ""
            strSql = strSql & " DELETE FROM TWINSA_WORKMEMO "
            strSql = strSql & " WHERE  WORKDATE = TO_DATE('" & StrSysDate & "','YYYY-MM-DD')"
            'StrSql = StrSql & " AND    DEPTCODE = '" & strDeptCode & "'"
            strSql = strSql & " AND    SABUN    = '" & StrSabun & "'      "
            
            If adoExecute(strSql) = False Then strChk = "N"
        
        End If
        
'MEMO 삭제 끝-----------------------------------------------------------------------------
'MEMO 입력 시작---------------------------------------------------------------------------
        If Trim(strRemark) <> "" Or Trim(StrRemark1) <> "" Or Trim(StrRemark2) <> "" Or Trim(StrRemark3) <> "" Or Trim(StrStrDate) <> "" Or (Left(strbun, 1) = "H" And strbun <> "off" And strbun <> "A1") Then
        
            strSql = ""
            strSql = strSql & " INSERT INTO TWINSA_WORKMEMO (WORKDATE, SABUN, DEPTCODE,BUN, REMARK, REMARK1, REMARK2, REMARK3, STARTDATE,ENDDATE)"
            strSql = strSql & " VALUES (TO_DATE('" & StrSysDate & "','YYYY-MM-DD') "
            strSql = strSql & ",         '" & StrSabun & "'"
            strSql = strSql & ",         '" & strDeptCode & "' "
            strSql = strSql & ",         '" & strbun & "' "
            strSql = strSql & ",         '" & strRemark & "'     "
            strSql = strSql & ",         '" & StrRemark1 & "'     "
            strSql = strSql & ",         '" & StrRemark2 & "'     "
            strSql = strSql & ",         '" & StrRemark3 & "'     "
            strSql = strSql & ",        TO_DATE('" & StrStrDate & "','YYYY-MM-DD') "
            strSql = strSql & ",        TO_DATE('" & StrEndDate & "','YYYY-MM-DD')) "
            If adoExecute(strSql) = False Then strChk = "N"
        
        End If

'MEMO 입력 끝-----------------------------------------------------------------------------

    Next i
'-------------------단시간 근태 입력 끝'-------------------
    
    adoConnect.BeginTrans
    
    If strChk = "Y" Then
        adoConnect.CommitTrans
        MsgBox StrSysDate & "일자에 근태를 체크하셨습니다.", vbInformation, "출근체크완료"
    Else
        adoConnect.RollbackTrans
    End If
    
    Call Clear_Spread
    Call Data_Search
    Call Data_Search2
    
End Sub

Private Sub cmdSearch_Click()
    chkTime.Value = 0
    Call Clear_Spread
    If DTDate.Value <> GStrSysDate Then If Save_Time_Chk = "N" Then CmdSave.Enabled = False
    
    Call Data_Search
    Call Data_Search2
    
    
'    Call Data_Search_Week

''2019-05-24 근무시간조회(부장, 팀장, 팀장대리, 관리자) , (시설팀, 영상의학과 파트장), (진료운영팀 500442김남순, 601177전병준)
'    If GstrJikwi = "1006" Or GstrJikwi = "3001" Or GstrJikwi = "3002" Or GStrGrade = 0 Or _
'        (GstrJikwi = "3003" And GStrDept(1) = "650100") Or _
'        (GstrJikwi = "3003" And GStrDept(1) = "311800") Or _
'        (GstrJikwi = "3003" And GStrDept(1) = "210400") Or _
'        (GStrDept(1) = "313000" And GstrIdnumber = "500442") Or _
'        (GStrDept(1) = "313000" And GstrIdnumber = "601177") Then
'
'        If chkWeek.Value = 0 Then Exit Sub
'
'        SSWeek.Visible = True
'        Call Data_Search_Week
'    End If

        If chkWeek.Value = 0 Then Exit Sub
        Call Data_Search_Week
End Sub

Private Sub DTDate_Change()
'    If GStrSysDate = Format(DTDate.Value, "yyyy-mm-dd") Then
'        chkTime.Visible = True
'    Else
'        chkTime.Visible = False
'    End If
End Sub

Private Sub DTDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{TAB}"

End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()

    Dim i             As Integer
    Dim StrSabun              As String * 6
    Dim AdoMemo               As ADODB.Recordset
    Dim AdoTime               As ADODB.Recordset
    Dim StrOverTime           As String
    Dim strTime               As String

    Call Clear_Spread
    Call Data_Bun_Setting(SS, 7, "H")
    Call Data_Bun_Setting(ss2, 5, "H2")
    
'    If GstrIdnumber <> "602121" Then
'        btnDetail.Visible = False
'    End If
    
    'SSWeek.Visible = False
    
    DTDate.Value = Format(GStrSysDate, "YYYY-MM-DD")
    
    If GStrGrade = 0 Then     '관리자는 전부서
        Call Data_Dept(CmbDept)
        CmbDept.Enabled = True
        'SSCHK.Visible = True
        ChkOK.Enabled = False
    Else
        'SSCHK.Visible = False
        
        For i = 1 To GIntDeptCnt
            CmbDept.AddItem GStrDeptName(i) & Space(25) & GStrDept(i)
        Next i
    
        '특수검사실 파트장 강지수, 검사실 통합관리
        If GstrIdnumber = "300150" Then
            CmbDept.AddItem "뇌파검사실" & Space(25) & "320401"
            CmbDept.AddItem "소화기능검사실" & Space(25) & "320402"
            CmbDept.AddItem "언어치료실" & Space(25) & "320800"
            CmbDept.AddItem "전정기능검사실" & Space(25) & "320407"
            CmbDept.AddItem "청력검사실" & Space(25) & "320406"
            CmbDept.AddItem "폐기능검사실" & Space(25) & "320405"
        End If
    
    
    End If
    
    For i = 0 To CmbDept.ListCount
        CmbDept.ListIndex = i
        If Right(CmbDept.text, 6) = GStrDept(1) Then CmbDept.ListIndex = i: Call CmbDept_LostFocus: Exit For
    Next i

    Call Data_Search
    
    SSTab1.Tab = 1
    
    Call Data_Search2   '단시간근로자
    
''2019-05-24 근무시간조회(팀장, 관리자,,, 시설팀/영상의학팀은 파트장까지)
'    If GstrJikwi = "1006" Or GstrJikwi = "3001" Or GstrJikwi = "3002" Or GStrGrade = 0 Or _
'        (GstrJikwi = "3003" And GStrDept(1) = "650100") Or _
'        (GstrJikwi = "3003" And GStrDept(1) = "311800") Or _
'        (GstrJikwi = "3003" And GStrDept(1) = "210400") Or _
'        (GStrDept(1) = "313000" And GstrIdnumber = "500442") Or _
'        (GStrDept(1) = "313000" And GstrIdnumber = "601177") Then
'        chkWeek.Visible = True
'        SSWeek.Visible = True
'        'Call Data_Search_Week
'    End If
    
    'Time 입력 시작-----------------------------------------------------------------------------------
    '당일 조회시 근태체크시간 일단 저장...! 저장안누르는 경우가 있어서..2014-07-16 류연주 추가
    If StrSysDate <> GStrSysDate Or Save_Time_Chk <> "Y" Then Exit Sub
    'NU 영양팀의 경우 자동 저장되도록 수정 2015-01-26 : 수당 계산에 사용하므로 중요함.
    
    For i = 1 To SS.MaxRows
        
        SS.Row = i
        SS.Col = 3: StrSabun = SS.text
        SS.Col = 6: strTime = Trim(SS.text)

        strSql = ""
        strSql = strSql & " SELECT *  FROM TWINSA_WORKTIME "
        strSql = strSql & " WHERE  WORKDATE = ? "
        strSql = strSql & " AND    SABUN    = ? "
        
        Call adoCmd(strSql, 1)
        With adoCommand
            .Parameters.Append .CreateParameter("workdate", adDate, adParamInput, 10, Format(StrSysDate, "yyyy-mm-dd"))
            .Parameters.Append .CreateParameter("sabun", adChar, adParamInput, 6, StrSabun)
        End With
        
        Set AdoMemo = adoCommand.Execute
        If AdoMemo.RecordCount > 0 Then
            
            strSql = ""
            strSql = strSql & " DELETE FROM TWINSA_WORKTIME "
            strSql = strSql & " WHERE  WORKDATE = ? "
            strSql = strSql & " AND    SABUN    = ? "
            
            Call adoCmd(strSql, 1)
            With adoCommand
                .Parameters.Append .CreateParameter("workdate", adDate, adParamInput, 10, Format(StrSysDate, "yyyy-mm-dd"))
                .Parameters.Append .CreateParameter("sabun", adChar, adParamInput, 6, StrSabun)
            End With
                    
            Set AdoTime = adoCommand.Execute
        
        End If
    
        If strTime <> "" Or StrOverTime <> "" Then
        
            strSql = ""
            strSql = strSql & " INSERT INTO TWINSA_WORKTIME (WORKDATE, "
            strSql = strSql & "                             SABUN, "
            strSql = strSql & "                             WORKTIME, "
            strSql = strSql & "                             OVERTIME) "
            strSql = strSql & "     VALUES (?,"
            strSql = strSql & "             ?, "
            strSql = strSql & "             ?, "
            strSql = strSql & "             ?) "
        
            Call adoCmd(strSql, 1)
            With adoCommand
                .Parameters.Append .CreateParameter("workdate", adDate, adParamInput, 10, Format(StrSysDate, "yyyy-mm-dd"))
                .Parameters.Append .CreateParameter("sabun", adChar, adParamInput, 6, StrSabun)
                .Parameters.Append .CreateParameter("WORKTIME", adChar, adParamInput, 5, strTime)
                .Parameters.Append .CreateParameter("OVERTIME", adChar, adParamInput, 5, StrOverTime)
            End With
                    
            Set AdoTime = adoCommand.Execute
        End If
    
    Next i
    'Time 입력 끝-----------------------------------------------------------------------------------
    
    
End Sub
Private Function Save_Time_Chk() As String

Dim AdoTime            As ADODB.Recordset
    
    
    'MsgBox "아침/점심 직원체온 및 증상을 사유란에 입력바랍니다.", vbCritical, "중요"
    
    '2020-02-07 신종코로나로 인한 직원 체온 체크 위해 근태 입력 시간 제한 안함 남수진 요청
    Save_Time_Chk = "Y"
    Exit Function
    '----------------------------------------------------------------------//
    
    strSql = "SELECT TO_CHAR(SYSDATE, 'HH24:MI') TIME1  FROM  DUAL"
        
    If adoSetOpen(strSql, AdoTime, 1) Then
        
        If Trim(GstrPassDept) = "NU" Then
            Save_Time_Chk = "Y"
            Exit Function
        End If
        
        '2015-03-12 인사팀 고수원 요청으로 근태체크 마감시간 10:30분으로 변경
        'If AdoTime.Fields("TIME1").Value & "" > "09:30" Then
        'If AdoTime.Fields("TIME1").Value & "" > "10:30" Then
'
'        If AdoTime.Fields("TIME1").Value & "" > "11:00" Then
'            MsgBox "1차 근태체크가 마감되었습니다. 조회만 가능합니다.", vbCritical, "확인"
'            Save_Time_Chk = "N"
'        Else
'            Save_Time_Chk = "Y"
'        End If


        If AdoTime.Fields("TIME1").Value & "" > "10:30" And AdoTime.Fields("TIME1").Value & "" < "14:00" Then
            MsgBox "1차 근태체크가 마감되었습니다. 조회만 가능합니다.", vbCritical, "확인"
            Save_Time_Chk = "N"
        ElseIf AdoTime.Fields("TIME1").Value & "" > "15:30" Then
            MsgBox "2차 근태체크가 마감되었습니다. 조회만 가능합니다.", vbCritical, "확인"
            Save_Time_Chk = "N"
        Else
            Save_Time_Chk = "Y"
        End If
        
''        '2020-02-07 신종코로나로 인한 직원 체온 체크 위해 근태 입력 시간 변경
''        If AdoTime.Fields("TIME1").Value & "" > "10:30" And AdoTime.Fields("TIME1").Value & "" < "15:30" Then
''            MsgBox "1차 근태체크가 마감되었습니다. 조회만 가능합니다.", vbCritical, "확인"
''            Save_Time_Chk = "N"
''        ElseIf AdoTime.Fields("TIME1").Value & "" > "17:30" Then
''            MsgBox "2차 근태체크가 마감되었습니다. 조회만 가능합니다.", vbCritical, "확인"
''            Save_Time_Chk = "N"
''        Else
''            Save_Time_Chk = "Y"
''        End If
        
    End If

    If Save_Time_Chk = "Y" Then

        strSql = ""
        strSql = strSql & " SELECT HOLYDAY "
        strSql = strSql & "  FROM TWBAS_JOBDATE "
        strSql = strSql & " WHERE JOBDATE = ? "
        Call adoCmd(strSql, 1)
        With adoCommand
            .Parameters.Append .CreateParameter("JOBDATE", adDate, adParamInput, 10, Format(DTDate.Value, "YYYY-MM-DD"))
        End With
        
        Set AdoTime = adoCommand.Execute
        If AdoTime.RecordCount > 0 Then
            If AdoTime.Fields("HOLYDAY").Value & "" = "*" Then
                MsgBox "공휴일은 근태체크를 하지 않습니다.조회만 가능합니다.", vbCritical, "확인"
                Save_Time_Chk = "N"
            End If
        End If
    End If

    If GStrGrade = 0 Then Save_Time_Chk = "Y"

End Function

Private Sub Clear_Spread()
    
    SS.Row = 1: SS.Row2 = SS.MaxRows
    SS.Col = 1: SS.Col2 = SS.MaxCols
    SS.BlockMode = True
    SS.Lock = False
    SS.text = ""
    SS.BlockMode = False
    
    SSWeek.ClearRange 1, 4, SSWeek.MaxCols, SSWeek.MaxRows, True
    
    ss2.MaxRows = 0
End Sub

Private Sub Data_Search()
Dim strDate             As String
Dim strDeptCode         As String
Dim strDataChk          As String
    
    If CmbPart.ListCount > 0 Then
        StrDept = Right(CmbPart.text, 6)
    Else
        StrDept = Right(CmbDept.text, 6)
    End If
    
    strDeptCode = StrDept
    
    '진료운영팀 통합관리 : 신경과, 뇌파검사실, 소화기능검사실, 폐기능검사실, 청력검사실, 전정기능검사실, 언어치료실, 수면다원검사실,근전도검사실,신경심리검사실
    If Right(CmbDept.text, 6) = "313000" And CmbPart.text = "ALL" Then
        strDeptCode = "313000','320402','320407','320800','320401','320405','320406', '320412', '320404', '320403', '310300"
    End If
    
    '원무팀은 파트별로 관리 가능하도록 함. 2022-09 -> 파트를 선택해도 입력은 원무팀 코드로
    If Right(CmbDept.text, 6) = "620100" Or Right(CmbDept.text, 6) = "620200" Then
        strDeptCode = Right(CmbDept.text, 6)
    End If
    
    ChkOK.Value = 0
    strDate = Format(DTDate.Value, "yyyymmdd")
    
    strSql = ""
    strSql = strSql & " SELECT TO_DATE('" & DTDate.Value & "','YYYY-MM-DD') SysDate1,                  "
    strSql = strSql & "        TO_CHAR(TO_DATE('" & DTDate.Value & "','YYYY-MM-DD'), 'YYYYMM') YYYYMM, "
    strSql = strSql & "        TO_CHAR(TO_DATE('" & DTDate.Value & "','YYYY-MM-DD'), 'DD') DAY,        "
    strSql = strSql & "        TO_CHAR(NEXT_DAY(TO_DATE('" & DTDate.Value & "','YYYY-MM-DD') - 7,2),'YYYY-MM-DD') Mon,    "
    strSql = strSql & "        TO_CHAR(TO_DATE('" & DTDate.Value & "','YYYY-MM-DD'), 'DAY') WEEK       "
    strSql = strSql & " FROM   DUAL                                                                    "
        
    If adoSetOpen(strSql, AdoDual, 1) = True Then
        StrSysDate = AdoDual.Fields("SysDate1").Value & ""
        strYYYYMM = AdoDual.Fields("YYYYMM").Value & ""
        StrDay = AdoDual.Fields("DAY").Value & ""
        StrMonDay = AdoDual.Fields("MON").Value & ""
        StrWeek = Trim(AdoDual.Fields("WEEK").Value & "")
    End If

'날짜 및 시간 체크-----------------------------------------------------------------------
    
    If GStrGrade <> 0 Then     '관리자는 전부서
        If StrSysDate < GStrSysDate Then             '오늘날짜와 조회일자가 다르면 저장 불가능
            CmdSave.Enabled = False
            
'            If StrSysDate = DateAdd("D", 1, GStrSysDate) And StrWeek = "SATURDAY" Then
'                CmdSave.Enabled = True
'                If Save_Time_Chk = "N" Then CmdSave.Enabled = False
'            Else
'                CmdSave.Enabled = False
'            End If
        ElseIf StrSysDate > GStrSysDate Then
            CmdSave.Enabled = True
        Else
            CmdSave.Enabled = True
            If Save_Time_Chk = "N" Then CmdSave.Enabled = False
        End If
    End If

'-----------------------------------------------------------------------------------
    strDataChk = ""
    strSql = ""
    strSql = strSql & "  SELECT * FROM TWINSA_WORKDAILY "
    strSql = strSql & "  WHERE  DEPTCODE in ('" & strDeptCode & "')"
    strSql = strSql & "  AND    YYYYMM   = '" & strYYYYMM & "'"
    strSql = strSql & "  AND    SUBSTR(SABUN,1,1)  NOT IN ('2','4','7','1')"             '의사제외
'    strSql = strSql & "  and    sabun not in ('100103','300192','300180','300305') "            '홍보 박두혁
    strSql = strSql & "  AND    DAY" & StrDay & " is not null "
    If adoSetOpen(strSql, AdoWork, 1) = True Then
        strDataChk = "Y"
    End If
'존재하지 않는다. 그럼 INSERT ROUNT을 먼저 갔다오자
'-->2022-02-15 무조건 해당부서의 없는 사람은 insert 하는걸로 변경
        
        strSql = ""
        strSql = strSql & "  INSERT INTO  TWINSA_WORKDAILY (YYYYMM, SABUN, DEPTCODE)"
        strSql = strSql & "  SELECT '" & strYYYYMM & "' YYYYMM, SABUN, DEPT  "
        strSql = strSql & "  FROM   TWINSA_MASTER"
        strSql = strSql & "  WHERE  DEPT    in ('" & strDeptCode & "')"
        strSql = strSql & "  AND (   status = 0 "
        strSql = strSql & "        OR (status = '2' AND TO_CHAR (statusdate, 'yyyymm') >= '" & strYYYYMM & "') "
        strSql = strSql & "       ) "
        strSql = strSql & "  AND    SUBSTR(SABUN,1,1)  NOT IN ('2','4','7','1')"            '의사제외
        strSql = strSql & "  AND    SABUN  NOT IN ('100103','300192','300180','300305') "    '홍보 박두혁
        strSql = strSql & "   AND last_day(indate) <  to_date('" & strYYYYMM & "01', 'yyyymmdd')  "
        strSql = strSql & "   AND SABUN NOT IN (SELECT sabun  "
        strSql = strSql & "                            FROM TWINSA_WORKDAILY A  "
        strSql = strSql & "                           WHERE     YYYYMM = '" & strYYYYMM & "'  "
        strSql = strSql & "                            AND      DEPT    in ('" & strDeptCode & "'))"
        
        adoConnect.BeginTrans
        
        If adoExecute(strSql) Then
            adoConnect.CommitTrans
        Else
            adoConnect.RollbackTrans
            MsgBox "등록에 문제가 있습니다. 전산개발팀 연락요망", vbCritical, "Error"
            Exit Sub
        End If
    
 '   End If

'Data불러오기
'-------------------------------------------------------------------------------
    StrDay = "DAY" & StrDay

    strSql = ""
    strSql = strSql & "   SELECT YYYYMM, a.deptcode dept, f.deptname, A.SABUN, B.NAMEK, C.NAME GRADE, NVL(" & StrDay & ",' ') Day, E.NAME, t.worktime, d.holyday, "
    strSql = strSql & "   decode(h.health_typ1,'Y', '대상', '') health_typ1, decode(trim(h.health_typ2),null, '', '0', '', '대상') health_typ2, "
    strSql = strSql & "   to_char(actdate_ge, 'yy/mm/dd') actdate_ge, to_char(actdate_sge, 'yy/mm/dd') actdate_sge, to_char(revdate, 'yy/mm/dd') revdate "
    strSql = strSql & "   FROM   TWINSA_WORKDAILY A, TWINSA_MASTER B,  "
    strSql = strSql & "          V_TWINSA_JIKWI C, TWNRS_BUN E, twbas_jobdate d, twinsa_dept f,  twsafe_health_employee H, "
    
    strSql = strSql & "         (SELECT * "
    strSql = strSql & "            FROM twinsa_worktime "
    strSql = strSql & "           WHERE workdate = TO_DATE ('" & DTDate.Value & "', 'yyyy-mm-dd')) t "
    
    strSql = strSql & "   WHERE  A.DEPTCODE     in ('" & strDeptCode & "')"

'원무팀은 파트별로 조회 가능
If (Right(CmbDept.text, 6) = "620100" Or Right(CmbDept.text, 6) = "620200") And CmbPart.text <> "ALL" Then
    strSql = strSql & "   and  b.DEPT2     = '" & Right(CmbPart.text, 6) & "' "
End If
    strSql = strSql & "   AND    b.contract not IN ('1016', '1017') " '단시간근로자 별도
    strSql = strSql & "   AND    YYYYMM         = '" & strYYYYMM & "'"
    strSql = strSql & "   AND    A.SABUN        = B.SABUN"
    'StrSql = StrSql & "   AND    A.SABUN        = '601233'"
    strSql = strSql & "   AND    A.SABUN        = T.SABUN(+)"
    strSql = strSql & "   AND    B.JIKWI        = C.CODE"
    strSql = strSql & "   AND    " & StrDay & " = E.BUN(+)"
    strSql = strSql & "   AND a.deptcode = f.dept"
    strSql = strSql & "   AND f.delmark IS NULL"
    strSql = strSql & "   AND    d.jobdate = TO_DATE ('" & DTDate.Value & "', 'yyyy-mm-dd') "
    strSql = strSql & "   AND   (b.statusdate IS NULL OR d.jobdate <= b.statusdate)  "
    strSql = strSql & "   AND   d.jobdate >= b.indate   "   '2024-11-18 입사일 이전에는 안보이도록 함
    strSql = strSql & "   AND    h.health_year(+) = '" & Left(strYYYYMM, 4) & "'"
    strSql = strSql & "   AND    A.sabun = h.empl_no(+) "

If StrSysDate >= GStrSysDate Then              '오늘날짜보다 조회일자가 작으면 퇴직자도 보여준다.
    strSql = strSql & "   AND    B.STATUS  in (0 , 2)"
    strSql = strSql & "   AND    A.DEPTCODE     = B.DEPT(+)"             '2014-04-25 이거 막지말것...
End If
'If strDataChk = "Y" And StrSysDate < DateAdd("d", -90, GStrSysDate) Then              ''과거 조회시 근태 저장된게 없으면 안불러오도록 함... 2024-10-11 월 중 부서이동시 계속나오는 현상
'    strSql = strSql & "   AND    " & StrDay & " is not null "                          '->다시막음... 3개월 초과후에 등록하는 경우가 많음 ...의과대학교학팀 등등
'End If
    strSql = strSql & "   AND    SUBSTR(B.SABUN,1,1) NOT IN ('2','4','7')"
    
    Select Case StrDept
        Case "310400": strSql = strSql & "  AND    A.SABUN <> '300180'  "              '정신과   '2006-08-29 손애리 선생 진료부로 변경
        Case "311900": strSql = strSql & "  AND    a.SABUN <> '300192'  "              '정신과   '2007-11-28 민철기 선생 진료부로 변경
    End Select
    
'    strSql = strSql & "   ORDER  BY C.CODE, B.NAMEK"
If StrDept = "320600" Then  '건강증진센터는 팀장-파트장-일검-특검-종검 순서로 보이도록...2016-03-14 이교승 파트장 요청
    strSql = strSql & "   ORDER BY DECODE(C.CODE, '3001','1','3002','2','3003','3','4'),  "
    strSql = strSql & "             decode(jikmu, '2007', 1, '3040', 2, '3041', 3, '3039', 4, 5), DECODE(C.CODE,3009,2, C.CODE), b.NAMEK                   " & vbLf
Else
'    StrSql = StrSql & "   ORDER BY DECODE(C.CODE,3009,2, C.CODE), b.NAMEK                   " & vbLf
    strSql = strSql & "    order by f.deptname, c.printranking, b.namek "

End If
    

    If adoSetOpen(strSql, AdoSet) = False Then Exit Sub
    
    SS.MaxRows = AdoSet.RecordCount
    SS.RowHeight(-1) = 12.3
    
    SSWeek.MaxRows = AdoSet.RecordCount + 3
    
    Do Until AdoSet.EOF
        SS.Row = AdoSet.AbsolutePosition
        SS.Col = 1:    SS.text = AdoSet.Fields("dept").Value & ""
        SS.Col = 2:    SS.text = AdoSet.Fields("deptname").Value & ""
        SS.Col = 3:    SS.text = AdoSet.Fields("SABUN").Value & ""
        SS.Col = 4:    SS.text = AdoSet.Fields("NAMEK").Value & ""
        SS.Col = 5:    SS.text = AdoSet.Fields("GRADE").Value & ""
        
        If chkWeek.Value = 1 Then   '주 근무시간 조회
            SSWeek.Row = AdoSet.AbsolutePosition + 3
            SSWeek.Col = 1:    SSWeek.text = AdoSet.Fields("SABUN").Value & ""
            SSWeek.Col = 2:    SSWeek.text = AdoSet.Fields("NAMEK").Value & ""
        End If
            
        GoSub Call_Memo
        SS.Col = 7:
        If Trim(AdoSet.Fields("DAY").Value) = "" Then
            If Trim(SS.text) = "" Then
                'If StrWeek = "SATURDAY" Or StrWeek = "SUNDAY" Then
                If StrWeek = "SATURDAY" Or StrWeek = "SUNDAY" Or Trim(AdoSet.Fields("holyday").Value) = "*" Then  '2020-01-28 토,일 외 공휴일도 비번으로 셋팅
                    SS.text = "비번" & Space(20) & "off"
                Else
                    SS.text = "정상근무" & Space(20) & "A1 "
                End If
                GoSub Call_WorkPaper '20140327 휴가서류등록내역 불러오기 인사팀 이수정 요청!!
            End If
        Else
            SS.text = AdoSet.Fields("NAME").Value & "" & Space(20) & AdoSet.Fields("Day").Value & ""
            ChkOK.Value = 1
        End If
        
        '오늘이면 근태체크기 읽어오고 아니면 데이타 가져오기
        If GStrSysDate = Format(DTDate.Value, "yyyy-mm-dd") Then
            'GoSub Call_Time
            
            SS.Col = 6 '오늘 미태그자 중에 시간 셋팅된것 있으면 보여주기..
            If SS.text = "" Then
                SS.text = AdoSet.Fields("worktime").Value & ""
            End If
        Else
        
            If AdoSet.Fields("worktime").Value & "" = "" Then
                'GoSub Call_Time
            Else
                SS.Col = 6:    SS.text = AdoSet.Fields("worktime").Value & ""
            End If
            
            If AdoSet.Fields("worktime").Value & "" >= "08:30" Then
                SS.Col = -1:    SS.ForeColor = RGB(0, 0, 255)
            ElseIf AdoSet.Fields("worktime").Value & "" = "" Then
                SS.Col = -1:    SS.ForeColor = RGB(160, 64, 160)
            Else
                SS.Col = -1:    SS.ForeColor = RGB(0, 0, 0)
            End If
        
        End If
        
        '검진 정보
        SS.Col = 13:    SS.text = AdoSet.Fields("health_typ1").Value & ""
        SS.Col = 14:    SS.text = AdoSet.Fields("actdate_ge").Value & ""
'                If AdoSet.Fields("health_typ1").Value & "" = "대상" And AdoSet.Fields("actdate_ge").Value & "" = "" Then
'                    '일반검진 내역 조회
'                    SS.Text = Get_HealthCheck("G", AdoSet.Fields("SABUN").Value & "", Left(strYYYYMM, 4))
'                End If
        SS.Col = 15:    SS.text = AdoSet.Fields("health_typ2").Value & ""
        SS.Col = 16:    SS.text = AdoSet.Fields("actdate_sge").Value & ""
'                If AdoSet.Fields("health_typ2").Value & "" = "대상" And AdoSet.Fields("actdate_sge").Value & "" = "" Then
'                    '특수검진 내역 조회
'                    SS.Text = Get_HealthCheck("S", AdoSet.Fields("SABUN").Value & "", Left(strYYYYMM, 4))
'                End If
        SS.Col = 17:    SS.text = AdoSet.Fields("revdate").Value & ""
        
        
        If SS.RowHeight(SS.Row) < SS.MaxTextRowHeight(SS.Row) Then SS.RowHeight(SS.Row) = SS.MaxTextRowHeight(SS.Row)
        AdoSet.MoveNext
    Loop

    AdoSet.Close
    Set AdoSet = Nothing

Exit Sub
'------------------------------------------------------------------------------------------------------------
Call_WorkPaper:
    
    Dim AdoWorkPaper       As ADODB.Recordset

    strSql = ""
    strSql = strSql & "  SELECT a.REMARK1, a.REMARK2, A.BUN, B.NAME,  "
    strSql = strSql & "         TO_CHAR(a.cdate1,'YYYY-MM-DD') STARTDATE, "
    strSql = strSql & "         TO_CHAR(a.cdate2,'YYYY-MM-DD') ENDDATE "
    strSql = strSql & "  FROM twinsa_work_paper A, TWNRS_BUN B "
'    strSql = strSql & " WHERE     A.bun = 'A3' "
'   2015-03-17 인사팀 고수원 요청으로 휴가 하프 자동 체크
    strSql = strSql & " WHERE     (A.bun in ('H', 'Hh', 'Hhh', 'HhA', 'HhP', 'Hs', 'Hf') "
    strSql = strSql & "            or (a.bun ='A3' AND (REMARK4 <>  '2' OR REMARK4 IS NULL))) "
    strSql = strSql & "       AND cdate2 >= TO_DATE('" & StrSysDate & "','YYYY-MM-DD')"
    strSql = strSql & "       AND cdate1 <= TO_DATE('" & StrSysDate & "','YYYY-MM-DD')"
    strSql = strSql & "       AND sabun = '" & AdoSet.Fields("SABUN").Value & "'"
    strSql = strSql & "       AND A.BUN = B.BUN "
    strSql = strSql & "       AND a.delmark is null "

   If adoSetOpen(strSql, AdoWorkPaper, 1) = True Then
   
        SS.Col = 8:    SS.text = AdoWorkPaper.Fields("STARTDATE").Value & ""
        SS.Col = 9:    SS.text = AdoWorkPaper.Fields("ENDDATE").Value & ""
        SS.Col = 10:    SS.text = AdoWorkPaper.Fields("REMARK2").Value & ""
        SS.Col = 7:    SS.text = AdoWorkPaper.Fields("NAME").Value & "" & Space(20) & AdoWorkPaper.Fields("BUN").Value & ""
    
        'If StrSysDate <> AdoWorkPaper.Fields("STARTDATE").Value Then SS.Col = 6:   SS.Lock = True
    Else
        
        '진료운영팀은 최근 리마크 불러옴 2024-03-21 장예은파트장 요청 -> 2024-07-04 막음 장예은파트장 요청
'        If strDeptCode = "313000" Then
'            GoSub Call_Memo2
'        End If
        
   End If
    
    adoSetClose AdoWorkPaper

Return
'------------------------------------------------------------------------------------------------------------
Call_Memo:
    
    Dim AdoMemo       As ADODB.Recordset


    strSql = ""
    strSql = strSql & "  SELECT REMARK, REMARK2, REMARK3, A.BUN, B.NAME,  "
    strSql = strSql & "         TO_CHAR(STARTDATE,'YYYY-MM-DD') STARTDATE, "
    strSql = strSql & "         TO_CHAR(ENDDATE,'YYYY-MM-DD') ENDDATE "
    strSql = strSql & "  FROM   TWINSA_WORKMEMO A, TWNRS_BUN B"
    strSql = strSql & "  WHERE  WORKDATE  = TO_DATE('" & StrSysDate & "','YYYY-MM-DD')"
    strSql = strSql & "  AND    SABUN     = '" & AdoSet.Fields("SABUN").Value & "'"
    strSql = strSql & "  AND    (A.BUN is null or A.BUN     = B.BUN) "

   If adoSetOpen(strSql, AdoMemo, 1) = True Then
   
        SS.Col = 8:    SS.text = AdoMemo.Fields("STARTDATE").Value & ""
        SS.Col = 9:    SS.text = AdoMemo.Fields("ENDDATE").Value & ""
        SS.Col = 10:    SS.text = AdoMemo.Fields("REMARK").Value & ""
        SS.Col = 7:    SS.text = AdoMemo.Fields("NAME").Value & "" & Space(20) & AdoMemo.Fields("BUN").Value & ""
'        SS.Col = 5:    SS.Lock = True
        SS.Col = 11:    SS.text = AdoMemo.Fields("REMARK2").Value & ""
        SS.Col = 12:    SS.text = AdoMemo.Fields("REMARK3").Value & ""
        If StrSysDate <> AdoMemo.Fields("STARTDATE").Value Then SS.Col = 8:   SS.Lock = True
   
   Else

        strSql = ""
        strSql = strSql & "  SELECT REMARK, REMARK2, REMARK3, A.BUN, B.NAME,  "
        strSql = strSql & "         TO_CHAR(STARTDATE,'YYYY-MM-DD') STARTDATE, "
        strSql = strSql & "         TO_CHAR(ENDDATE,'YYYY-MM-DD') ENDDATE "
        strSql = strSql & "  FROM   TWINSA_WORKMEMO A, TWNRS_BUN B"
        strSql = strSql & "  WHERE  STARTDATE <= TO_DATE('" & StrSysDate & "','YYYY-MM-DD')"
        strSql = strSql & "  AND    ENDDATE   >= TO_DATE('" & StrSysDate & "','YYYY-MM-DD')"
        strSql = strSql & "  AND    SABUN     = '" & AdoSet.Fields("SABUN").Value & "'"
        strSql = strSql & "  AND    A.BUN     = B.BUN"
        strSql = strSql & "  ORDER  BY ENDDATE DESC"
        
        If adoSetOpen(strSql, AdoMemo, 1) Then
            
            SS.Col = 8:    SS.text = AdoMemo.Fields("STARTDATE").Value & "":  SS.Lock = True
            SS.Col = 9:    SS.text = AdoMemo.Fields("ENDDATE").Value & ""
            SS.Col = 10:    SS.text = AdoMemo.Fields("REMARK").Value & ""
            SS.Col = 7:    SS.text = AdoMemo.Fields("NAME").Value & "" & Space(20) & AdoMemo.Fields("BUN").Value & ""
'            SS.Col = 9:    SS.Text = AdoMemo.Fields("REMARK2").Value & ""
'            SS.Col = 10:    SS.Text = AdoMemo.Fields("REMARK3").Value & ""
'            SS.Col = 5:    SS.Lock = True
            
        If StrSysDate <> AdoMemo.Fields("STARTDATE").Value Then SS.Col = 8:   SS.Lock = True
        
        End If
    
    End If
    
    
    adoSetClose AdoMemo

Return

'------------------------------------------------------------------------------------------------------------
Call_Memo2:

    strSql = "  "
    strSql = strSql & "   SELECT workdate, REMARK "
    strSql = strSql & "   FROM   TWINSA_WORKMEMO A "
    strSql = strSql & "   WHERE  WORKDATE  <=  TO_DATE('" & StrSysDate & "','YYYY-MM-DD')"
    strSql = strSql & "   and    WORKDATE  >=  TO_DATE('" & StrSysDate & "','YYYY-MM-DD') - 7 "
    strSql = strSql & "   AND    SABUN     = '" & AdoSet.Fields("SABUN").Value & "'"
    strSql = strSql & "   and bun <> 'off' "
    strSql = strSql & "   AND ROWNUM = 1 "
    strSql = strSql & "   ORDER BY workdate desc "

   If adoSetOpen(strSql, AdoMemo, 1) = True Then
        SS.Col = 10:    SS.text = AdoMemo.Fields("REMARK").Value & ""
    End If
    
    
    adoSetClose AdoMemo

Return

'------------------------------------------------------------------------------------------------------------
Call_Time:
    
    Dim AdoTime       As ADODB.Recordset

    strSql = ""
    strSql = strSql & "SELECT substr(eltime,1,2) || ':' || substr(eltime,3,2) eltime "
    strSql = strSql & "  FROM eventlog "
    strSql = strSql & " WHERE eldate = ? "
    strSql = strSql & "   AND SUBSTR (eluserid, 11, 6) = ? "
    strSql = strSql & "   AND elccode IN ('1', '2', '3') "

    Call adoCmd(strSql, 1)
    
    With adoCommand
        .Parameters.Append .CreateParameter("eldate", adChar, adParamInput, 8, strDate)
        .Parameters.Append .CreateParameter("eluserid", adChar, adParamInput, 16, AdoSet.Fields("SABUN").Value)
    End With
            
    Set AdoTime = adoCommand.Execute
    If AdoTime.RecordCount > 0 Then
        SS.Col = 6:    SS.text = AdoTime.Fields("eltime").Value & ""
        
        If AdoTime.Fields("ELTIME").Value & "" >= "08:30" Then
            SS.Col = -1:    SS.ForeColor = RGB(0, 0, 255)
        ElseIf AdoTime.Fields("ELTIME").Value & "" = "" Then
'            SS.Col = -1:    SS.ForeColor = RGB(0, 255, 0)
            SS.Col = -1:    SS.ForeColor = RGB(160, 64, 160)
        Else
            SS.Col = -1:    SS.ForeColor = RGB(0, 0, 0)
        End If
    Else
        SS.Col = -1:    SS.ForeColor = RGB(160, 64, 160)
    End If
    
    adoSetClose AdoTime

Return

End Sub

Private Sub Data_Search2()
Dim strDate             As String
Dim strDeptCode         As String

    If CmbPart.ListCount > 0 Then
        StrDept = Right(CmbPart.text, 6)
    Else
        StrDept = Right(CmbDept.text, 6)
    End If
    
    strDeptCode = StrDept
    
    '진료운영팀 통합관리 : 신경과, 뇌파검사실, 소화기능검사실, 폐기능검사실, 청력검사실, 전정기능검사실, 언어치료실, 수면다원검사실,근전도검사실,신경심리검사실
    If Right(CmbDept.text, 6) = "313000" And CmbPart.text = "ALL" Then
        strDeptCode = "313000','320402','320407','320800','320401','320405','320406', '320412', '320404', '320403', '310300"
    End If
    
    '원무팀은 파트별로 관리 가능하도록 함. 2022-09 -> 파트를 선택해도 입력은 원무팀 코드로
    If Right(CmbDept.text, 6) = "620100" Or Right(CmbDept.text, 6) = "620200" Then
        strDeptCode = Right(CmbDept.text, 6)
    End If
    
    
    strDate = Format(DTDate.Value, "yyyymmdd")
    
 '   End If

'Data불러오기
'-------------------------------------------------------------------------------
    'StrDay = "DAY" & StrDay

    strSql = ""
    strSql = strSql & "   SELECT YYYYMM, a.deptcode dept, f.deptname, A.SABUN, B.NAMEK, C.NAME GRADE, NVL(" & StrDay & ",' ') Day, E.NAME, t.worktime, d.holyday, "
    strSql = strSql & "   decode(h.health_typ1,'Y', '대상', '') health_typ1, decode(trim(h.health_typ2),null, '', '0', '', '대상') health_typ2, "
    strSql = strSql & "   to_char(actdate_ge, 'yy/mm/dd') actdate_ge, to_char(actdate_sge, 'yy/mm/dd') actdate_sge, to_char(revdate, 'yy/mm/dd') revdate "
    strSql = strSql & "   FROM   TWINSA_WORKDAILY A, TWINSA_MASTER B,  "
    strSql = strSql & "          V_TWINSA_JIKWI C, TWNRS_BUN E, twbas_jobdate d, twinsa_dept f,  twsafe_health_employee H, "
    
    strSql = strSql & "         (SELECT * "
    strSql = strSql & "            FROM twinsa_worktime "
    strSql = strSql & "           WHERE workdate = TO_DATE ('" & DTDate.Value & "', 'yyyy-mm-dd')) t "
    
    strSql = strSql & "   WHERE  A.DEPTCODE     in ('" & strDeptCode & "')"

'원무팀은 파트별로 조회 가능
If (Right(CmbDept.text, 6) = "620100" Or Right(CmbDept.text, 6) = "620200") And CmbPart.text <> "ALL" Then
    strSql = strSql & "   and  b.DEPT2     = '" & Right(CmbPart.text, 6) & "' "
End If
    strSql = strSql & "   AND    b.contract IN ('1016', '1017') " '단시간근로자 별도
    strSql = strSql & "   AND    YYYYMM         = '" & strYYYYMM & "'"
    strSql = strSql & "   AND    A.SABUN        = B.SABUN"
    strSql = strSql & "   AND    A.SABUN        = T.SABUN(+)"
    strSql = strSql & "   AND    B.JIKWI        = C.CODE"
    strSql = strSql & "   AND    " & StrDay & " = E.BUN(+)"
    strSql = strSql & "   AND a.deptcode = f.dept"
    strSql = strSql & "   AND f.delmark IS NULL"
    strSql = strSql & "   AND    d.jobdate = TO_DATE ('" & DTDate.Value & "', 'yyyy-mm-dd') "
    strSql = strSql & "   AND   (b.statusdate IS NULL OR d.jobdate <= b.statusdate)  "
    
    strSql = strSql & "   AND    h.health_year(+) = '" & Left(strYYYYMM, 4) & "'"
    strSql = strSql & "   AND    A.sabun = h.empl_no(+) "

If StrSysDate >= GStrSysDate Then              '오늘날짜보다 조회일자가 작으면 퇴직자도 보여준다.
    strSql = strSql & "   AND    B.STATUS  in (0 , 2)"
    strSql = strSql & "   AND    A.DEPTCODE     = B.DEPT(+)"             '2014-04-25 이거 막지말것...
End If
    strSql = strSql & "   AND    SUBSTR(B.SABUN,1,1) NOT IN ('2','4','7')"
    
    Select Case StrDept
        Case "310400": strSql = strSql & "  AND    A.SABUN <> '300180'  "              '정신과   '2006-08-29 손애리 선생 진료부로 변경
        Case "311900": strSql = strSql & "  AND    a.SABUN <> '300192'  "              '정신과   '2007-11-28 민철기 선생 진료부로 변경
    End Select

If StrDept = "320600" Then  '건강증진센터는 팀장-파트장-일검-특검-종검 순서로 보이도록...2016-03-14 이교승 파트장 요청
    strSql = strSql & "   ORDER BY DECODE(C.CODE, '3001','1','3002','2','3003','3','4'),  "
    strSql = strSql & "             decode(jikmu, '2007', 1, '3040', 2, '3041', 3, '3039', 4, 5), DECODE(C.CODE,3009,2, C.CODE), b.NAMEK                   " & vbLf
Else
    strSql = strSql & "    order by f.deptname, c.printranking, b.namek "

End If
    
    If adoSetOpen(strSql, AdoSet) = False Then
        'SSTab1.TabVisible(1) = False    '단기근로자가 없으면 안보이게!
        SSTab1.Tab = 0
        Exit Sub
    End If
    
    SSTab1.Tab = 1
    
    ss2.MaxRows = AdoSet.RecordCount
    ss2.RowHeight(-1) = 12.3
    
    Do Until AdoSet.EOF
        ss2.Row = AdoSet.AbsolutePosition
        ss2.Col = 1:    ss2.text = AdoSet.Fields("dept").Value & ""
        ss2.Col = 2:    ss2.text = AdoSet.Fields("deptname").Value & ""
        ss2.Col = 3:    ss2.text = AdoSet.Fields("SABUN").Value & ""
        ss2.Col = 4:    ss2.text = AdoSet.Fields("NAMEK").Value & ""
            
        GoSub Call_Memo
        ss2.Col = 5:
        If Trim(AdoSet.Fields("DAY").Value) = "" Then
            If Trim(ss2.text) = "" Then
                If StrWeek = "SATURDAY" Or StrWeek = "SUNDAY" Or Trim(AdoSet.Fields("holyday").Value) = "*" Then  '2020-01-28 토,일 외 공휴일도 비번으로 셋팅
                    ss2.text = "비번" & Space(20) & "off"
                Else
                    ss2.text = "정상근무" & Space(20) & "A1 "
                    GoSub Call_Memo_Time  '최근 입력된 시간 세팅해주기!
                End If
            End If
        Else
            ss2.text = AdoSet.Fields("NAME").Value & "" & Space(20) & AdoSet.Fields("Day").Value & ""
            
        End If
        
        '검진 정보
        ss2.Col = 12:    ss2.text = AdoSet.Fields("health_typ1").Value & ""
        ss2.Col = 13:    ss2.text = AdoSet.Fields("actdate_ge").Value & ""
        ss2.Col = 14:    ss2.text = AdoSet.Fields("health_typ2").Value & ""
        ss2.Col = 15:    ss2.text = AdoSet.Fields("actdate_sge").Value & ""
        ss2.Col = 16:    ss2.text = AdoSet.Fields("revdate").Value & ""
        
        
        If ss2.RowHeight(ss2.Row) < ss2.MaxTextRowHeight(ss2.Row) Then ss2.RowHeight(ss2.Row) = ss2.MaxTextRowHeight(ss2.Row)
        AdoSet.MoveNext
    Loop

    AdoSet.Close
    Set AdoSet = Nothing

Exit Sub

'------------------------------------------------------------------------------------------------------------
Call_Memo:
    
    Dim AdoMemo       As ADODB.Recordset


    strSql = ""
    strSql = strSql & "  SELECT REMARK, REMARK1, REMARK2, REMARK3, A.BUN, B.NAME,  "
    strSql = strSql & "         TO_CHAR(STARTDATE,'YYYY-MM-DD') STARTDATE, "
    strSql = strSql & "         TO_CHAR(ENDDATE,'YYYY-MM-DD') ENDDATE, "
    strSql = strSql & "         substr(remark1,instr(remark1,'s')+1,5) Time1, "
    strSql = strSql & "         substr(remark1,instr(remark1,'e')+1,5) Time2, "
    strSql = strSql & "         substr(remark1,instr(remark1,'t')+1) Time3 "
    strSql = strSql & "  FROM   TWINSA_WORKMEMO A, TWNRS_BUN B"
    strSql = strSql & "  WHERE  WORKDATE  = TO_DATE('" & StrSysDate & "','YYYY-MM-DD')"
    strSql = strSql & "  AND    SABUN     = '" & AdoSet.Fields("SABUN").Value & "'"
    strSql = strSql & "  AND    (A.BUN is null or A.BUN     = B.BUN) "

   If adoSetOpen(strSql, AdoMemo, 1) = True Then
   
        ss2.Col = 6:    ss2.text = AdoMemo.Fields("Time1").Value & ""
        ss2.Col = 7:    ss2.text = AdoMemo.Fields("Time2").Value & ""
        ss2.Col = 8:    ss2.text = AdoMemo.Fields("Time3").Value & ""
        ss2.Col = 9:    ss2.text = AdoMemo.Fields("REMARK").Value & ""
        ss2.Col = 5:    ss2.text = AdoMemo.Fields("NAME").Value & "" & Space(20) & AdoMemo.Fields("BUN").Value & ""

        ss2.Col = 10:    ss2.text = AdoMemo.Fields("REMARK2").Value & ""
        ss2.Col = 11:    ss2.text = AdoMemo.Fields("REMARK3").Value & ""
'    Else
'        ss2.Col = 6:    ss2.Text = "00:00"
'        ss2.Col = 7:    ss2.Text = "00:00"
'        ss2.Col = 8:    ss2.Text = "0.0"
    End If
    
    
    adoSetClose AdoMemo

Return

'------------------------------------------------------------------------------------------------------------
Call_Memo_Time:

    strSql = "  "
    strSql = strSql & "   SELECT workdate, REMARK1,   "
    strSql = strSql & "          substr(remark1,instr(remark1,'s')+1,5) Time1,  "
    strSql = strSql & "          substr(remark1,instr(remark1,'e')+1,5) Time2,  "
    strSql = strSql & "          substr(remark1,instr(remark1,'t')+1) Time3  "
    strSql = strSql & "   FROM   TWINSA_WORKMEMO A "
    strSql = strSql & "   WHERE  WORKDATE  <=  TO_DATE('" & StrSysDate & "','YYYY-MM-DD')"
    strSql = strSql & "   AND    SABUN     = '" & AdoSet.Fields("SABUN").Value & "'"
    strSql = strSql & "   and bun <> 'off' "
    strSql = strSql & "   AND ROWNUM = 1 "
    strSql = strSql & "   ORDER BY workdate desc "


   If adoSetOpen(strSql, AdoMemo, 1) = True Then
        ss2.Col = 6:    ss2.text = AdoMemo.Fields("Time1").Value & ""
        ss2.Col = 7:    ss2.text = AdoMemo.Fields("Time2").Value & ""
        ss2.Col = 8:    ss2.text = AdoMemo.Fields("Time3").Value & ""
    End If
    
    
    adoSetClose AdoMemo

Return


End Sub

Private Sub SS_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
Dim strDate            As String

    SS.Col = Col: SS.Row = Row
    If Trim(SS.text) = "" Then Exit Sub
    
    
    Select Case Col
        Case 7:
            If Right(SS.text, 3) = "A1 " Then
                SS.Col = 8:     SS.text = ""
                SS.Col = 9:     SS.text = ""
                SS.Col = 10:     SS.text = ""
            End If
            
        Case 8:
            SS.Col = 8:     SS.text = Date_Format(SS.text): strDate = SS.text
            SS.Col = 9:     If Trim(SS.text) = "" Then SS.text = Date_Format(strDate)
        Case 9:
        SS.Col = 9:     SS.text = Date_Format(SS.text)
    End Select
End Sub

Private Sub SS_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sWord As String
    
    'ctrl + f
    If KeyCode = 70 And Shift = 2 Then
        sWord = InputBox("교직원번호 또는 이름을 입력하세요", "검색")
        
        For i = 1 To SS.MaxRows
           
            '이름
            SS.Row = i: SS.Col = 4
            If Trim(SS.text) = Trim(sWord) Then
                SS.SetFocus
                SS.SetSelection 2, i, 2, i
                SS.SetActiveCell 2, i
                Exit Sub
            End If
            
            SS.Row = i: SS.Col = 3
            '교직원번호
            If Trim(SS.text) = Trim(sWord) Then
                SS.SetFocus
                SS.SetSelection 1, i, 1, i
                SS.SetActiveCell 1, i
                Exit Sub
            End If
            
        Next i
    End If
End Sub

Private Sub Data_Search_Week()
Dim StrJobDate   As String

Dim StrColor     As String
Dim StrSabun     As String
Dim j, i         As Single

Dim Stryymm      As String
Dim StrDD        As String
Dim StrDay       As String

Dim AdoSub       As ADODB.Recordset
Dim nTotAmt(3)     As Double

    If SSWeek.DataRowCnt = 3 Then Exit Sub
    If DTDate.Value < "2019-06-01" Then SSWeek.Visible = False:  Exit Sub
'    chkWeek.Visible = True
'    SSWeek.Visible = True
    
    strSql = ""
    strSql = strSql & "  SELECT jobdate, holyday, TO_CHAR (jobdate, 'day') day"
    strSql = strSql & "    FROM twbas_jobdate a, "
    strSql = strSql & "         (SELECT jobdate fdate, jobdate + 6 AS tdate "
    strSql = strSql & "            FROM TWBAS_JOBDATE "
    strSql = strSql & "           WHERE     TO_CHAR (jobdate, 'd') = 1 "
    strSql = strSql & "                 AND jobdate > ? - 7 "
    strSql = strSql & "                 AND jobdate < ? + 1) b "
    strSql = strSql & "   WHERE a.jobdate >= b.fdate AND a.jobdate <= b.tdate "
    strSql = strSql & "ORDER BY jobdate "
    
    Call adoCmd(strSql, 1)
    With adoCommand
        .Parameters.Append .CreateParameter("jobdate", adDate, adParamInput, 10, DTDate.Value)
        .Parameters.Append .CreateParameter("jobdate", adDate, adParamInput, 10, DTDate.Value)
    End With
    
    Set AdoSet = adoCommand.Execute
    If AdoSet.RecordCount = 0 Then Exit Sub
    
'    SSWeek.BlockMode = True
    Screen.MousePointer = vbHourglass
    
    
    i = 0
    j = 0

'Data불러오기
'-------------------------------------------------------------------------------
    Do Until AdoSet.EOF
        
        StrJobDate = AdoSet.Fields("jobdate").Value & ""
'        strMM = Format(StrJobDate, "mm")
'        StrYY = Format(StrJobDate, "yyyy")
        StrDD = Format(StrJobDate, "dd")
        Stryymm = Format(StrJobDate, "yyyymm")
        
        StrDay = "Day" & StrDD
        
        If AdoSet.Fields("holyday").Value = "*" Then
            StrColor = &HC0&
        Else
            If Trim(AdoSet.Fields("day").Value) & "" = "토요일" Or Trim(AdoSet.Fields("day").Value & "") = "saturday" Then
                StrColor = &HC00000
            Else
                StrColor = &H0&
            End If
        End If
    
        SSWeek.Row = 2
        SSWeek.Col = 3 + i:     SSWeek.ForeColor = StrColor
        SSWeek.Col = 11 + i:    SSWeek.ForeColor = StrColor
        SSWeek.Col = 19 + i:    SSWeek.ForeColor = StrColor
    
    
        SSWeek.Row = 3
        SSWeek.Col = 3 + i:     SSWeek.text = StrDD:  SSWeek.ForeColor = StrColor
        SSWeek.Col = 11 + i:    SSWeek.text = StrDD:  SSWeek.ForeColor = StrColor
        SSWeek.Col = 19 + i:    SSWeek.text = StrDD:  SSWeek.ForeColor = StrColor
                
'        StrSql = ""
'        StrSql = StrSql & " select   NVL(" & StrDay & ",' ') Day, e.worktime   "
'        StrSql = StrSql & "   FROM   TWINSA_WORKDAILY A, TWNRS_BUN E "
'        StrSql = StrSql & "   WHERE  a.sabun = ? "
'        StrSql = StrSql & "   AND    YYYYMM  = ? "
'        StrSql = StrSql & "   AND    " & StrDay & " = E.BUN(+)"
                
        If StrJobDate < "2019-06-01" Then
            '2019-06-01 이전데이타는 조회하지 않는다
        Else
            strSql = ""
            strSql = strSql & " SELECT SUM (WORKTIME) WORKTIME, "
            strSql = strSql & "       SUM (ADDTIME) ADDTIME, "
            strSql = strSql & "       SUM (CALLTIME) CallTime "
            strSql = strSql & "  FROM (SELECT NVL(e.worktime, 0) WORKTIME, 0 ADDTIME, 0 CALLTIME "
            strSql = strSql & "          FROM TWINSA_WORKDAILY A, TWNRS_BUN E "
            strSql = strSql & "         WHERE a.sabun = ? "
            strSql = strSql & "           AND YYYYMM = ? "
            strSql = strSql & "           AND " & StrDay & " = E.BUN(+)"
            strSql = strSql & "        UNION ALL "
            strSql = strSql & "        SELECT DECODE (C.HOLYDAY, '*', WORKTIME2, WORKTIME) WORKTIME, "
            strSql = strSql & "               0 ADDTIME, "
            strSql = strSql & "               0 CALLTIME "
            strSql = strSql & "          FROM twdang_master A, "
            strSql = strSql & "           (SELECT * "
            strSql = strSql & "              FROM TWDANG_TIME A "
            strSql = strSql & "             WHERE JOBDATE = "
            strSql = strSql & "                      (SELECT MAX (JOBDATE) "
            strSql = strSql & "                         FROM TWDANG_TIME "
            strSql = strSql & "                        WHERE     JOBDATE <= ?  "
            strSql = strSql & "                              AND DEPTCODE = A.DEPTCODE)) B, "
            strSql = strSql & "           TWBAS_JOBDATE C "
            strSql = strSql & "         WHERE SABUN = ? "
            strSql = strSql & "           AND BDATE = ? "
            strSql = strSql & "           AND A.DEPTCODE = B.DEPTCODE "
            strSql = strSql & "           AND A.DANGCODE = B.DANGCODE "
            strSql = strSql & "           AND A.SEQNO = B.SEQNO "
            strSql = strSql & "           AND A.BDATE = C.JOBDATE "
            strSql = strSql & "        UNION ALL "
            strSql = strSql & "        SELECT 0, "
            strSql = strSql & "               round(TRUNC ( ( (END_TIME - START_TIME) * 1440) + 0.5) / 60, 1) ADDTIME, "
            strSql = strSql & "               0 CALLTIME "
            strSql = strSql & "          FROM twinsa_workadd_master A, twinsa_workadd_detail T "
            strSql = strSql & "         WHERE A.SABUN = ? "
            strSql = strSql & "           AND A.WORKDATE = ? "
            strSql = strSql & "           AND A.DELMARK = 'S' AND a.seqno = t.seqno(+) "
            strSql = strSql & "        UNION ALL "
            strSql = strSql & "        SELECT 0, 0, round(TRUNC ( ( (totime - fromtime) * 1440) + 0.5) /60, 1) totaltime "
            strSql = strSql & "          FROM TWINSA_WORK_CALL "
            strSql = strSql & "         WHERE SABUN = ? "
            strSql = strSql & "           AND CALLDATE = ? "
            strSql = strSql & "         ) "
                    
            For j = 4 To SSWeek.MaxRows
                SSWeek.Col = 1
                SSWeek.Row = j:      StrSabun = SSWeek.text
                
                Call adoCmd(strSql, 1)
                With adoCommand
                    .Parameters.Append .CreateParameter("sabun", adChar, adParamInput, 6, StrSabun)
                    .Parameters.Append .CreateParameter("YYYYMM", adChar, adParamInput, 6, Stryymm)
                    
                    .Parameters.Append .CreateParameter("WORKDATE", adDate, adParamInput, 10, Format(StrJobDate, "YYYY-MM-DD"))
                    .Parameters.Append .CreateParameter("sabun", adChar, adParamInput, 6, StrSabun)
                    .Parameters.Append .CreateParameter("WORKDATE", adDate, adParamInput, 10, Format(StrJobDate, "YYYY-MM-DD"))
                
                    .Parameters.Append .CreateParameter("sabun", adChar, adParamInput, 6, StrSabun)
                    .Parameters.Append .CreateParameter("WORKDATE", adDate, adParamInput, 10, Format(StrJobDate, "YYYY-MM-DD"))
                
                    .Parameters.Append .CreateParameter("sabun", adChar, adParamInput, 6, StrSabun)
                    .Parameters.Append .CreateParameter("CALLDATE", adDate, adParamInput, 10, Format(StrJobDate, "YYYY-MM-DD"))
                
                End With
                
                Set AdoSub = adoCommand.Execute
                If AdoSub.RecordCount > 0 Then
                    SSWeek.Col = 3 + i:          SSWeek.text = IIf(AdoSub.Fields("WORKTIME").Value & "" = 0, "", AdoSub.Fields("WORKTIME").Value & "")
                    SSWeek.Col = 11 + i:         SSWeek.text = IIf(AdoSub.Fields("AddTIME").Value & "" = 0, "", AdoSub.Fields("AddTIME").Value & "")
                    SSWeek.Col = 19 + i:         SSWeek.text = IIf(AdoSub.Fields("CallTIME").Value & "" = 0, "", AdoSub.Fields("CallTIME").Value & "")
                Else
                    SSWeek.Col = 3 + i:       SSWeek.text = ""
                    SSWeek.Col = 11 + i:      SSWeek.text = ""
                    SSWeek.Col = 19 + i:      SSWeek.text = ""
                End If
                
            Next j
        End If
        i = i + 1
        AdoSet.MoveNext
    Loop
    
    AdoSet.Close
    Set AdoSet = Nothing
    
    i = 0
    j = 0
    
    For i = 4 To SSWeek.MaxRows
        
        nTotAmt(1) = 0
        nTotAmt(2) = 0
        nTotAmt(3) = 0
        
        For j = 0 To 6
            SSWeek.Row = i
            SSWeek.Col = j + 3:      nTotAmt(1) = nTotAmt(1) + IIf(SSWeek.text = "", 0, SSWeek.text)
            SSWeek.Col = j + 11:     nTotAmt(2) = nTotAmt(2) + IIf(SSWeek.text = "", 0, SSWeek.text)
            SSWeek.Col = j + 19:     nTotAmt(3) = nTotAmt(3) + IIf(SSWeek.text = "", 0, SSWeek.text)
        Next j
    
        SSWeek.Col = 10:   SSWeek.text = nTotAmt(1)
        SSWeek.Col = 18:   SSWeek.text = nTotAmt(2)
        SSWeek.Col = 26:   SSWeek.text = nTotAmt(3)
    
        SSWeek.Col = 27:   SSWeek.text = nTotAmt(1) + nTotAmt(2) + nTotAmt(3)
        If nTotAmt(1) + nTotAmt(2) + nTotAmt(3) > 46 Then
            SSWeek.ForeColor = RGB(255, 0, 0)
        ElseIf nTotAmt(1) + nTotAmt(2) + nTotAmt(3) > 40 Then
            SSWeek.ForeColor = RGB(0, 0, 255)
        Else
            SSWeek.ForeColor = RGB(0, 0, 0)
        End If
    
    Next i

'    SSWeek.BlockMode = False
    Screen.MousePointer = vbDefault

End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    ss2.Col = Col: ss2.Row = Row
    If Trim(ss2.text) = "" Then Exit Sub
    
    Select Case Col
        Case 5:
            If Right(ss2.text, 3) = "off" Then
                ss2.Col = 6:     ss2.text = ""
                ss2.Col = 7:     ss2.text = ""
                ss2.Col = 8:     ss2.text = ""
            End If
    End Select
End Sub
