VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form15 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇæÒÇä"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9540
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form15.frx":0000
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   9540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "ÖÇíÚÇÊ"
      Height          =   495
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   6840
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ç ÇæÒÇä"
      Height          =   495
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   6840
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      Caption         =   "ÇÝÒæÏä ãæÇÏ"
      Height          =   1695
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox Combo2 
         Height          =   465
         Left            =   2280
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ËÈÊ"
         Height          =   465
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   4215
      End
      Begin VB.ComboBox Combo4 
         Height          =   465
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Text            =   "Combo4"
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox Combo5 
         Height          =   465
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Text            =   "Combo5"
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÞÏÇÑ"
         Height          =   495
         Index           =   7
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2760
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\PraticGroup\End Of Money To Account\1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\PraticGroup\End Of Money To Account\1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "g_ozanunder"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Form15.frx":2CFA
      Height          =   4215
      Left            =   360
      TabIndex        =   10
      Top             =   1920
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   29
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "qq"
         Caption         =   "äÇã ãÇÏå"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1080
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "meghdar"
         Caption         =   "ãÞÏÇÑ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1080
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\PraticGroup\End Of Money To Account\1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\PraticGroup\End Of Money To Account\1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "infomavad"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form15.frx":2D0F
      Height          =   6495
      Left            =   5040
      TabIndex        =   11
      Top             =   240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   11456
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   26
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ãÍÕæáÇÊ"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "idmavad"
         Caption         =   "ßÏ ãÍÕæá"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1080
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "mavad"
         Caption         =   "äÇã ãÍÕæá"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1080
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2234.835
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4800
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\PraticGroup\End Of Money To Account\1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\PraticGroup\End Of Money To Account\1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "P_g_ozan"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   13
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   6240
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "ãÌãæÚ :"
      Height          =   495
      Index           =   12
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   375
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu mnumain1 
      Caption         =   "ÊÚÑíÝ"
      Begin VB.Menu mnumavadmasraf 
         Caption         =   "ãæÇÏ Çæáíå ãÕÑÝí"
      End
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Dim q As Integer, commove As Integer

Private Sub Combo2_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
Dim tmp
tmp = SendMessage(Combo2.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text1(5).SetFocus
End Sub

Private Sub Command1_Click()
Dim db1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
db1.Open Form3.Text10.Text
  rs1.Open "DELETE FROM P_g_ozan", db1
db1.Close

Adodc1.Recordset.MoveFirst
Do
  Adodc2.ConnectionString = Form3.Text10.Text
  Adodc2.CommandType = adCmdUnknown
  Adodc2.RecordSource = "select * from g_ozanunder where (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmavad)) + ")"
  Adodc2.Refresh
  If Adodc2.Recordset.RecordCount > 0 Then
    Adodc2.Recordset.MoveFirst
    Do
      Adodc3.Refresh
      Adodc3.Recordset.AddNew
      Adodc3.Recordset.Fields!idmavad = Adodc1.Recordset.Fields!idmavad
      Adodc3.Recordset.Fields!mavad = Adodc1.Recordset.Fields!mavad
      Adodc3.Recordset.Fields!idmade = Adodc2.Recordset.Fields!idmade
      Adodc3.Recordset.Fields!meghdar = Adodc2.Recordset.Fields!meghdar
      Adodc3.Recordset.Fields!qq = Adodc2.Recordset.Fields!qq
      Adodc3.Recordset.Fields!j5 = Adodc2.Recordset.Fields!j5
      Adodc3.Recordset.Fields!j6 = Adodc2.Recordset.Fields!j6
      db1.Open Form3.Text10.Text
        rs1.Open "SELECT * FROM g_gardeshmavad WHERE (nomade=1) AND (idmade=" + Trim(Str(Adodc2.Recordset.Fields!idmade)) + ")", db1
          Adodc3.Recordset.Fields!j51 = rs1.Fields!masrafteydoremeghdar
          Adodc3.Recordset.Fields!j61 = rs1.Fields!masrafteydoremablagh
        rs1.Close
      db1.Close
      
      Adodc3.Recordset.Fields!j52 = Adodc2.Recordset.Fields!j5 - Adodc3.Recordset.Fields!j51
      Adodc3.Recordset.Fields!j62 = Adodc2.Recordset.Fields!j6 - Adodc3.Recordset.Fields!j61

      Adodc3.Recordset.Fields!j52 = 0
      Adodc3.Recordset.Fields!j62 = 0

      Adodc3.Recordset.Update
      Adodc2.Recordset.MoveNext
    Loop Until Adodc2.Recordset.EOF = True
  End If
  
  Adodc1.Recordset.MoveNext
Loop Until Adodc1.Recordset.EOF = True
Form9.Show
End Sub

Private Sub Command2_Click()
Dim db1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

db1.Open Form3.Text10.Text
  rs1.Open "DELETE FROM P_g_ozan", db1
db1.Close
      
      Adodc3.Refresh

db1.Open Form3.Text10.Text
  Form8.Adodc1.Recordset.MoveFirst
  Do
    rs1.Open "SELECT * FROM g_gardeshmavad WHERE (nomade=1) AND (idmade=" + Trim(Str(Form8.Adodc1.Recordset.Fields!idmavad)) + ")", db1
    rs2.Open "SELECT SUM(j5) As j51, SUM(j6) As j61 FROM g_ozanunder WHERE (idmade='" + Trim(Str(Form8.Adodc1.Recordset.Fields!idmavad)) + "')", db1
      Adodc3.Recordset.AddNew
      Adodc3.Recordset.Fields!idmade = Form8.Adodc1.Recordset.Fields!idmavad
      Adodc3.Recordset.Fields!qq = Form8.Adodc1.Recordset.Fields!mavad
      Adodc3.Recordset.Fields!j5 = rs1.Fields!masrafteydoremeghdar
      Adodc3.Recordset.Fields!j6 = rs1.Fields!masrafteydoremablagh
      Adodc3.Recordset.Fields!j51 = rs2.Fields!j51
      Adodc3.Recordset.Fields!j61 = rs2.Fields!j61
      If IsNull(rs2.Fields!j61) = False Then
        Adodc3.Recordset.Fields!j52 = Val(Adodc3.Recordset.Fields!j5) - Val(Adodc3.Recordset.Fields!j51)
        Adodc3.Recordset.Fields!j62 = Val(Adodc3.Recordset.Fields!j6) - Val(Adodc3.Recordset.Fields!j61)
      End If
      Adodc3.Recordset.Update
    rs2.Close
    rs1.Close
    Form8.Adodc1.Recordset.MoveNext
  Loop Until Form8.Adodc1.Recordset.EOF = True
db1.Close
Form11.Show
End Sub

Private Sub Command3_Click()
If (Combo2.ListIndex = -1) Or (Text1(5).Text = "") Then
  MsgBox "áØÝÇ ÊãÇãí ÝíáÏ åÇ ÑÇ Ê˜ãíá äãÇííÏ", vbCritical + vbMsgBoxRight, ""
  Exit Sub
End If

If Adodc2.Recordset.RecordCount > 0 Then
  Adodc2.Recordset.MoveFirst
  Do
    If Adodc2.Recordset.Fields!idmade = Combo4.List(Combo2.ListIndex) Then
      Exit Sub
    End If
    Adodc2.Recordset.MoveNext
  Loop Until Adodc2.Recordset.EOF = True
End If

Adodc2.Refresh
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields!idmahsol = Adodc1.Recordset.Fields!idmavad
Adodc2.Recordset.Fields!idmade = Combo4.List(Combo2.ListIndex)
Adodc2.Recordset.Fields!meghdar = Text1(5).Text
Adodc2.Recordset.Fields!j5 = 0
Adodc2.Recordset.Fields!qq = Combo2.Text
Adodc2.Recordset.Update
Label2(13).Caption = Val(Label2(13).Caption) + (Val(Text1(5).Text) * Val(Combo5.List(Combo2.ListIndex)))
Combo2.SetFocus
MsgBox "ÇØáÇÚÇÊ ÈÇ ãæÝÞíÊ ËÈÊ ÔÏ", vbInformation + vbMsgBoxRight, ""

End Sub

Private Sub DataGrid1_Click()
'On Error Resume Next
If Adodc1.Recordset.RecordCount > 0 Then
  commove = 1
  
  Adodc2.ConnectionString = Form3.Text10.Text
  Adodc2.CommandType = adCmdUnknown
  Adodc2.RecordSource = "select sum(meghdar) as amin12 from g_ozanunder where (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmavad)) + ")"
  Adodc2.Refresh
  If IsNull(Adodc2.Recordset.Fields!amin12) Then
    Label2(13).Caption = 0
  Else
    Label2(13).Caption = Adodc2.Recordset.Fields!amin12
  End If
  
  Adodc2.ConnectionString = Form3.Text10.Text
  Adodc2.CommandType = adCmdUnknown
  Adodc2.RecordSource = "select * from g_ozanunder where (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmavad)) + ")"
  Adodc2.Refresh
  If Adodc2.Recordset.RecordCount > 0 Then
    Adodc2.Recordset.MoveFirst
    Do
      Form8.Adodc1.Recordset.Find "idmavad=" + Adodc2.Recordset.Fields!idmade, , adSearchForward, 1
      DataGrid2.Col = 0
      DataGrid2.Text = Form8.Adodc1.Recordset.Fields!mavad
      Adodc2.Recordset.MoveNext
    Loop Until Adodc2.Recordset.EOF = True
  End If
  DataGrid2.Refresh
End If
End Sub

Private Sub DataGrid2_BeforeDelete(Cancel As Integer)
Label2(13).Caption = Val(Label2(13).Caption) - Val(Adodc2.Recordset.Fields!meghdar)
End Sub

Private Sub Form_Activate()
commove = 0

Form8.Adodc1.ConnectionString = Form3.Text10.Text
Form8.Adodc1.CommandType = adCmdUnknown
Form8.Adodc1.RecordSource = "select * from g_infomavad"
Form8.Adodc1.Refresh

If Form8.Adodc1.Recordset.RecordCount > 0 Then
  Combo2.Clear
  Combo4.Clear
  Combo5.Clear
  Form8.Adodc1.Recordset.Sort = "idmavad"
  Form8.Adodc1.Recordset.MoveFirst
  Do
    Combo2.AddItem Form8.Adodc1.Recordset.Fields!mavad
    Combo4.AddItem Form8.Adodc1.Recordset.Fields!idmavad
    Combo5.AddItem Form8.Adodc1.Recordset.Fields!zarib
    Form8.Adodc1.Recordset.MoveNext
  Loop Until Form8.Adodc1.Recordset.EOF = True
End If

Adodc2.ConnectionString = Form3.Text10.Text
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "select * from g_ozanunder WHERE idmahsol=0"
Adodc2.Refresh

Label2(13).Caption = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
Me.Hide
End Sub

Private Sub List4_DblClick()
If List4.ListIndex <> -1 Then
  q = Adodc3.Recordset.RecordCount
  Adodc3.Refresh
  Adodc3.Recordset.AddNew
  Adodc3.Recordset.Fields!idmahsol = Adodc1.Recordset.Fields!idmahsol
  Adodc3.Recordset.Fields!rad = Adodc1.Recordset.Fields!rad
  Adodc3.Recordset.Fields!rad1 = q + 1
  Adodc3.Recordset.Fields!Name = List4.List(List4.ListIndex)
  List4.RemoveItem (List4.ListIndex)
  Adodc3.Recordset.Update
End If
End Sub

Private Sub mnumavadmasraf_Click()
Form8.Show
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  If Index = 5 Then Command3.SetFocus
End If
End Sub
