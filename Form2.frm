VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ãÕÑÝ ÑíÇáí ãæÇÏ ÏÑ ÊæáíÏ ÑÇäæá ÏæÑå 6"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9675
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   9675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "ÑÏÇÒÔ"
      Height          =   495
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
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
      Bindings        =   "Form2.frx":2CFA
      Height          =   5415
      Left            =   4920
      TabIndex        =   0
      Top             =   720
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   9551
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Form2.frx":2D0F
      Height          =   5415
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   9551
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
         DataField       =   "qq"
         Caption         =   "äÇã ãÇÏå"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "j6"
         Caption         =   "ãÞÏÇÑ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2009.764
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1440
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc Adodc3 
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
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Command1_Click()
If Adodc1.Recordset.RecordCount > 0 Then
  ProgressBar1.Min = 0
  ProgressBar1.Max = Adodc1.Recordset.RecordCount
  ProgressBar1.Value = 0
  Adodc1.Recordset.MoveFirst
  Do
    ProgressBar1.Value = ProgressBar1.Value + 1
    Adodc2.ConnectionString = Form3.Text10.Text
    Adodc2.CommandType = adCmdUnknown
    Adodc2.RecordSource = "select * from g_ozanunder where (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmavad)) + ")"
    Adodc2.Refresh

    If Adodc2.Recordset.RecordCount > 0 Then
      Adodc2.Recordset.MoveFirst
      Do
        Adodc3.ConnectionString = Form3.Text10.Text
        Adodc3.CommandType = adCmdUnknown
        Adodc3.RecordSource = "select sum(j5) as j51 from g_ozanunder where (idmade='" + Trim(Str(Adodc2.Recordset.Fields!idmade)) + "')"
        Adodc3.Refresh
        w = Adodc3.Recordset.Fields!j51
        
        Form7.Adodc1.CommandType = adCmdUnknown
        Form7.Adodc1.RecordSource = "Select * From g_gardeshmavad Where (nomade=1) and (idmade=" + Trim(Str(Adodc2.Recordset.Fields!idmade)) + ")"
        Form7.Adodc1.Refresh
        If Form7.Adodc1.Recordset.RecordCount > 0 Then
          q = Form7.Adodc1.Recordset.Fields!masrafteydoremablagh
        Else
          q = 0
        End If
'        MsgBox q
'        MsgBox w
'        MsgBox Adodc2.Recordset.Fields!j5
        If w <> 0 Then
          Adodc2.Recordset.Fields!j6 = Round((q / w) * Adodc2.Recordset.Fields!j5)
        Else
          Adodc2.Recordset.Fields!j6 = 0
        End If
        Adodc2.Recordset.Update
        Adodc2.Recordset.MoveNext
      Loop Until Adodc2.Recordset.EOF = True
    End If
    Adodc1.Recordset.MoveNext
  Loop Until Adodc1.Recordset.EOF = True
End If
End Sub

Private Sub DataGrid1_Click()
If Adodc1.Recordset.RecordCount > 0 Then
  commove = 1
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

Private Sub Form_Activate()
commove = 0

Form8.Adodc1.ConnectionString = Form3.Text10.Text
Form8.Adodc1.CommandType = adCmdUnknown
Form8.Adodc1.RecordSource = "select * from g_infomavad"
Form8.Adodc1.Refresh

Adodc1.ConnectionString = Form3.Text10.Text
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from infomavad where (nogra='1') ORDER BY idmavad"
Adodc1.Refresh

Adodc2.ConnectionString = Form3.Text10.Text
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "select * from g_ozanunder WHERE idmahsol=0"
Adodc2.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
Me.Hide
End Sub

