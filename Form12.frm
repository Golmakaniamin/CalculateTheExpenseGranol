VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form12 
   Caption         =   "«‰Õ—«›"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form12"
   RightToLeft     =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "ç«Å"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Å—œ«“‘"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form12.frx":0000
      Height          =   6735
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   11880
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
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "idmade"
         Caption         =   "òœ „«œÂ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "qq"
         Caption         =   "‰«„ „«œÂ"
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
      BeginProperty Column02 
         DataField       =   "m1"
         Caption         =   "„ﬁœ«— „Õ«”»Â"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "r1"
         Caption         =   "—Ì«· „Õ«”»Â"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "m2"
         Caption         =   "„ﬁœ«— „’—›"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "r2"
         Caption         =   "—Ì«· „’—›"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "m3"
         Caption         =   "„ﬁœ«— «‰Õ—«›"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "r3"
         Caption         =   "—Ì«· «‰Õ—«›"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "e3"
         Caption         =   "—Ì«· „Õ«”»Â ‘œÂ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column08 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   1680
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
      RecordSource    =   "Enheraf"
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
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim db1 As New ADODB.Connection
Dim db2 As New ADODB.Connection
Dim rs(3) As New ADODB.Recordset

db1.Open Form3.Text10.Text
  rs(0).Open "DELETE FROM Enheraf", db1
db1.Close


db1.Open Form3.Text10.Text
  rs(0).Open "SELECT idmade, qq, SUM(j5) As rssum1, SUM(j6) As rssum2 FROM g_ozanunder GROUP BY idmade,qq", db1
    rs(0).MoveFirst
    Do
      Adodc1.Refresh
      Adodc1.Recordset.AddNew
      Adodc1.Recordset.Fields!idmade = rs(0).Fields!idmade
      Adodc1.Recordset.Fields!qq = rs(0).Fields!qq
      Adodc1.Recordset.Fields!m1 = rs(0).Fields!rssum1
      Adodc1.Recordset.Fields!r1 = rs(0).Fields!rssum2
      rs(1).Open "SELECT * FROM g_gardeshmavad WHERE (nomade=1) AND (idmade=" + rs(0).Fields!idmade + ")", db1
        Adodc1.Recordset.Fields!m2 = rs(1).Fields!masrafteydoremeghdar
        Adodc1.Recordset.Fields!r2 = rs(1).Fields!masrafteydoremablagh
      rs(1).Close
      Adodc1.Recordset.Fields!m3 = Adodc1.Recordset.Fields!m1 - Adodc1.Recordset.Fields!m2
      Adodc1.Recordset.Fields!r3 = Adodc1.Recordset.Fields!r1 - Adodc1.Recordset.Fields!r2
      If ((Adodc1.Recordset.Fields!r2 <> 0) And (Adodc1.Recordset.Fields!m2 <> 0)) Then
        Adodc1.Recordset.Fields!e3 = (Adodc1.Recordset.Fields!r2 / Adodc1.Recordset.Fields!m2) * Adodc1.Recordset.Fields!m3
      Else
        Adodc1.Recordset.Fields!e3 = 0
      End If
      Adodc1.Recordset.Update
      rs(0).MoveNext
    Loop Until rs(0).EOF = True
  rs(0).Close
db1.Close

Adodc1.ConnectionString = Form3.Text10.Text
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "SELECT * FROM Enheraf ORDER BY idmade"
Adodc1.Refresh
DataGrid1.Refresh


db1.Open Form3.Text10.Text
  rs(0).Open "SELECT SUM(m1) As m11,SUM(r1) As r11,SUM(m2) As m21,SUM(r2) As r21,SUM(m3) As m31,SUM(r3) As r31,SUM(e3) As e31 FROM Enheraf", db1
     Adodc1.Refresh
      Adodc1.Recordset.AddNew
      Adodc1.Recordset.Fields!idmade = 1000
      Adodc1.Recordset.Fields!qq = "Ã„⁄ ò· :"
      Adodc1.Recordset.Fields!m1 = rs(0).Fields!m11
      Adodc1.Recordset.Fields!r1 = rs(0).Fields!r11
      Adodc1.Recordset.Fields!m2 = rs(0).Fields!m21
      Adodc1.Recordset.Fields!r2 = rs(0).Fields!r21
      Adodc1.Recordset.Fields!m3 = rs(0).Fields!m31
      Adodc1.Recordset.Fields!r3 = rs(0).Fields!r31
      Adodc1.Recordset.Fields!e3 = rs(0).Fields!e31
      Adodc1.Recordset.Update
  rs(0).Close
db1.Close

Adodc1.ConnectionString = Form3.Text10.Text
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "SELECT * FROM Enheraf ORDER BY idmade"
Adodc1.Refresh
DataGrid1.Refresh

Adodc1.Refresh
DataGrid1.Refresh

Adodc1.Refresh
DataGrid1.Refresh

Adodc1.Refresh
DataGrid1.Refresh

Adodc1.Refresh
DataGrid1.Refresh
End Sub

Private Sub Command2_Click()
Form13.Show
End Sub

Private Sub Form_Activate()
Adodc1.ConnectionString = Form3.Text10.Text
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "SELECT * FROM Enheraf"
Adodc1.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
Me.Hide
End Sub
