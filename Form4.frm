VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ãÕÑÝ ÑíÇáí  ãæÇÏ ÏÑ ÊæáíÏ ÑÇäæá ÏæÑå 3"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10425
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ç"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   7800
      Width           =   10095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "store2"
      DataSource      =   "Adodc2"
      Height          =   495
      Index           =   1
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "store1"
      DataSource      =   "Adodc2"
      Height          =   495
      Index           =   0
      Left            =   7800
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Text            =   "1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "store3"
      DataSource      =   "Adodc2"
      Height          =   495
      Index           =   2
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "store4"
      DataSource      =   "Adodc2"
      Height          =   495
      Index           =   3
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form4.frx":2CFA
      Height          =   6975
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   12303
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   23
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "mavad"
         Caption         =   "äÇã ãÍÕæá"
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
         DataField       =   "tolid"
         Caption         =   "ÊæáíÏ"
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
         DataField       =   "bahamavad"
         Caption         =   "ÈåÇí ãæÇÏ Çæáíå"
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
      BeginProperty Column03 
         DataField       =   "dastmozd"
         Caption         =   "ÏÓÊãÒÏ"
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
      BeginProperty Column04 
         DataField       =   "sarbar"
         Caption         =   "ÓÑÈÇÑ"
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
      BeginProperty Column05 
         DataField       =   "estelak"
         Caption         =   "ÇÓÊåáÇ˜"
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
      BeginProperty Column06 
         DataField       =   "gheymattamam"
         Caption         =   "ÞíãÊ ÊãÇã ÔÏå æÇÍÏ"
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   120
      Top             =   720
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   794
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2880
      Top             =   240
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
      RecordSource    =   "marahelnameasl"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   450
      Left            =   1440
      Top             =   720
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   794
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
      RecordSource    =   "sarbar_4"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   450
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   794
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
      RecordSource    =   "p_Gra2"
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
      Alignment       =   1  'Right Justify
      Caption         =   "ÈåÇí ÏÓÊãÒÏ"
      Height          =   495
      Index           =   4
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "ÈåÇí ãÓ"
      Height          =   495
      Index           =   0
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "ÓÑÈÇÑ"
      Height          =   495
      Index           =   1
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "ÇÓÊåáÇ˜"
      Height          =   495
      Index           =   2
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim db1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
db1.Open Form3.Text10.Text
rs1.Open "DELETE FROM p_Gra2", db1
db1.Close
Adodc1.Recordset.MoveFirst
Do
  Adodc4.Refresh
  Adodc4.Recordset.AddNew
  Adodc4.Recordset.Fields!idmavad = Adodc1.Recordset.Fields!idmavad
  Adodc4.Recordset.Fields!mavad = Adodc1.Recordset.Fields!mavad
  Adodc4.Recordset.Fields!tolid = Adodc1.Recordset.Fields!tolid
  Adodc4.Recordset.Fields!bahamavad = Adodc1.Recordset.Fields!bahamavad
  Adodc4.Recordset.Fields!dastmozd = Adodc1.Recordset.Fields!dastmozd
  Adodc4.Recordset.Fields!sarbar = Adodc1.Recordset.Fields!sarbar
  Adodc4.Recordset.Fields!estelak = Adodc1.Recordset.Fields!estelak
  Adodc4.Recordset.Fields!gheymattamam = Adodc1.Recordset.Fields!gheymattamam
  Adodc4.Recordset.Update
  Adodc1.Recordset.MoveNext
Loop Until Adodc1.Recordset.EOF = True
Form10.Show
End Sub

Public Sub Form_Activate()
On Error Resume Next
Adodc1.ConnectionString = Form3.Text10.Text
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from infomavad where (nogra='1') ORDER BY idmavad"
Adodc1.Refresh

Adodc2.ConnectionString = Form3.Text10.Text
Adodc2.CommandType = adCmdUnknown
Adodc2.RecordSource = "select * from marahelnameasl"
Adodc2.Refresh
Adodc2.Recordset.Find "name='gra'", , adSearchForward, 1

Adodc3.ConnectionString = Form3.Text10.Text
Adodc3.CommandType = adCmdUnknown
Adodc3.RecordSource = "select * from sarbar_4 WHERE (rad=998)"
Adodc3.Refresh

Form7.Adodc1.CommandType = adCmdUnknown
Form7.Adodc1.RecordSource = "Select sum(masrafteydoremablagh) as masrafteydoremablagh1 From g_gardeshmavad Where (nomade=1) "
Form7.Adodc1.Refresh
q = Form7.Adodc1.Recordset.Fields!masrafteydoremablagh1
Adodc2.Recordset.Fields!store1 = q
Adodc2.Recordset.Fields!store2 = Adodc3.Recordset.Fields!dastmozd
Adodc2.Recordset.Fields!store3 = Val(Adodc3.Recordset.Fields!sarbarvahed) + Val(Adodc3.Recordset.Fields!sarbarjazb)
Adodc2.Recordset.Fields!store4 = Adodc3.Recordset.Fields!estehlak
Adodc2.Recordset.Update


Adodc1.Recordset.MoveFirst
Do
  Form7.Adodc1.CommandType = adCmdUnknown
  Form7.Adodc1.RecordSource = "Select * From g_gardeshmavad Where (nomade=2) and (idmade=" + Trim(Str(Adodc1.Recordset.Fields!idmavad)) + ")"
  Form7.Adodc1.Refresh
  If Form7.Adodc1.Recordset.RecordCount > 0 Then
    Adodc1.Recordset.Fields!tolid = Form7.Adodc1.Recordset.Fields!kharidteydoremeghdar
    w = Adodc1.Recordset.Fields!tolid
  Else
    Adodc1.Recordset.Fields!tolid = 0
    w = Adodc1.Recordset.Fields!tolid
  End If
  
  Form15.Adodc2.ConnectionString = Form3.Text10.Text
  Form15.Adodc2.CommandType = adCmdUnknown
  Form15.Adodc2.RecordSource = "select sum(j6) as j61 from g_ozanunder where (idmahsol=" + Trim(Str(Adodc1.Recordset.Fields!idmavad)) + ")"
  Form15.Adodc2.Refresh
  If IsNull(Form15.Adodc2.Recordset.Fields!j61) Then
    Adodc1.Recordset.Fields!bahamavad = 0
  Else
    Adodc1.Recordset.Fields!bahamavad = Form15.Adodc2.Recordset.Fields!j61
  End If
  
  Form15.Adodc1.ConnectionString = Form3.Text10.Text
  Form15.Adodc1.CommandType = adCmdUnknown
  Form15.Adodc1.RecordSource = "select sum(tolid) as tolid1 from infomavad "
  Form15.Adodc1.Refresh
  q = Form15.Adodc1.Recordset.Fields!tolid1
  
  If (q <> 0) Then
    r1 = (Val(Text1(1).Text) / q * w)
    r2 = (Val(Text1(2).Text) / q * w)
    r3 = (Val(Text1(3).Text) / q * w)
  Else
    r1 = 0
    r2 = 0
    r3 = 0
  End If
  Adodc1.Recordset.Fields!dastmozd = Round(r1)
  Adodc1.Recordset.Fields!sarbar = Round(r2)
  Adodc1.Recordset.Fields!estelak = Round(r3)
  Adodc1.Recordset.Fields!gheymattamam = Round(Val(r1) + Val(r2) + Val(r3) + Val(Adodc1.Recordset.Fields!bahamavad))
  Adodc1.Recordset.Update
  
  Adodc1.Recordset.MoveNext
Loop Until Adodc1.Recordset.EOF = True
Adodc1.Recordset.Sort = "idmavad"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
Me.Hide
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  Select Case Index
    Case 0
      Text1(1).SetFocus
      
    Case 1
      Text1(2).SetFocus
      
    Case 2
      Text1(3).SetFocus
      
    Case 3
      Text1(0).SetFocus
      
  End Select
  Adodc2.Recordset.Fields!store1 = Text1(0).Text
  Adodc2.Recordset.Fields!store2 = Text1(1).Text
  Adodc2.Recordset.Fields!store3 = Text1(2).Text
  Adodc2.Recordset.Fields!store4 = Text1(3).Text
  Adodc2.Recordset.Update
End If
End Sub

