VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ê—œ‘ „Ê«œ «Ê·ÌÂ"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9840
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   465
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Text            =   "Combo3"
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ç«Å"
      Height          =   735
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   5160
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1560
      Top             =   3840
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
      CommandType     =   8
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
      RecordSource    =   ""
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
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      Caption         =   "ò«·«"
      Height          =   495
      Left            =   6600
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   360
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   3840
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
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.CommandButton Command1 
      Caption         =   "À» "
      Height          =   975
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Frame Frame4 
      Caption         =   "‰ﬁ· »Â  Ê·Ìœ "
      Height          =   1695
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   2040
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   9
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   10
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„ﬁœ«—"
         Height          =   495
         Index           =   3
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€"
         Height          =   495
         Index           =   6
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ê«—œÂ ÿÌ œÊ—Â"
      Height          =   1695
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   120
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   3
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   2
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€"
         Height          =   495
         Index           =   2
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„ﬁœ«—"
         Height          =   495
         Index           =   1
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "„ÊÃÊœÌ Å«Ì«‰ œÊ—Â"
      Height          =   1695
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   4200
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   16
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   15
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€"
         Height          =   495
         Index           =   16
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„ﬁœ«—"
         Height          =   495
         Index           =   15
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "ò‰ —· - ÷«Ì⁄«   Ê·Ìœ"
      Height          =   1695
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   4200
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   14
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   13
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„ﬁœ«—"
         Height          =   495
         Index           =   14
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€"
         Height          =   495
         Index           =   13
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "›—Ê‘ ÿÌ œÊ—Â"
      Height          =   1695
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2040
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   12
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   11
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€"
         Height          =   495
         Index           =   12
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„ﬁœ«—"
         Height          =   495
         Index           =   11
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "¬„«œÂ »—«Ì „’—›"
      Height          =   2295
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1920
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   8
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   7
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   6
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€"
         Height          =   495
         Index           =   8
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "›Ì"
         Height          =   495
         Index           =   7
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„ﬁœ«—"
         Height          =   495
         Index           =   5
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "„ÊÃÊœÌ «Ê· œÊ—Â"
      Height          =   1695
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   0
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€"
         Height          =   495
         Index           =   0
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„ﬁœ«—"
         Height          =   495
         Index           =   4
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "„Ê«œ «Ê·ÌÂ „’—›Ì"
      Height          =   495
      Left            =   8160
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   465
      Left            =   6600
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.ComboBox Combo3 
      Height          =   465
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "Combo3"
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   2880
      Top             =   3840
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
      RecordSource    =   "p_g_gardeshmavad"
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "‰«„ „«œÂ"
      Height          =   495
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
On Error Resume Next
For q = 0 To 16
  Text1(q).Text = 0
Next q
If Option1.Value = True Then Adodc1.RecordSource = "Select * From g_gardeshmavad Where (nomade=1) and (idmade=" + Combo3.List(Combo1.ListIndex) + ")"
If Option2.Value = True Then Adodc1.RecordSource = "Select * From g_gardeshmavad Where (nomade=2) and (idmade=" + Combo3.List(Combo1.ListIndex) + ")"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
  Text1(0).Text = Adodc1.Recordset.Fields!moneyonedoremeghdar
  Text1(1).Text = Adodc1.Recordset.Fields!moneyonedoremablagh
  Text1(2).Text = Adodc1.Recordset.Fields!kharidteydoremeghdar
  Text1(3).Text = Adodc1.Recordset.Fields!kharidteydoremablagh
  Text1(6).Text = Adodc1.Recordset.Fields!mojodiamademasrafmeghdar
  Text1(7).Text = Adodc1.Recordset.Fields!mojodiamademasraffi
  Text1(8).Text = Adodc1.Recordset.Fields!mojodiamademasrafmablagh
  Text1(9).Text = Adodc1.Recordset.Fields!masrafteydoremeghdar
  Text1(10).Text = Adodc1.Recordset.Fields!masrafteydoremablagh
  Text1(11).Text = Adodc1.Recordset.Fields!foroshteydoremeghdar
  Text1(12).Text = Adodc1.Recordset.Fields!foroshteydoremablagh
  Text1(13).Text = Adodc1.Recordset.Fields!zayeatmeghdar
  Text1(14).Text = Adodc1.Recordset.Fields!zayeatmablagh
  Text1(15).Text = Adodc1.Recordset.Fields!mojodipayandoremeghdar
  Text1(16).Text = Adodc1.Recordset.Fields!mojodipayandoremablagh
End If
Adodc1.RecordSource = "Select * From ghardeshmavad "
Adodc1.Refresh
End Sub

Private Sub Command1_Click()
Adodc1.RecordSource = "Select * From ghardeshmavad"
Adodc1.Refresh

If Option1.Value = True Then Adodc1.RecordSource = "Select * From g_gardeshmavad Where (nomade=1) and (idmade=" + Combo3.List(Combo1.ListIndex) + ")"
If Option2.Value = True Then Adodc1.RecordSource = "Select * From g_gardeshmavad Where (nomade=2) and (idmade=" + Combo3.List(Combo1.ListIndex) + ")"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount = 0 Then
  Adodc1.Refresh
  Adodc1.Recordset.AddNew
  Adodc1.Recordset.Fields!idmade = Combo3.List(Combo1.ListIndex)
  If Option1.Value = True Then Adodc1.Recordset.Fields!nomade = 1
  If Option2.Value = True Then Adodc1.Recordset.Fields!nomade = 2
  Adodc1.Recordset.Fields!moneyonedoremeghdar = Text1(0).Text
  Adodc1.Recordset.Fields!moneyonedoremablagh = Text1(1).Text
  Adodc1.Recordset.Fields!kharidteydoremeghdar = Text1(2).Text
  Adodc1.Recordset.Fields!kharidteydoremablagh = Text1(3).Text
  Adodc1.Recordset.Fields!naghlazgeranolmeghdar = "-"
  Adodc1.Recordset.Fields!naghlazgeranolmablagh = "-"
  Adodc1.Recordset.Fields!mojodiamademasrafmeghdar = Text1(6).Text
  Adodc1.Recordset.Fields!mojodiamademasraffi = Text1(7).Text
  Adodc1.Recordset.Fields!mojodiamademasrafmablagh = Text1(8).Text
  Adodc1.Recordset.Fields!masrafteydoremeghdar = Text1(9).Text
  Adodc1.Recordset.Fields!masrafteydoremablagh = Text1(10).Text
  Adodc1.Recordset.Fields!foroshteydoremeghdar = Text1(11).Text
  Adodc1.Recordset.Fields!foroshteydoremablagh = Text1(12).Text
  Adodc1.Recordset.Fields!zayeatmeghdar = Text1(13).Text
  Adodc1.Recordset.Fields!zayeatmablagh = Text1(14).Text
  Adodc1.Recordset.Fields!mojodipayandoremeghdar = Text1(15).Text
  Adodc1.Recordset.Fields!mojodipayandoremablagh = Text1(16).Text
  Adodc1.Recordset.Update
  Adodc1.Refresh
End If

If Adodc1.Recordset.RecordCount > 0 Then
  Adodc1.Recordset.Fields!moneyonedoremeghdar = Text1(0).Text
  Adodc1.Recordset.Fields!moneyonedoremablagh = Text1(1).Text
  Adodc1.Recordset.Fields!kharidteydoremeghdar = Text1(2).Text
  Adodc1.Recordset.Fields!kharidteydoremablagh = Text1(3).Text
  Adodc1.Recordset.Fields!naghlazgeranolmeghdar = "-"
  Adodc1.Recordset.Fields!naghlazgeranolmablagh = "-"
  Adodc1.Recordset.Fields!mojodiamademasrafmeghdar = Text1(6).Text
  Adodc1.Recordset.Fields!mojodiamademasraffi = Text1(7).Text
  Adodc1.Recordset.Fields!mojodiamademasrafmablagh = Text1(8).Text
  Adodc1.Recordset.Fields!masrafteydoremeghdar = Text1(9).Text
  Adodc1.Recordset.Fields!masrafteydoremablagh = Text1(10).Text
  Adodc1.Recordset.Fields!foroshteydoremeghdar = Text1(11).Text
  Adodc1.Recordset.Fields!foroshteydoremablagh = Text1(12).Text
  Adodc1.Recordset.Fields!zayeatmeghdar = Text1(13).Text
  Adodc1.Recordset.Fields!zayeatmablagh = Text1(14).Text
  Adodc1.Recordset.Fields!mojodipayandoremeghdar = Text1(15).Text
  Adodc1.Recordset.Fields!mojodipayandoremablagh = Text1(16).Text
  Adodc1.Recordset.Update
  Adodc1.Refresh
End If

End Sub

Private Sub Command2_Click()
Dim db1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
db1.Open Form3.Text10.Text
rs1.Open "DELETE FROM p_g_gardeshmavad", db1
db1.Close
Adodc2.RecordSource = "SELECT * FROM g_gardeshmavad"
Adodc2.Refresh
Adodc2.Recordset.MoveFirst
Do
  Adodc3.Refresh
  Adodc3.Recordset.AddNew
  Adodc3.Recordset.Fields!idmade = Adodc2.Recordset.Fields!idmade
  Adodc3.Recordset.Fields!nomade = Adodc2.Recordset.Fields!nomade
  Adodc3.Recordset.Fields!moneyonedoremeghdar = Adodc2.Recordset.Fields!moneyonedoremeghdar
  Adodc3.Recordset.Fields!moneyonedoremablagh = Adodc2.Recordset.Fields!moneyonedoremablagh
  Adodc3.Recordset.Fields!kharidteydoremeghdar = Adodc2.Recordset.Fields!kharidteydoremeghdar
  Adodc3.Recordset.Fields!kharidteydoremablagh = Adodc2.Recordset.Fields!kharidteydoremablagh
  Adodc3.Recordset.Fields!naghlazgeranolmeghdar = Val(Adodc2.Recordset.Fields!naghlazgeranolmeghdar)
  Adodc3.Recordset.Fields!naghlazgeranolmablagh = Val(Adodc2.Recordset.Fields!naghlazgeranolmablagh)
  Adodc3.Recordset.Fields!mojodiamademasrafmeghdar = Adodc2.Recordset.Fields!mojodiamademasrafmeghdar
  Adodc3.Recordset.Fields!mojodiamademasraffi = Adodc2.Recordset.Fields!mojodiamademasraffi
  Adodc3.Recordset.Fields!mojodiamademasrafmablagh = Adodc2.Recordset.Fields!mojodiamademasrafmablagh
  Adodc3.Recordset.Fields!masrafteydoremeghdar = Adodc2.Recordset.Fields!masrafteydoremeghdar
  Adodc3.Recordset.Fields!masrafteydoremablagh = Adodc2.Recordset.Fields!masrafteydoremablagh
  Adodc3.Recordset.Fields!foroshteydoremeghdar = Adodc2.Recordset.Fields!foroshteydoremeghdar
  Adodc3.Recordset.Fields!foroshteydoremablagh = Adodc2.Recordset.Fields!foroshteydoremablagh
  Adodc3.Recordset.Fields!zayeatmeghdar = Adodc2.Recordset.Fields!zayeatmeghdar
  Adodc3.Recordset.Fields!zayeatmablagh = Adodc2.Recordset.Fields!zayeatmablagh
  Adodc3.Recordset.Fields!mojodipayandoremeghdar = Adodc2.Recordset.Fields!mojodipayandoremeghdar
  Adodc3.Recordset.Fields!mojodipayandoremablagh = Adodc2.Recordset.Fields!mojodipayandoremablagh
  If Adodc2.Recordset.Fields!nomade = 1 Then
    Form8.Adodc1.Recordset.Find "idmavad=" + Trim(Str(Adodc2.Recordset.Fields!idmade)), , adSearchForward, 1
    Adodc3.Recordset.Fields!Name = Form8.Adodc1.Recordset.Fields!mavad
    Adodc3.Recordset.Fields!zay = "÷«Ì⁄« "
  End If
  
  If Adodc2.Recordset.Fields!nomade = 2 Then
    Form15.Adodc1.Recordset.Find "idmavad=" + Trim(Str(Adodc2.Recordset.Fields!idmade)), , adSearchForward, 1
    Adodc3.Recordset.Fields!Name = Form15.Adodc1.Recordset.Fields!mavad
    Adodc3.Recordset.Fields!zay = "«÷«›Â „’—›  Ê·Ìœ"
  End If
  
  Adodc3.Recordset.Update
  Adodc2.Recordset.MoveNext
Loop Until Adodc2.Recordset.EOF = True
Form6.Show
End Sub

Private Sub Form_Activate()
On Error Resume Next
Call Option1_Click
For q = 0 To 16
  Text1(q).Text = 0
Next q
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
End Sub

Private Sub Option1_Click()
Combo1.Clear
Combo3.Clear
Form8.Adodc1.Recordset.Sort = "idmavad"
Form8.Adodc1.Recordset.MoveFirst
Do
  Combo1.AddItem Form8.Adodc1.Recordset.Fields!mavad
  Combo3.AddItem Form8.Adodc1.Recordset.Fields!idmavad
  Form8.Adodc1.Recordset.MoveNext
Loop Until Form8.Adodc1.Recordset.EOF = True
End Sub

Private Sub Option2_Click()
Combo1.Clear
Combo2.Clear
Combo3.Clear

Form15.Adodc1.Recordset.Sort = "idmavad"
Form15.Adodc1.Recordset.MoveFirst
Do
  Combo1.AddItem Form15.Adodc1.Recordset.Fields!mavad
  Combo2.AddItem Form15.Adodc1.Recordset.Fields!mastebach
  Combo3.AddItem Form15.Adodc1.Recordset.Fields!idmavad
  Form15.Adodc1.Recordset.MoveNext
Loop Until Form15.Adodc1.Recordset.EOF = True
End Sub

Private Sub Text1_Change(Index As Integer)

If (Option1.Value = True) And (Combo1.ListIndex <> -1) Then
  Text1(6).Text = Round(Val(Text1(2).Text) + Val(Text1(0).Text))
  Text1(8).Text = Round(Val(Text1(3).Text) + Val(Text1(1).Text))
  If Text1(6).Text = 0 Then
    Text1(7).Text = 0
  Else
    Text1(7).Text = Round(Val(Text1(8).Text) / Val(Text1(6).Text))
  End If

  Text1(9).Text = Round(Val(Text1(6).Text) - (Val(Text1(11).Text) + Val(Text1(13).Text) + Val(Text1(15).Text)))

  If Val(Text1(6).Text) <> 0 Then
    Text1(10).Text = Round(Val(Text1(8).Text) / Val(Text1(6).Text) * Val(Text1(9).Text))
    Text1(12).Text = Round(Val(Text1(8).Text) / Val(Text1(6).Text) * Val(Text1(11).Text))
    Text1(14).Text = Round(Val(Text1(8).Text) / Val(Text1(6).Text) * Val(Text1(13).Text))
  End If
  Text1(16).Text = Round(Val(Text1(8).Text) - (Val(Text1(14).Text) + Val(Text1(12).Text) + Val(Text1(10).Text)))
End If


If (Option2.Value = True) And (Combo1.ListIndex <> -1) Then
  If Combo2.List(Combo1.ListIndex) = 1 Then
    Form4.Adodc1.Recordset.Find "idmavad=" + Combo3.List(Combo1.ListIndex), , adSearchForward, 1
    Text1(3).Text = Form4.Adodc1.Recordset.Fields!gheymattamam
    Text1(6).Text = Round(Val(Text1(2).Text) + Val(Text1(0).Text))
    Text1(8).Text = Round(Val(Text1(3).Text) + Val(Text1(1).Text))
    
    If Text1(6).Text = 0 Then
      Text1(7).Text = 0
    Else
      Text1(7).Text = Round(Val(Text1(8).Text) / Val(Text1(6).Text))
    End If
    
    Text1(9).Text = Round(Val(Text1(6).Text) - ((Val(Text1(11).Text) + Val(Text1(13).Text) + Val(Text1(15).Text))))
    
    If Val(Text1(6).Text) <> 0 Then
      Text1(10).Text = Round(Val(Text1(8).Text) / Val(Text1(6).Text) * Val(Text1(9).Text))
      Text1(12).Text = Round(Val(Text1(8).Text) / Val(Text1(6).Text) * Val(Text1(11).Text))
      Text1(14).Text = Round(Val(Text1(8).Text) / Val(Text1(6).Text) * Val(Text1(13).Text))
      Text1(16).Text = Round(Val(Text1(8).Text) / Val(Text1(6).Text) * Val(Text1(15).Text))
    End If
  Else
    Adodc2.CommandType = adCmdUnknown
    Adodc2.RecordSource = "SELECT * FROM ghardeshmavad WHERE (nomade=1) AND (idmade=" + Combo3.List(Combo1.ListIndex) + ")"
    Adodc2.Refresh
    If (Adodc2.Recordset.RecordCount = 0) Then
      Text1(9).Text = 0
    Else
      Text1(9).Text = Round(Adodc2.Recordset.Fields!naghlazgeranolmeghdar)
    End If

    Form4.Adodc1.Recordset.Find "idmavad=" + Combo3.List(Combo1.ListIndex), , adSearchForward, 1
    If IsNull(Form4.Adodc1.Recordset.Fields!gheymattamam) = False Then
      Text1(3).Text = Form4.Adodc1.Recordset.Fields!gheymattamam
    Else
      Text1(3).Text = 0
    End If
    Text1(6).Text = Round(Val(Text1(2).Text) + Val(Text1(0).Text))
    Text1(8).Text = Round(Val(Text1(3).Text) + Val(Text1(1).Text))
    If Text1(6).Text = 0 Then
      Text1(7).Text = 0
    Else
      Text1(7).Text = Round(Val(Text1(8).Text) / Val(Text1(6).Text))
    End If
  
    If Val(Text1(6).Text) <> 0 Then
      Text1(10).Text = Round(Val(Text1(8).Text) / Val(Text1(6).Text) * Val(Text1(9).Text))
      Text1(12).Text = Round(Val(Text1(8).Text) / Val(Text1(6).Text) * Val(Text1(11).Text))
    
      Text1(13).Text = Round(Val(Text1(6).Text) - (Val(Text1(9).Text) + Val(Text1(11).Text) + Val(Text1(15).Text)))
      Text1(14).Text = Round(Val(Text1(8).Text) / Val(Text1(6).Text) * Val(Text1(13).Text))

      Text1(16).Text = Round(Val(Text1(8).Text) / Val(Text1(6).Text) * Val(Text1(15).Text))
    End If
  End If
End If

For q = 0 To 16
  If Val(Text1(Index).Text) < 0 Then
    Text1(Index).BackColor = RGB(255, 0, 0)
    Text1(Index).ForeColor = RGB(255, 255, 255)
  Else
    Text1(Index).BackColor = RGB(255, 255, 255)
    Text1(Index).ForeColor = RGB(0, 0, 0)
  End If
Next q

End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If ((KeyAscii >= 48) And (KeyAscii <= 57)) Or KeyAscii = 8 Then
Else
  KeyAscii = 0
End If
End Sub
