VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form27 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÑÏÔ ãæÇÏ Çæáíå"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   13230
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form27.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   13230
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3600
      Top             =   2880
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
      Caption         =   "ËÈÊ"
      Height          =   495
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "äÞá ÇÒ æÇÍÏ ÑÇäæá ÓÇÒí Øí ÏæÑå"
      Height          =   1695
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   960
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   4
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   5
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÞÏÇÑ"
         Height          =   495
         Index           =   3
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÈáÛ"
         Height          =   495
         Index           =   6
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "ÎÑíÏ Øí ÏæÑå"
      Height          =   1695
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   960
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   3
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   2
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÈáÛ"
         Height          =   495
         Index           =   2
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÞÏÇÑ"
         Height          =   495
         Index           =   1
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "ãæÌæÏí ÇíÇä ÏæÑå "
      Height          =   1695
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   3360
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   16
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   15
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÈáÛ"
         Height          =   495
         Index           =   16
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÞÏÇÑ"
         Height          =   495
         Index           =   15
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "ÖÇíÚÇÊ"
      Height          =   1695
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   3360
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   14
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   13
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÞÏÇÑ"
         Height          =   495
         Index           =   14
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÈáÛ"
         Height          =   495
         Index           =   13
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "ÝÑæÔ Øí ÏæÑå"
      Height          =   1695
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3360
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   12
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   11
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÈáÛ"
         Height          =   495
         Index           =   12
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÞÏÇÑ"
         Height          =   495
         Index           =   11
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "ãÕÑÝ Øí ÏæÑå"
      Height          =   1695
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3360
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   10
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   9
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÞÏÇÑ"
         Height          =   495
         Index           =   10
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÈáÛ"
         Height          =   495
         Index           =   9
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "ãæÌæÏí ãæÇÏ ÂãÇÏå ãÕÑÝ"
      Height          =   2295
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   960
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   8
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   7
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   6
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÈáÛ"
         Height          =   495
         Index           =   8
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Ýí"
         Height          =   495
         Index           =   7
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÞÏÇÑ"
         Height          =   495
         Index           =   5
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ãæÌæÏí Çæá ÏæÑå"
      Height          =   1695
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   0
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   495
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÈáÛ"
         Height          =   495
         Index           =   0
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÞÏÇÑ"
         Height          =   495
         Index           =   4
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      Caption         =   "ãæÇÏ Çæáíå ˜ã˜í"
      Height          =   495
      Left            =   9000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "ãæÇÏ Çæáíå ãÕÑÝí"
      Height          =   495
      Left            =   10800
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   465
      Left            =   5880
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.ComboBox Combo3 
      Height          =   465
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "Combo3"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "äÇã ãÇÏå"
      Height          =   495
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu mnutarif 
      Caption         =   "ÊÚÇÑíÝ"
      Begin VB.Menu mnumavadkomaki 
         Caption         =   "ãæÇÏ Çæáíå ˜ã˜í"
      End
   End
End
Attribute VB_Name = "Form27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
For q = 0 To 16
  Text1(q).Text = 0
Next q
If Option1.Value = True Then Adodc1.RecordSource = "Select * From ghardeshmavad Where (nomade=1) and (idmade=" + Combo3.List(Combo1.ListIndex) + ")"
If Option2.Value = True Then Adodc1.RecordSource = "Select * From ghardeshmavad Where (nomade=2) and (idmade=" + Combo3.List(Combo1.ListIndex) + ")"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
  Text1(0).Text = Adodc1.Recordset.Fields!moneyonedoremeghdar
  Text1(1).Text = Adodc1.Recordset.Fields!moneyonedoremablagh
  Text1(2).Text = Adodc1.Recordset.Fields!kharidteydoremeghdar
  Text1(3).Text = Adodc1.Recordset.Fields!kharidteydoremablagh
  Text1(4).Text = Adodc1.Recordset.Fields!naghlazgeranolmeghdar
  Text1(5).Text = Adodc1.Recordset.Fields!naghlazgeranolmablagh
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

If Option1.Value = True Then Adodc1.RecordSource = "Select * From ghardeshmavad Where (nomade=1) and (idmade=" + Combo3.List(Combo1.ListIndex) + ")"
If Option2.Value = True Then Adodc1.RecordSource = "Select * From ghardeshmavad Where (nomade=2) and (idmade=" + Combo3.List(Combo1.ListIndex) + ")"
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
  Adodc1.Recordset.Fields!naghlazgeranolmeghdar = Text1(4).Text
  Adodc1.Recordset.Fields!naghlazgeranolmablagh = Text1(5).Text
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
  Adodc1.Recordset.Fields!naghlazgeranolmeghdar = Text1(4).Text
  Adodc1.Recordset.Fields!naghlazgeranolmablagh = Text1(5).Text
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

Private Sub Form_Activate()
Call Option1_Click
For q = 0 To 16
  Text1(q).Text = 0
Next q
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
End Sub

Private Sub mnumavadkomaki_Click()
Form8.Show
End Sub

Private Sub MSHFlexGrid1_Click()

End Sub

Private Sub Option1_Click()
Combo1.Clear
Combo3.Clear
Form4.Adodc1.Recordset.Sort = "idmavad"
Form4.Adodc1.Recordset.MoveFirst
Do
  Combo1.AddItem Form4.Adodc1.Recordset.Fields!mavad
  Combo3.AddItem Form4.Adodc1.Recordset.Fields!idmavad
  Form4.Adodc1.Recordset.MoveNext
Loop Until Form4.Adodc1.Recordset.EOF = True
End Sub

Private Sub Option2_Click()
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

Private Sub Text1_Change(Index As Integer)
Text1(6).Text = Round(Val(Text1(4).Text) + Val(Text1(2).Text) + Val(Text1(0).Text))
Text1(8).Text = Round(Val(Text1(5).Text) + Val(Text1(3).Text) + Val(Text1(1).Text))
If Text1(6).Text = 0 Then
  Text1(7).Text = 0
Else
  Text1(7).Text = Round(Val(Text1(8).Text) / Val(Text1(6).Text))
End If
Text1(9).Text = Round(Val(Text1(6).Text) - (Val(Text1(11).Text) + Val(Text1(13).Text) + Val(Text1(15).Text)))

Text1(10).Text = Round(Val(Text1(7).Text) * Val(Text1(9).Text))
Text1(12).Text = Round(Val(Text1(7).Text) * Val(Text1(11).Text))
Text1(14).Text = Round(Val(Text1(7).Text) * Val(Text1(13).Text))
Text1(16).Text = Round(Val(Text1(8).Text) - (Val(Text1(14).Text) + Val(Text1(12).Text) + Val(Text1(10).Text)))
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
