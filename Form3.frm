VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{DF1D4B1E-D56E-4A40-BA98-2CC06080E796}#1.0#0"; "Tiny.ocx"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "’›ÕÂ «’·Ì ê—«‰Ê· ”«“Ì"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":2CFA
   RightToLeft     =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   1125
      Left            =   360
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin TINYLib.Tiny Tiny1 
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   1296
      _StockProps     =   0
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   6360
      Visible         =   0   'False
      Width           =   4455
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP2 
      Height          =   8115
      Left            =   -240
      TabIndex        =   2
      Top             =   -240
      Visible         =   0   'False
      Width           =   9960
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   -1  'True
      uiMode          =   "mini"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   17568
      _cy             =   14314
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP 
      Height          =   8115
      Left            =   -240
      TabIndex        =   1
      Top             =   -240
      Width           =   9960
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   -1  'True
      uiMode          =   "mini"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   17568
      _cy             =   14314
   End
   Begin VB.Menu mnugranol 
      Caption         =   "ê—«‰Ê· ”«“Ì"
      Begin VB.Menu mnuozang 
         Caption         =   "«Ê“«‰"
      End
      Begin VB.Menu mnugardeshg 
         Caption         =   "ê—œ‘ „Ê«œ «Ê·ÌÂ"
      End
      Begin VB.Menu mnu1 
         Caption         =   "„’—› „ﬁœ«—Ì „Ê«œ œ—  Ê·Ìœ ê—«‰Ê· œÊ—Â 5"
      End
      Begin VB.Menu mnu2 
         Caption         =   "„’—› —Ì«·Ì „Ê«œ œ—  Ê·Ìœ ê—«‰Ê· œÊ—Â 6"
      End
      Begin VB.Menu mnu3 
         Caption         =   "„’—› —Ì«·Ì  „Ê«œ œ—  Ê·Ìœ ê—«‰Ê· œÊ—Â 3"
      End
      Begin VB.Menu mnuprogress 
         Caption         =   "Å—œ«“‘"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub Form_Initialize()
 InitCommonControls
End Sub

Private Sub Form_Activate()
DoEvents

Text10.Text = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + "D:\PraticGroup\End Of Money To Account" + "\1.mdb" + ";Persist Security Info=False"
DoEvents

Form7.Adodc1.ConnectionString = Form3.Text10.Text
Form7.Adodc1.CommandType = adCmdUnknown
Form7.Adodc1.RecordSource = "select * from g_gardeshmavad"
Form7.Adodc1.Refresh

Form8.Adodc1.ConnectionString = Form3.Text10.Text
Form8.Adodc1.CommandType = adCmdUnknown
Form8.Adodc1.RecordSource = "select * from g_infomavad"
Form8.Adodc1.Refresh

Form15.Adodc1.ConnectionString = Form3.Text10.Text
Form15.Adodc1.CommandType = adCmdUnknown
Form15.Adodc1.RecordSource = "select * from infomavad where (nogra='1') ORDER BY idmavad"
Form15.Adodc1.Refresh

Form1.Adodc1.ConnectionString = Form3.Text10.Text
Form1.Adodc1.CommandType = adCmdUnknown
Form1.Adodc1.RecordSource = "select * from infomavad ORDER BY idmavad"
Form1.Adodc1.Refresh

Form2.Adodc1.ConnectionString = Form3.Text10.Text
Form2.Adodc1.CommandType = adCmdUnknown
Form2.Adodc1.RecordSource = "select * from infomavad ORDER BY idmavad"
Form2.Adodc1.Refresh

Form15.Adodc2.ConnectionString = Form3.Text10.Text
Form15.Adodc2.CommandType = adCmdUnknown
Form15.Adodc2.RecordSource = "select * from g_ozanunder"
Form15.Adodc2.Refresh

End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
  MsgBox "‰—„ «›“«— œ— Õ«·  «Ã—« „Ì »«‘œ", vbCritical + vbMsgBoxRight, ""
  End
End If

WMP.URL = App.Path + "/1.avi"
WMP.Controls.play
DoEvents

'Tiny1.Initialize = True
'If Tiny1.TinyErrCode = 0 Then
'  Tiny1.UserPassWord = "61F9F7776F8AAFFCC29D6C8DE83A1C1"
'  Tiny1.SpecialID = "v25f192510******"
'  Tiny1.ShowTinyInfo = True
'  DoEvents
'  If Tiny1.TinyErrCode = 0 Then
'    If Tiny1.DataPartition = "PraticGroup" Then
'      DoEvents
'      If Tiny1.SerialNumber = "2019-8805-1157" Then
'        DoEvents
'      Else
'        MsgBox "⁄œ„ ‘‰«”«ÌÌ „ÕÌÿ", vbCritical + vbMsgBoxRight, ""
'        End
'      End If
'    Else
'      MsgBox "⁄œ„ ‘‰«”«ÌÌ „ÕÌÿ", vbCritical + vbMsgBoxRight, ""
'      End
'    End If
'  Else
'    MsgBox "⁄œ„ ‘‰«”«ÌÌ „ÕÌÿ", vbCritical + vbMsgBoxRight, ""
'    End
'  End If
'  Tiny1.ShowTinyInfo = False
'Else
'  If Tiny1.TinyErrCode = 1 Then
'    MsgBox "⁄œ„ ‘‰«”«ÌÌ ﬁ›·", vbCritical + vbMsgBoxRight, ""
'  End If
'  End
'End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
File1.Path = Left(App.Path, 3)
File1.Pattern = "*.tmp"
For q = 0 To File1.ListCount - 1
  Kill Left(App.Path, 3) + File1.List(q)
Next q

File1.Path = App.Path + "\"
File1.Pattern = "*.tmp"
For q = 0 To File1.ListCount - 1
  Kill App.Path + "\" + File1.List(q)
Next q

End
End Sub

Private Sub mnu1_Click()
Form1.Show
Me.Hide
End Sub

Private Sub mnu2_Click()
Form2.Show
Me.Hide
End Sub

Private Sub mnu3_Click()
Form4.Show
Me.Hide
End Sub

Private Sub mnugardeshg_Click()
Form7.Show
Me.Hide
End Sub

Private Sub mnuozang_Click()
Form15.Show
Me.Hide
End Sub

Private Sub mnuprogress_Click()
Form5.Show
Me.Hide
End Sub

Private Sub WMP_PlayStateChange(ByVal NewState As Long)
If NewState = 8 Then
  WMP2.Controls.stop
  WMP2.URL = App.Path + "/2.avi"
  WMP2.Controls.play
  WMP.Visible = False
  WMP2.Visible = True
End If
End Sub

Private Sub WMP2_PlayStateChange(ByVal NewState As Long)
If NewState = 1 Then WMP2.Controls.play
End Sub


