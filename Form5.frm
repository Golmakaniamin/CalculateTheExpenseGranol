VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÑÏÇÒÔ"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2205
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "ÔÑæÚ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   26.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   9975
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   1508
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As New ADODB.Connection
Dim rs(10) As New ADODB.Recordset

Private Sub Command1_Click()
Command1.Enabled = False
db1.Open Form3.Text10.Text
ProgressBar1.Min = 0
ProgressBar1.Max = 5
ProgressBar1.Value = 0

Call Form1.Command1_Click
ProgressBar1.Value = 1

Call Form2.Command1_Click
ProgressBar1.Value = 2

Call Form4.Form_Activate
ProgressBar1.Value = 3

Call Form4.Form_Activate
ProgressBar1.Value = 4

Form15.Adodc1.ConnectionString = Form3.Text10.Text
Form15.Adodc1.CommandType = adCmdUnknown
Form15.Adodc1.RecordSource = "select * from infomavad where (nogra='1') ORDER BY idmavad"
Form15.Adodc1.Refresh
Form15.Adodc1.Recordset.MoveFirst
Do
  Form7.Adodc1.RecordSource = "Select * From g_gardeshmavad Where (nomade=2) and (idmade=" + Trim(Str(Form15.Adodc1.Recordset.Fields!idmavad)) + ")"
  Form7.Adodc1.Refresh
  If Form7.Adodc1.Recordset.RecordCount > 0 Then
    rs(4).Open "SELECT * FROM infomavad WHERE (idmavad=" + Trim(Str(Form7.Adodc1.Recordset.Fields!idmade)) + ")", db1
      If rs(4).Fields!mastebach = 1 Then
        Form4.Adodc1.Recordset.Find "idmavad=" + Trim(Str(Form15.Adodc1.Recordset.Fields!idmavad)), , adSearchForward, 1
        Form7.Adodc1.Recordset.Fields!kharidteydoremablagh = Form4.Adodc1.Recordset.Fields!gheymattamam
      
        Form7.Adodc1.Recordset.Fields!mojodiamademasrafmeghdar = Round(Val(Form7.Adodc1.Recordset.Fields!kharidteydoremeghdar) + Val(Form7.Adodc1.Recordset.Fields!moneyonedoremeghdar))
        Form7.Adodc1.Recordset.Fields!mojodiamademasrafmablagh = Round(Val(Form7.Adodc1.Recordset.Fields!kharidteydoremablagh) + Val(Form7.Adodc1.Recordset.Fields!moneyonedoremablagh))
       
        If Form7.Adodc1.Recordset.Fields!mojodiamademasrafmeghdar = 0 Then
          Form7.Adodc1.Recordset.Fields!mojodiamademasraffi = 0
        Else
          Form7.Adodc1.Recordset.Fields!mojodiamademasraffi = Round(Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmeghdar))
        End If
       
        Form7.Adodc1.Recordset.Fields!masrafteydoremeghdar = Round(Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) - ((Val(Form7.Adodc1.Recordset.Fields!foroshteydoremeghdar) + Val(Form7.Adodc1.Recordset.Fields!zayeatmeghdar) + Val(Form7.Adodc1.Recordset.Fields!mojodipayandoremeghdar))))
        
        If Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) <> 0 Then
          Form7.Adodc1.Recordset.Fields!masrafteydoremablagh = Round(Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Round(Val(Form7.Adodc1.Recordset.Fields!masrafteydoremeghdar)))
          Form7.Adodc1.Recordset.Fields!foroshteydoremablagh = Round(Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Val(Form7.Adodc1.Recordset.Fields!foroshteydoremeghdar))
          Form7.Adodc1.Recordset.Fields!zayeatmablagh = Round(Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Val(Form7.Adodc1.Recordset.Fields!zayeatmeghdar))
          Form7.Adodc1.Recordset.Fields!mojodipayandoremablagh = Round(Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Val(Form7.Adodc1.Recordset.Fields!mojodipayandoremeghdar))
        Else
          Form7.Adodc1.Recordset.Fields!masrafteydoremablagh = 0
          Form7.Adodc1.Recordset.Fields!foroshteydoremablagh = 0
          Form7.Adodc1.Recordset.Fields!zayeatmablagh = 0
          Form7.Adodc1.Recordset.Fields!mojodipayandoremablagh = 0
        End If
        Form7.Adodc1.Recordset.Update
      Else
        rs(1).Open "SELECT Count(nomade) As rsnumber FROM ghardeshmavad WHERE (nomade=1) AND (idmade=" + Trim(Str(Form15.Adodc1.Recordset.Fields!idmavad)) + ")", db1
          If rs(1).Fields!rsnumber > 0 Then
            rs(0).Open "SELECT * FROM ghardeshmavad WHERE (nomade=1) AND (idmade=" + Trim(Str(Form15.Adodc1.Recordset.Fields!idmavad)) + ")", db1
              Form7.Adodc1.Recordset.Fields!masrafteydoremeghdar = rs(0).Fields!naghlazgeranolmeghdar
'              MsgBox rs(0).Fields!naghlazgeranolmeghdar
            rs(0).Close
          Else
            Form7.Adodc1.Recordset.Fields!masrafteydoremeghdar = 0
          End If
        rs(1).Close
        
        Form4.Adodc1.Recordset.Find "idmavad=" + Trim(Str(Form15.Adodc1.Recordset.Fields!idmavad)), , adSearchForward, 1
        Form7.Adodc1.Recordset.Fields!kharidteydoremablagh = Form4.Adodc1.Recordset.Fields!gheymattamam
        Form7.Adodc1.Recordset.Fields!mojodiamademasrafmeghdar = Round(Val(Form7.Adodc1.Recordset.Fields!kharidteydoremeghdar) + Val(Form7.Adodc1.Recordset.Fields!moneyonedoremeghdar))
        Form7.Adodc1.Recordset.Fields!mojodiamademasrafmablagh = Round(Val(Form7.Adodc1.Recordset.Fields!kharidteydoremablagh) + Val(Form7.Adodc1.Recordset.Fields!moneyonedoremablagh))
      
        If Form7.Adodc1.Recordset.Fields!mojodiamademasrafmeghdar = 0 Then
          Form7.Adodc1.Recordset.Fields!mojodiamademasraffi = 0
        Else
          Form7.Adodc1.Recordset.Fields!mojodiamademasraffi = Round(Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmeghdar))
        End If
    
        If Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) <> 0 Then
          Form7.Adodc1.Recordset.Fields!masrafteydoremablagh = Round(Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Round(Val(Form7.Adodc1.Recordset.Fields!masrafteydoremeghdar)))
          Form7.Adodc1.Recordset.Fields!foroshteydoremablagh = Round(Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Val(Form7.Adodc1.Recordset.Fields!foroshteydoremeghdar))
  
          Form7.Adodc1.Recordset.Fields!zayeatmeghdar = Round(Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) - (Val(Form7.Adodc1.Recordset.Fields!masrafteydoremeghdar) + Val(Form7.Adodc1.Recordset.Fields!foroshteydoremeghdar) + Val(Form7.Adodc1.Recordset.Fields!mojodipayandoremeghdar)))
          Form7.Adodc1.Recordset.Fields!zayeatmablagh = Round(Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Val(Form7.Adodc1.Recordset.Fields!zayeatmeghdar))

          Form7.Adodc1.Recordset.Fields!mojodipayandoremablagh = Round(Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmablagh) / Val(Form7.Adodc1.Recordset.Fields!mojodiamademasrafmeghdar) * Val(Form7.Adodc1.Recordset.Fields!mojodipayandoremeghdar))
        Else
          Form7.Adodc1.Recordset.Fields!masrafteydoremablagh = 0
          Form7.Adodc1.Recordset.Fields!foroshteydoremablagh = 0
          Form7.Adodc1.Recordset.Fields!zayeatmeghdar = 0
          Form7.Adodc1.Recordset.Fields!zayeatmablagh = 0
          Form7.Adodc1.Recordset.Fields!mojodipayandoremablagh = 0
        End If
        Form7.Adodc1.Recordset.Update
      End If
    rs(4).Close
  End If
  Form15.Adodc1.Recordset.MoveNext
Loop Until Form15.Adodc1.Recordset.EOF = True

ProgressBar1.Value = 5
db1.Close
Command1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
Me.Hide
End Sub
