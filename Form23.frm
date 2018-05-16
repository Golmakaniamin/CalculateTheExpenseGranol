VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form23 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ò‰ —· ê—œ‘ „”"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12495
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form23.frx":0000
   LinkTopic       =   "Form23"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   12495
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "ÃœÊ· œÊ"
      TabPicture(0)   =   "Form23.frx":2CFA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DataGrid2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ÃœÊ· Ìò"
      TabPicture(1)   =   "Form23.frx":2D16
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Form23.frx":2D32
         Height          =   7935
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   13996
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   29
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
         ColumnCount     =   28
         BeginProperty Column00 
            DataField       =   "rad"
            Caption         =   "—œÌ›"
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
            DataField       =   "name"
            Caption         =   "‰«„ „—Õ·Â"
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
         BeginProperty Column02 
            DataField       =   "sanaveye1"
            Caption         =   "À«‰ÊÌÂ („ﬁœ«—)"
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
         BeginProperty Column03 
            DataField       =   "sanaveye2"
            Caption         =   "À«‰ÊÌÂ (—Ì«·)"
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
         BeginProperty Column04 
            DataField       =   "nahaee1"
            Caption         =   "‰Â«ÌÌ („ﬁœ«—)"
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
         BeginProperty Column05 
            DataField       =   "nahaee2"
            Caption         =   "‰Â«ÌÌ (—Ì«·)"
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
         BeginProperty Column06 
            DataField       =   "Koreh1"
            Caption         =   "òÊ—Â („ﬁœ«—)"
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
         BeginProperty Column07 
            DataField       =   "Koreh2"
            Caption         =   "òÊ—Â (—Ì«·)"
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
         BeginProperty Column08 
            DataField       =   "Taab1"
            Caption         =   " «» („ﬁœ«—)"
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
         BeginProperty Column09 
            DataField       =   "Taab2"
            Caption         =   " «» (—Ì«·)"
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
         BeginProperty Column10 
            DataField       =   "Sterander1_61"
            Caption         =   "«” —‰œ— 6+1 („ﬁœ«—)"
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
         BeginProperty Column11 
            DataField       =   "Sterander1_62"
            Caption         =   "«” —‰œ— 6+1 (—Ì«·)"
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
         BeginProperty Column12 
            DataField       =   "Sterander1_361"
            Caption         =   "«” —‰œ— 36+1 (—Ì«·)"
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
         BeginProperty Column13 
            DataField       =   "Sterander1_362"
            Caption         =   "«” —‰œ— 36+1 („ﬁœ«—)"
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
         BeginProperty Column14 
            DataField       =   "Sterander1_41"
            Caption         =   "«” —‰œ— 4+1 („ﬁœ«—)"
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
         BeginProperty Column15 
            DataField       =   "Sterander1_42"
            Caption         =   "«” —‰œ— 4+1 (—Ì«·)"
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
         BeginProperty Column16 
            DataField       =   "DramToester1"
            Caption         =   "œ—«„  ÊÌ” — („ﬁœ«—)"
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
         BeginProperty Column17 
            DataField       =   "DramToester2"
            Caption         =   "œ—«„  ÊÌ” — (—Ì«·)"
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
         BeginProperty Column18 
            DataField       =   "Mokhaberat1"
            Caption         =   "„Œ«»—«  („ﬁœ«—)"
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
         BeginProperty Column19 
            DataField       =   "Mokhaberat2"
            Caption         =   "„Œ«»—«  (—Ì«·)"
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
         BeginProperty Column20 
            DataField       =   "Exteroder1"
            Caption         =   "«ò” —Êœ— („ﬁœ«—)"
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
         BeginProperty Column21 
            DataField       =   "Exteroder2"
            Caption         =   "«ò” —Êœ— (—Ì«·)"
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
         BeginProperty Column22 
            DataField       =   "Bastebandi1"
            Caption         =   "»” Â »‰œÌ („ﬁœ«—)"
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
         BeginProperty Column23 
            DataField       =   "Bastebandi2"
            Caption         =   "»” Â »‰œÌ (—Ì«·)"
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
         BeginProperty Column24 
            DataField       =   "AnbarMahsol1"
            Caption         =   "«‰»«— „Õ’Ê· („ﬁœ«—)"
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
         BeginProperty Column25 
            DataField       =   "AnbarMahsol2"
            Caption         =   "«‰»«— „Õ’Ê· (—Ì«·)"
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
         BeginProperty Column26 
            DataField       =   "sum1"
            Caption         =   "Ã„⁄ („ﬁœ«—)"
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
         BeginProperty Column27 
            DataField       =   "sum2"
            Caption         =   "Ã„⁄ (—Ì«·)"
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
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
            EndProperty
            BeginProperty Column13 
            EndProperty
            BeginProperty Column14 
            EndProperty
            BeginProperty Column15 
            EndProperty
            BeginProperty Column16 
            EndProperty
            BeginProperty Column17 
            EndProperty
            BeginProperty Column18 
            EndProperty
            BeginProperty Column19 
            EndProperty
            BeginProperty Column20 
            EndProperty
            BeginProperty Column21 
            EndProperty
            BeginProperty Column22 
            EndProperty
            BeginProperty Column23 
            EndProperty
            BeginProperty Column24 
            EndProperty
            BeginProperty Column25 
            EndProperty
            BeginProperty Column26 
            EndProperty
            BeginProperty Column27 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form23.frx":2D47
         Height          =   7935
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   13996
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   29
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
            DataField       =   "rad"
            Caption         =   "—œÌ›"
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
            DataField       =   "name"
            Caption         =   "‰«„ „—Õ·Â"
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
         BeginProperty Column02 
            DataField       =   "mojodiavalmeghdar"
            Caption         =   "„ÊÃÊœÌ «Ê· œÊ—Â"
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
         BeginProperty Column03 
            DataField       =   "varedeteydoremeghdar"
            Caption         =   "Ê«—œÂ ÿÌ œÊ—Â"
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
         BeginProperty Column04 
            DataField       =   "enteghalbade"
            Caption         =   "«‰ ﬁ«· »Â Ê«Õœ »⁄œ"
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
         BeginProperty Column05 
            DataField       =   "zayeat"
            Caption         =   "÷«Ì⁄«   Ê·Ìœ"
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
         BeginProperty Column06 
            DataField       =   "mojodienddore"
            Caption         =   "„ÊÃÊœÌ Å«Ì«‰ œÊ—Â"
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "kontrolgardeshmes"
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
Attribute VB_Name = "Form23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Adodc1.Recordset.MoveFirst
Do
  Select Case Adodc1.Recordset.Fields!rad
    Case 1
      Form9.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = Form9.Adodc1.Recordset.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = Form9.Adodc1.Recordset.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = Form9.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = Form9.Adodc1.Recordset.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      Form9.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc1.Recordset.Fields!zayeat = Form9.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Update
      Call KontrolGardeshmes.amin1(Adodc1.Recordset.Fields!rad)

    Case 2
      Form10.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = Form10.Adodc1.Recordset.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = Form10.Adodc1.Recordset.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = Form10.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = Form10.Adodc1.Recordset.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      Form10.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc1.Recordset.Fields!zayeat = Form10.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Update
      Call KontrolGardeshmes.amin2(Adodc1.Recordset.Fields!rad)
      
    Case 3
      Form11.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = Form11.Adodc1.Recordset.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = Form11.Adodc1.Recordset.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = Form11.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = Form11.Adodc1.Recordset.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      Form11.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc1.Recordset.Fields!zayeat = Form11.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Update
      Call KontrolGardeshmes.amin3(Adodc1.Recordset.Fields!rad)

    Case 4
      Form13.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = Form13.Adodc1.Recordset.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = Form13.Adodc1.Recordset.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = Form13.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = Form13.Adodc1.Recordset.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      Form13.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc1.Recordset.Fields!zayeat = Form13.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Update
      Call KontrolGardeshmes.amin4(Adodc1.Recordset.Fields!rad)

    Case 5
      Form1.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = Form1.Adodc1.Recordset.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = Form1.Adodc1.Recordset.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = Form1.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = Form1.Adodc1.Recordset.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      Form1.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc1.Recordset.Fields!zayeat = Form1.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Update
      Call KontrolGardeshmes.amin5(Adodc1.Recordset.Fields!rad)

    Case 6
      Form14.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = Form14.Adodc1.Recordset.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = Form14.Adodc1.Recordset.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = Form14.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = Form14.Adodc1.Recordset.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      Form14.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc1.Recordset.Fields!zayeat = Form14.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Update
      Call KontrolGardeshmes.amin6(Adodc1.Recordset.Fields!rad)

    Case 7
      Form16.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = Form16.Adodc1.Recordset.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = Form16.Adodc1.Recordset.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = Form16.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = Form16.Adodc1.Recordset.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      Form16.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc1.Recordset.Fields!zayeat = Form16.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Update
      Call KontrolGardeshmes.amin7(Adodc1.Recordset.Fields!rad)
      
    Case 8
      Form17.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = Form17.Adodc1.Recordset.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = Form17.Adodc1.Recordset.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = Form17.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = Form17.Adodc1.Recordset.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      Form17.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc1.Recordset.Fields!zayeat = Form17.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Update
      Call KontrolGardeshmes.amin8(Adodc1.Recordset.Fields!rad)

    Case 9
      Form18.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = Form18.Adodc1.Recordset.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = Form18.Adodc1.Recordset.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = Form18.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = Form18.Adodc1.Recordset.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      Form18.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc1.Recordset.Fields!zayeat = Form18.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Update
      Call KontrolGardeshmes.amin9(Adodc1.Recordset.Fields!rad)
      
    Case 10
      Form19.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = Form19.Adodc1.Recordset.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = Form19.Adodc1.Recordset.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = Form19.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = Form19.Adodc1.Recordset.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      Form19.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc1.Recordset.Fields!zayeat = Form19.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Update
      Call KontrolGardeshmes.amin10(Adodc1.Recordset.Fields!rad)
      
    Case 11
      Form20.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = Form20.Adodc1.Recordset.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = Form20.Adodc1.Recordset.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = Form20.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = Form20.Adodc1.Recordset.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      Form20.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc1.Recordset.Fields!zayeat = Form20.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Update
      Call KontrolGardeshmes.amin11(Adodc1.Recordset.Fields!rad)
      
    Case 12
      Form21.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = Form21.Adodc1.Recordset.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = Form21.Adodc1.Recordset.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = Form21.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = Form21.Adodc1.Recordset.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      Form21.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc1.Recordset.Fields!zayeat = Form21.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Update
      Call KontrolGardeshmes.amin12(Adodc1.Recordset.Fields!rad)

    Case 13
      Form22.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = Form22.Adodc1.Recordset.Fields!mojodiavalmeghdar
      Adodc1.Recordset.Fields!varedeteydoremeghdar = Form22.Adodc1.Recordset.Fields!tolidteydoremeghdar
      Adodc1.Recordset.Fields!enteghalbade = Form22.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Fields!mojodienddore = Form22.Adodc1.Recordset.Fields!mojodiendmeghdar
      Adodc1.Recordset.Update
      Form22.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc1.Recordset.Fields!zayeat = Form22.Adodc1.Recordset.Fields!naghlbebadmeghdar
      Adodc1.Recordset.Update
      Call KontrolGardeshmes.amin13(Adodc1.Recordset.Fields!rad)

  End Select
  Adodc1.Recordset.MoveNext
Loop Until Adodc1.Recordset.EOF = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
Me.Hide
End Sub
