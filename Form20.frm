VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form20 
   Caption         =   "«ò” —Êœ—"
   ClientHeight    =   10815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12975
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form20.frx":0000
   LinkTopic       =   "Form20"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10815
   ScaleWidth      =   12975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "»Â —Ê“ —”«‰Ì"
      Height          =   465
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "store4"
      DataSource      =   "Adodc2"
      Height          =   495
      Index           =   3
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   3
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
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   2
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
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "store2"
      DataSource      =   "Adodc2"
      Height          =   495
      Index           =   1
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   11400
      Top             =   2160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      Caption         =   "Adodc3"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   9615
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   16960
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "ÃœÊ· „ﬁœ«—Ì —Ì«·Ì"
      TabPicture(0)   =   "Form20.frx":2CFA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DataGrid2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ÃœÊ· ﬁÌ„   „«„ ‘œÂ Ê«Õœ"
      TabPicture(1)   =   "Form20.frx":2D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "ÃœÊ· „ﬁœ«—Ì"
      TabPicture(2)   =   "Form20.frx":2D32
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "DataGrid3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "Form20.frx":2D4E
         Height          =   9015
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   15901
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   23
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
         ColumnCount     =   9
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
            Caption         =   "‰«„ „Õ’Ê·"
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
            DataField       =   "kodemahsol"
            Caption         =   "òœ „Õ’Ê·"
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
            DataField       =   "gothr"
            Caption         =   "ﬁÿ—"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1065
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "size"
            Caption         =   "”«Ì“"
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
            DataField       =   "mojodiavalmeghdar"
            Caption         =   "„ÊÃÊœÌ «Ê· œÊ—Â"
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
         BeginProperty Column06 
            DataField       =   "tolidteydoremeghdar"
            Caption         =   " Ê·Ìœ ÿÌ œÊ—Â"
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
         BeginProperty Column07 
            DataField       =   "naghlbebadmeghdar"
            Caption         =   "‰ﬁ· »Â Ê«Õœ »⁄œ"
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
         BeginProperty Column08 
            DataField       =   "mojodiendmeghdar"
            Caption         =   "„ÊÃÊœÌ «‰ Â«Ì œÊ—Â"
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
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Alignment       =   3
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form20.frx":2D63
         Height          =   9015
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   15901
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
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "rad"
            Caption         =   "—œÌ›"
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
            DataField       =   "name"
            Caption         =   "‰«„ „Õ’Ê·"
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
            DataField       =   "kodemahsol"
            Caption         =   "òœ „Õ’Ê·"
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
            DataField       =   "gothr"
            Caption         =   "ﬁÿ—"
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
            DataField       =   "size"
            Caption         =   "”«Ì“"
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
            DataField       =   "standard8"
            Caption         =   "«” «‰œ«—œ  Ê·Ìœ œ— 8 ”«⁄ "
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
            DataField       =   "mezantolidmostaghim"
            Caption         =   "„Ì“«‰  Ê·Ìœ „” ﬁÌ„"
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
         BeginProperty Column07 
            DataField       =   "zaribtahsimdarsaat"
            Caption         =   "÷—Ì»  Â”Ì„"
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
         BeginProperty Column08 
            DataField       =   "mavadaval"
            Caption         =   "„Ê«œ «Ê·ÌÂ"
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
         BeginProperty Column09 
            DataField       =   "zaribdastmozd"
            Caption         =   "÷—Ì» œ” „“œ"
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
         BeginProperty Column10 
            DataField       =   "dastmozd"
            Caption         =   "œ” „“œ"
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
         BeginProperty Column11 
            DataField       =   "zaribsarbar"
            Caption         =   "÷—Ì» ”—»«—"
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
         BeginProperty Column12 
            DataField       =   "sarbar"
            Caption         =   "”—»«—"
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
         BeginProperty Column13 
            DataField       =   "estelak"
            Caption         =   "«” Â·«ò"
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
         BeginProperty Column14 
            DataField       =   "gheymattamam"
            Caption         =   "ﬁÌ„   „«„ ‘œÂ"
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
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
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
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Form20.frx":2D78
         Height          =   9015
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   15901
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
         ColumnCount     =   16
         BeginProperty Column00 
            DataField       =   "rad"
            Caption         =   "—œÌ›"
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
            DataField       =   "name"
            Caption         =   "‰«„ „Õ’Ê·"
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
            DataField       =   "kodemahsol"
            Caption         =   "òœ „Õ’Ê·"
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
            DataField       =   "gothr"
            Caption         =   "ﬁÿ—"
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
            DataField       =   "size"
            Caption         =   "”«Ì“"
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
            DataField       =   "mojodiavalmeghdar"
            Caption         =   "„ÊÃÊœÌ «Ê· œÊ—Â „ﬁœ«—"
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
            DataField       =   "mojodiavalmemoney"
            Caption         =   "„ÊÃÊœÌ «Ê· œÊ—Â „»·€"
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
         BeginProperty Column07 
            DataField       =   "tolidteydoremeghdar"
            Caption         =   " Ê·Ìœ ÿÌ œÊ—Â „ﬁœ«—"
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
         BeginProperty Column08 
            DataField       =   "tolidteydoremoney"
            Caption         =   " Ê·Ìœ ÿÌ œÊ—Â „»·€"
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
         BeginProperty Column09 
            DataField       =   "masrafmeghdar"
            Caption         =   "¬„«œÂ „’—› „ﬁœ«—"
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
         BeginProperty Column10 
            DataField       =   "masrafmoney"
            Caption         =   "¬„«œÂ „’—› „»·€"
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
         BeginProperty Column11 
            DataField       =   "fey"
            Caption         =   "›Ì"
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
         BeginProperty Column12 
            DataField       =   "naghlbebadmeghdar"
            Caption         =   "‰ﬁ· »Â Ê«Õœ »⁄œ „ﬁœ«—"
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
         BeginProperty Column13 
            DataField       =   "naghlbebadmoney"
            Caption         =   "‰ﬁ· »Â Ê«Õœ »⁄œ „»·€"
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
         BeginProperty Column14 
            DataField       =   "mojodiendmeghdar"
            Caption         =   "„ÊÃÊœÌ ¬Œ— œÊ—Â „ﬁœ«—"
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
         BeginProperty Column15 
            DataField       =   "mojodiendmoney"
            Caption         =   "„ÊÃÊœÌ ¬Œ— œÊ—Â „»·€"
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
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
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
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   10920
      Top             =   1440
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10920
      Top             =   1200
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
      RecordSource    =   "Taab"
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
      Height          =   375
      Left            =   120
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      Caption         =   "Adodc3"
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
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«” Â·«ò"
      Height          =   495
      Index           =   2
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "”—»«—"
      Height          =   495
      Index           =   1
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "»Â«Ì „”"
      Height          =   495
      Index           =   0
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "»Â«Ì œ” „“œ"
      Height          =   495
      Index           =   4
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Ã„⁄ :"
      Height          =   495
      Index           =   3
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmp0(20) As String, introwforall As Integer

Private Sub DataGrid_sum_help()
'On Error Resume Next
  For q = 0 To 20
    tmp0(q) = 0
  Next q
  Adodc3.Refresh
  
  Adodc3.Recordset.MoveFirst
  Do
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!standard8 <> "") Then tmp0(0) = Val(tmp0(0)) + Val(Adodc3.Recordset.Fields!standard8)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!mezantolidmostaghim <> "") Then tmp0(1) = Val(tmp0(1)) + Val(Adodc3.Recordset.Fields!mezantolidmostaghim)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!zaribtahsimdarsaat <> "") Then tmp0(2) = Val(tmp0(2)) + Val(Adodc3.Recordset.Fields!zaribtahsimdarsaat)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!mavadaval <> "") Then tmp0(3) = Val(tmp0(3)) + Val(Adodc3.Recordset.Fields!mavadaval)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!zaribdastmozd <> "") Then tmp0(4) = Val(tmp0(4)) + Val(Adodc3.Recordset.Fields!zaribdastmozd)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!dastmozd <> "") Then tmp0(5) = Val(tmp0(5)) + Val(Adodc3.Recordset.Fields!dastmozd)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!zaribsarbar <> "") Then tmp0(6) = Val(tmp0(6)) + Val(Adodc3.Recordset.Fields!zaribsarbar)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!sarbar <> "") Then tmp0(7) = Val(tmp0(7)) + Val(Adodc3.Recordset.Fields!sarbar)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!estelak <> "") Then tmp0(8) = Val(tmp0(8)) + Val(Adodc3.Recordset.Fields!estelak)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!gheymattamam <> "") Then tmp0(9) = Val(tmp0(9)) + Val(Adodc3.Recordset.Fields!gheymattamam)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!mojodiavalmeghdar <> "") Then tmp0(10) = Val(tmp0(10)) + Val(Adodc3.Recordset.Fields!mojodiavalmeghdar)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!mojodiavalmemoney <> "") Then tmp0(11) = Val(tmp0(11)) + Val(Adodc3.Recordset.Fields!mojodiavalmemoney)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!tolidteydoremeghdar <> "") Then tmp0(12) = Val(tmp0(12)) + Val(Adodc3.Recordset.Fields!tolidteydoremeghdar)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!tolidteydoremoney <> "") Then tmp0(13) = Val(tmp0(13)) + Val(Adodc3.Recordset.Fields!tolidteydoremoney)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!naghlbebadmoney <> "") Then tmp0(14) = Val(tmp0(14)) + Val(Adodc3.Recordset.Fields!naghlbebadmoney)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!naghlbebadmeghdar <> "") Then tmp0(15) = Val(tmp0(15)) + Val(Adodc3.Recordset.Fields!naghlbebadmeghdar)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!mojodiendmeghdar <> "") Then tmp0(16) = Val(tmp0(16)) + Val(Adodc3.Recordset.Fields!mojodiendmeghdar)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!mojodiendmoney <> "") Then tmp0(17) = Val(tmp0(17)) + Val(Adodc3.Recordset.Fields!mojodiendmoney)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!masrafmeghdar <> "") Then tmp0(18) = Val(tmp0(18)) + Val(Adodc3.Recordset.Fields!masrafmeghdar)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!masrafmoney <> "") Then tmp0(19) = Val(tmp0(19)) + Val(Adodc3.Recordset.Fields!masrafmoney)
    If (Adodc3.Recordset.Fields!rad <> "99997") And (Adodc3.Recordset.Fields!rad <> "99999") And (Adodc3.Recordset.Fields!fey <> "") Then tmp0(20) = Val(tmp0(20)) + Val(Adodc3.Recordset.Fields!fey)
    Adodc3.Recordset.MoveNext
  Loop Until Adodc3.Recordset.EOF = True
  
  
'  Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
'  tmp0(0) = tmp0(0) - Val(Adodc1.Recordset.Fields!standard8)
'  tmp0(1) = tmp0(1) - Val(Adodc1.Recordset.Fields!mezantolidmostaghim)
'  tmp0(2) = tmp0(2) - Val(Adodc1.Recordset.Fields!zaribtahsimdarsaat)
'  tmp0(3) = tmp0(3) - Val(Adodc1.Recordset.Fields!mavadaval)
'  tmp0(4) = tmp0(4) - Val(Adodc1.Recordset.Fields!zaribdastmozd)
'  tmp0(5) = tmp0(5) - Val(Adodc1.Recordset.Fields!dastmozd)
'  tmp0(6) = tmp0(6) - Val(Adodc1.Recordset.Fields!zaribsarbar)
'  tmp0(7) = tmp0(7) - Val(Adodc1.Recordset.Fields!sarbar)
'  tmp0(8) = tmp0(8) - Val(Adodc1.Recordset.Fields!estelak)
'  tmp0(9) = tmp0(9) - Val(Adodc1.Recordset.Fields!gheymattamam)
'  tmp0(10) = tmp0(10) - Val(Adodc1.Recordset.Fields!mojodiavalmeghdar)
'  tmp0(11) = tmp0(11) - Val(Adodc1.Recordset.Fields!mojodiavalmemoney)
'  tmp0(12) = tmp0(12) - Val(Adodc1.Recordset.Fields!tolidteydoremeghdar)
'  tmp0(13) = tmp0(13) - Val(Adodc1.Recordset.Fields!tolidteydoremoney)
'  tmp0(14) = tmp0(14) - Val(Adodc1.Recordset.Fields!naghlbebadmoney)
'  tmp0(15) = tmp0(15) - Val(Adodc1.Recordset.Fields!naghlbebadmeghdar)
'  tmp0(16) = tmp0(16) - Val(Adodc1.Recordset.Fields!mojodiendmeghdar)
'  tmp0(17) = tmp0(17) - Val(Adodc1.Recordset.Fields!mojodiendmoney)
  
  Adodc3.Refresh
  Adodc3.Recordset.Find "rad=99999", , adSearchForward, 1
  Adodc3.Recordset.Fields!standard8 = tmp0(0)
  Adodc3.Recordset.Fields!mezantolidmostaghim = tmp0(1)
  Adodc3.Recordset.Fields!zaribtahsimdarsaat = tmp0(2)
  Adodc3.Recordset.Fields!mavadaval = tmp0(3)
  Adodc3.Recordset.Fields!zaribdastmozd = tmp0(4)
  Adodc3.Recordset.Fields!dastmozd = tmp0(5)
  Adodc3.Recordset.Fields!zaribsarbar = tmp0(6)
  Adodc3.Recordset.Fields!sarbar = tmp0(7)
  Adodc3.Recordset.Fields!estelak = tmp0(8)
  Adodc3.Recordset.Fields!gheymattamam = tmp0(9)
  Adodc3.Recordset.Fields!mojodiavalmeghdar = tmp0(10)
  Adodc3.Recordset.Fields!mojodiavalmemoney = tmp0(11)
  Adodc3.Recordset.Fields!tolidteydoremeghdar = tmp0(12)
  Adodc3.Recordset.Fields!tolidteydoremoney = tmp0(13)
  Adodc3.Recordset.Fields!naghlbebadmoney = tmp0(14)
  Adodc3.Recordset.Fields!naghlbebadmeghdar = tmp0(15)
  Adodc3.Recordset.Fields!mojodiendmeghdar = tmp0(16)
  Adodc3.Recordset.Fields!mojodiendmoney = tmp0(17)
  Adodc3.Recordset.Fields!masrafmeghdar = tmp0(18)
  Adodc3.Recordset.Fields!masrafmoney = tmp0(19)
  Adodc3.Recordset.Fields!fey = tmp0(20)
  Adodc3.Recordset.Update
  Adodc3.Refresh
  
  Adodc1.Refresh
  Adodc1.Recordset.Find "rad=" + Trim(Str(introwforall)), , adSearchForward, 1
End Sub

Private Sub DataGrid1_AfterColUpdate_help()
If (Val(Adodc1.Recordset.Fields!rad) <> 99999) And (Val(Adodc1.Recordset.Fields!rad) <> 99997) Then
  '÷—Ì»  Â”Ì„

  q = Adodc1.Recordset.Fields!standard8
  If IsNull(q) = True Then q = 0
  
  w = Adodc1.Recordset.Fields!mezantolidmostaghim
  If IsNull(w) = True Then w = 0
  
  r = (q * w) / 8
  
  Adodc1.Recordset.Fields!zaribtahsimdarsaat = Round(r)
  
  '„Ê«œ «Ê·ÌÂ
  
  Adodc3.Refresh
  Adodc3.Recordset.Find "rad=99999", , adSearchForward, 1
  If Adodc3.Recordset.Fields!zaribtahsimdarsaat <> 0 Then
    q = Adodc3.Recordset.Fields!zaribtahsimdarsaat
    If q = "" Then q = 0
    If IsNull(q) = True Then q = 0
    
    w = Adodc1.Recordset.Fields!zaribtahsimdarsaat
    If w = "" Then w = 0
    If IsNull(w) = True Then w = 0
    
    r = Val(Text1(0).Text) / q * w
    If r = "" Then r = 0
    If IsNull(r) = True Then r = 0
    
    w1 = Adodc1.Recordset.Fields!zaribdastmozd
    
    If w1 = "" Then w1 = 0
    If IsNull(w1) = True Then w1 = 0
    r1 = Val(w1) * (Val(Text1(1).Text) / q * w)
    
    w1 = Adodc1.Recordset.Fields!zaribsarbar
    If w1 = "" Then w1 = 0
    If IsNull(w1) = True Then w1 = 0
    r2 = Val(w1) * (Val(Text1(2).Text) / q * w)
    r3 = Val(w1) * (Val(Text1(3).Text) / q * w)
  Else
    r = 0
    r1 = 0
    r2 = 0
    r3 = 0
  End If
  
'  r = Round(r)
'  r1 = Round(r1)
'  r2 = Round(r2)
'  r3 = Round(r3)
  
  Adodc1.Recordset.Fields!mavadaval = Round(r)
  Adodc1.Recordset.Fields!dastmozd = Round(r1)
  Adodc1.Recordset.Fields!sarbar = Round(r2)
  Adodc1.Recordset.Fields!estelak = Round(r3)
  Adodc1.Recordset.Fields!gheymattamam = Round(Val(r) + Val(r1) + Val(r2) + Val(r3))
  Adodc1.Recordset.Update
  Call DataGrid_sum_help
End If
End Sub

Private Sub DataGrid2_AfterColUpdate_help()

If (Val(Adodc1.Recordset.Fields!rad) <> 99999) And (Val(Adodc1.Recordset.Fields!rad) <> 99998) And (Val(Adodc1.Recordset.Fields!rad) <> 99997) Then
  q = Adodc1.Recordset.Fields!mojodiavalmeghdar
  w = Adodc1.Recordset.Fields!mojodiavalmemoney
  e = Adodc1.Recordset.Fields!tolidteydoremeghdar
  Adodc1.Recordset.Fields!tolidteydoremoney = Round(Adodc1.Recordset.Fields!gheymattamam)
  r = Adodc1.Recordset.Fields!tolidteydoremoney
  
  Adodc1.Recordset.Fields!masrafmeghdar = Val(q) + Val(e)
  Adodc1.Recordset.Fields!masrafmoney = Val(w) + Val(r)
  
  If Val(Adodc1.Recordset.Fields!masrafmeghdar) <> 0 Then
    Adodc1.Recordset.Fields!fey = Round(Val(Adodc1.Recordset.Fields!masrafmoney) / Val(Adodc1.Recordset.Fields!masrafmeghdar))
  Else
    Adodc1.Recordset.Fields!fey = 0
  End If
  Adodc1.Recordset.Fields!naghlbebadmoney = Val(Adodc1.Recordset.Fields!naghlbebadmeghdar) * Val(Adodc1.Recordset.Fields!fey)
  Adodc1.Recordset.Fields!mojodiendmoney = Val(Adodc1.Recordset.Fields!mojodiendmeghdar) * Val(Adodc1.Recordset.Fields!fey)
  Adodc1.Recordset.Update
  Call DataGrid_sum_help
End If
End Sub

Private Sub DataGrid3_AfterColUpdate_help()

If (Val(Adodc1.Recordset.Fields!rad) <> 99999) And (Val(Adodc1.Recordset.Fields!rad) <> 99998) And (Val(Adodc1.Recordset.Fields!rad) <> 99997) Then
  
  DataGrid3.Col = 5
  q = DataGrid3.Text
  
  
  DataGrid3.Col = 6
  w = DataGrid3.Text

  DataGrid1.Col = 6
  DataGrid1.Text = DataGrid3.Text

  DataGrid3.Col = 8
  e = DataGrid3.Text

  DataGrid3.Col = 7
  DataGrid3.Text = (Val(q) + Val(w)) - Val(e)
  DataGrid3.Refresh

  Call DataGrid1_AfterColUpdate_help
  Call DataGrid_sum_help
End If
End Sub

Private Sub Command1_Click()
q = 0
w = 0
Form9.Adodc1.Recordset.MoveFirst
Do
  If Form9.Adodc1.Recordset.Fields!Name = "«ò” —Êœ—" Then
    q = Val(q) + Val(Form9.Adodc1.Recordset.Fields!naghlbebadmeghdar)
    w = Val(w) + Val(Form9.Adodc1.Recordset.Fields!naghlbebadmoney)
  End If
  Form9.Adodc1.Recordset.MoveNext
Loop Until Form9.Adodc1.Recordset.EOF = True

Form10.Adodc1.Recordset.MoveFirst
Do
  If Form10.Adodc1.Recordset.Fields!Name = "«ò” —Êœ—" Then
    q = Val(q) + Val(Form10.Adodc1.Recordset.Fields!naghlbebadmeghdar)
    w = Val(w) + Val(Form10.Adodc1.Recordset.Fields!naghlbebadmoney)
  End If
  Form10.Adodc1.Recordset.MoveNext
Loop Until Form10.Adodc1.Recordset.EOF = True

Form11.Adodc1.Recordset.MoveFirst
Do
  If Form11.Adodc1.Recordset.Fields!Name = "«ò” —Êœ—" Then
    q = Val(q) + Val(Form11.Adodc1.Recordset.Fields!naghlbebadmeghdar)
    w = Val(w) + Val(Form11.Adodc1.Recordset.Fields!naghlbebadmoney)
  End If
  Form11.Adodc1.Recordset.MoveNext
Loop Until Form11.Adodc1.Recordset.EOF = True

Form13.Adodc1.Recordset.MoveFirst
Do
  If Form13.Adodc1.Recordset.Fields!Name = "«ò” —Êœ—" Then
    q = Val(q) + Val(Form13.Adodc1.Recordset.Fields!naghlbebadmeghdar)
    w = Val(w) + Val(Form13.Adodc1.Recordset.Fields!naghlbebadmoney)
  End If
  Form13.Adodc1.Recordset.MoveNext
Loop Until Form13.Adodc1.Recordset.EOF = True

Form15.Adodc3.CommandType = adCmdUnknown
Form15.Adodc3.RecordSource = "SELECT * FROM ozanmasir WHERE name= '«ò” —Êœ—' "
Form15.Adodc3.Refresh
If Form15.Adodc3.Recordset.RecordCount > 0 Then
  Form15.Adodc3.Recordset.MoveFirst
  Do
    Adodc3.ConnectionString = Form3.Text10.Text
    Adodc3.CommandType = adCmdUnknown
    Adodc3.RecordSource = "select * from Exteroder WHERE (idmahsol=" + Trim(Str(Form15.Adodc3.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Form15.Adodc3.Recordset.Fields!rad)) + ")"
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount = 0 Then
      Form15.Adodc1.CommandType = adCmdUnknown
      Form15.Adodc1.RecordSource = "SELECT * FROM ozanmain WHERE (idmahsol=" + Trim(Str(Form15.Adodc3.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Form15.Adodc3.Recordset.Fields!rad)) + ")"
      Form15.Adodc1.Refresh
      Adodc1.Refresh
      Adodc1.Recordset.AddNew
      Adodc1.Recordset.Fields!rad = Form15.Adodc1.Recordset.Fields!rad
      Adodc1.Recordset.Fields!idmahsol = Form15.Adodc1.Recordset.Fields!idmahsol
      Adodc1.Recordset.Fields!kodemahsol = Form15.Adodc1.Recordset.Fields!kodemahsol
      Adodc1.Recordset.Fields!gothr = Form15.Adodc1.Recordset.Fields!gothr
      Adodc1.Recordset.Fields!Size = Form15.Adodc1.Recordset.Fields!Size
      Form2.Adodc1.Recordset.Find "idmahsol=" + Trim(Str(Form15.Adodc1.Recordset.Fields!idmahsol))
      Adodc1.Recordset.Fields!Name = Form2.Adodc1.Recordset.Fields!mahsol
      Adodc1.Recordset.Fields!standard8 = 1
      Adodc1.Recordset.Fields!mezantolidmostaghim = 1
      Adodc1.Recordset.Fields!zaribtahsimdarsaat = 1
      Adodc1.Recordset.Fields!mavadaval = 1
      Adodc1.Recordset.Fields!zaribdastmozd = 1
      Adodc1.Recordset.Fields!dastmozd = 1
      Adodc1.Recordset.Fields!zaribsarbar = 1
      Adodc1.Recordset.Fields!sarbar = 1
      Adodc1.Recordset.Fields!estelak = 1
      Adodc1.Recordset.Fields!gheymattamam = 1
      Adodc1.Recordset.Fields!mojodiavalmeghdar = 1
      Adodc1.Recordset.Fields!mojodiavalmemoney = 1
      Adodc1.Recordset.Fields!tolidteydoremeghdar = 1
      Adodc1.Recordset.Fields!tolidteydoremoney = 1
      Adodc1.Recordset.Fields!naghlbebadmoney = 1
      Adodc1.Recordset.Fields!naghlbebadmeghdar = 1
      Adodc1.Recordset.Fields!mojodiendmeghdar = 1
      Adodc1.Recordset.Fields!mojodiendmoney = 1
      Adodc1.Recordset.Fields!masrafmeghdar = 1
      Adodc1.Recordset.Fields!masrafmoney = 1
      Adodc1.Recordset.Fields!fey = 1
      Adodc1.Recordset.Update
    Else
      Form15.Adodc1.CommandType = adCmdUnknown
      Form15.Adodc1.RecordSource = "SELECT * FROM ozanmain WHERE (idmahsol=" + Trim(Str(Form15.Adodc3.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Form15.Adodc3.Recordset.Fields!rad)) + ")"
      Form15.Adodc1.Refresh
      Adodc3.Recordset.Fields!kodemahsol = Form15.Adodc1.Recordset.Fields!kodemahsol
      Adodc3.Recordset.Fields!gothr = Form15.Adodc1.Recordset.Fields!gothr
      Adodc3.Recordset.Fields!Size = Form15.Adodc1.Recordset.Fields!Size
      Form2.Adodc1.Recordset.Find "idmahsol=" + Trim(Str(Form15.Adodc1.Recordset.Fields!idmahsol))
      Adodc3.Recordset.Fields!Name = Form2.Adodc1.Recordset.Fields!mahsol
      Adodc3.Recordset.Update
      Adodc3.Refresh
    End If
    
    Adodc4.CommandType = adCmdUnknown
    Adodc4.RecordSource = "SELECT * FROM ozanmasir WHERE (idmahsol=" + Trim(Str(Form15.Adodc3.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Form15.Adodc3.Recordset.Fields!rad)) + ") ORDER BY rad1 ASC"
    Adodc4.Refresh
    If Adodc4.Recordset.RecordCount > 0 Then
      Adodc4.Recordset.Find "name= '«ò” —Êœ—'"
      Adodc4.Recordset.MovePrevious
      If Adodc4.Recordset.BOF = False Then
        Select Case Adodc4.Recordset.Fields!Name
        
          Case " «»"
            Form1.Adodc3.RecordSource = "SELECT * From Taab WHERE (idmahsol=" + Trim(Str(Adodc4.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc4.Recordset.Fields!rad)) + ")"
            Form1.Adodc3.Refresh
            If Form1.Adodc3.Recordset.RecordCount > 0 Then
              q = Val(q) + Val(Form1.Adodc3.Recordset.Fields!naghlbebadmeghdar)
              w = Val(w) + Val(Form1.Adodc3.Recordset.Fields!naghlbebadmoney)
            End If
            
          Case "«” —‰œ— 6 +1"
            Form14.Adodc3.RecordSource = "SELECT * From Sterander1_6 WHERE (idmahsol=" + Trim(Str(Adodc4.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc4.Recordset.Fields!rad)) + ")"
            Form14.Adodc3.Refresh
            If Form14.Adodc3.Recordset.RecordCount > 0 Then
              q = Val(q) + Val(Form14.Adodc3.Recordset.Fields!naghlbebadmeghdar)
              w = Val(w) + Val(Form14.Adodc3.Recordset.Fields!naghlbebadmoney)
            End If
            
          Case "«” —‰œ— 36 + 1"
            Form16.Adodc3.RecordSource = "SELECT * From Sterander1_36 WHERE (idmahsol=" + Trim(Str(Adodc4.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc4.Recordset.Fields!rad)) + ")"
            Form16.Adodc3.Refresh
            If Form16.Adodc3.Recordset.RecordCount > 0 Then
              q = Val(q) + Val(Form16.Adodc3.Recordset.Fields!naghlbebadmeghdar)
              w = Val(w) + Val(Form16.Adodc3.Recordset.Fields!naghlbebadmoney)
            End If
            
          Case "«” —‰œ— 4 + 1"
            Form17.Adodc3.RecordSource = "SELECT * From Sterander1_4 WHERE (idmahsol=" + Trim(Str(Adodc4.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc4.Recordset.Fields!rad)) + ")"
            Form17.Adodc3.Refresh
            If Form17.Adodc3.Recordset.RecordCount > 0 Then
              q = Val(q) + Val(Form17.Adodc3.Recordset.Fields!naghlbebadmeghdar)
              w = Val(w) + Val(Form17.Adodc3.Recordset.Fields!naghlbebadmoney)
            End If
                        
          Case "œ—«„  ÊÌ” —"
            Form18.Adodc3.RecordSource = "SELECT * From DramToester WHERE (idmahsol=" + Trim(Str(Adodc4.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc4.Recordset.Fields!rad)) + ")"
            Form18.Adodc3.Refresh
            If Form18.Adodc3.Recordset.RecordCount > 0 Then
              q = Val(q) + Val(Form18.Adodc3.Recordset.Fields!naghlbebadmeghdar)
              w = Val(w) + Val(Form18.Adodc3.Recordset.Fields!naghlbebadmoney)
            End If
                        
          Case "„Œ«»—« Ì"
            Form19.Adodc3.RecordSource = "SELECT * From Mokhaberat WHERE (idmahsol=" + Trim(Str(Adodc4.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc4.Recordset.Fields!rad)) + ")"
            Form19.Adodc3.Refresh
            If Form19.Adodc3.Recordset.RecordCount > 0 Then
              q = Val(q) + Val(Form19.Adodc3.Recordset.Fields!naghlbebadmeghdar)
              w = Val(w) + Val(Form19.Adodc3.Recordset.Fields!naghlbebadmoney)
            End If
                        
          Case "»” Â »‰œÌ"
            Form21.Adodc3.RecordSource = "SELECT * From Bastebandi WHERE (idmahsol=" + Trim(Str(Adodc4.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc4.Recordset.Fields!rad)) + ")"
            Form21.Adodc3.Refresh
            If Form21.Adodc3.Recordset.RecordCount > 0 Then
              q = Val(q) + Val(Form21.Adodc3.Recordset.Fields!naghlbebadmeghdar)
              w = Val(w) + Val(Form21.Adodc3.Recordset.Fields!naghlbebadmoney)
            End If
                        
          Case "«‰»«— „Õ’Ê·"
            Form22.Adodc3.RecordSource = "SELECT * From AnbarMahsol WHERE (idmahsol=" + Trim(Str(Adodc4.Recordset.Fields!idmahsol)) + ") and (rad=" + Trim(Str(Adodc4.Recordset.Fields!rad)) + ")"
            Form22.Adodc3.Refresh
            If Form22.Adodc3.Recordset.RecordCount > 0 Then
              q = Val(q) + Val(Form22.Adodc3.Recordset.Fields!naghlbebadmeghdar)
              w = Val(w) + Val(Form22.Adodc3.Recordset.Fields!naghlbebadmoney)
            End If
                        
        End Select
      End If
    End If
    Form15.Adodc3.Recordset.MoveNext
  Loop Until Form15.Adodc3.Recordset.EOF = True
End If

Adodc3.ConnectionString = Form3.Text10.Text
Adodc3.CommandType = adCmdUnknown
Adodc3.RecordSource = "select * from Exteroder ORDER BY rad ASC"
Adodc3.Refresh

Adodc1.Recordset.Find "rad=99997", , adSearchForward, 1
Adodc1.Recordset.Fields!tolidteydoremeghdar = q
Adodc1.Recordset.Fields!tolidteydoremoney = w
Adodc1.Recordset.Update

Adodc1.Refresh
DataGrid3.Refresh
Adodc1.Refresh
DataGrid3.Refresh
Adodc2.Recordset.Fields!store1 = w
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
  introwforall = Adodc1.Recordset.Fields!rad
  Call DataGrid1_AfterColUpdate_help
End Sub

Private Sub DataGrid2_AfterColEdit(ByVal ColIndex As Integer)
  introwforall = Adodc1.Recordset.Fields!rad
  Call DataGrid2_AfterColUpdate_help
End Sub

Private Sub DataGrid3_AfterColEdit(ByVal ColIndex As Integer)
  introwforall = Adodc1.Recordset.Fields!rad
  Call DataGrid3_AfterColUpdate_help
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
Me.Hide
End Sub

Private Sub Text1_Change(Index As Integer)
Label1.Caption = Val(Text1(0).Text) + Val(Text1(1).Text) + Val(Text1(2).Text) + Val(Text1(3).Text)
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






