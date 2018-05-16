VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form24 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«” Â·«ﬂ"
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
   Icon            =   "Form24.frx":0000
   LinkTopic       =   "Form24"
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
      TabCaption(0)   =   "ÃœÊ· „Õ«”»Â «” Â·«ﬂ œ«—«∆ÌÂ«Ì À«»  "
      TabPicture(0)   =   "Form24.frx":2CFA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DataGrid2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "÷—«Ì» ›‰Ì  ”ÂÌ„ «” Â·«ﬂ"
      TabPicture(1)   =   "Form24.frx":2D16
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Form24.frx":2D32
         Height          =   7935
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   13996
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
         ColumnCount     =   21
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
            Caption         =   "„—«Õ·  Ê·Ìœ"
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
            DataField       =   "vasayelnaglmarkazi"
            Caption         =   "Ê”«Ì· ‰ﬁ·ÌÂ œ› — „—ﬂ“Ì"
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
            DataField       =   "skarghahi"
            Caption         =   "”«Œ „«‰ ﬂ«—ê«ÂÌ"
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
            DataField       =   "sedari"
            Caption         =   "”«Œ „«‰ «œ«—Ì"
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
            DataField       =   "snaghahban"
            Caption         =   "”«Œ „«‰ ‰êÂ»«‰Ì Ê ⁄„Ê„Ì"
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
            DataField       =   "mashinkarkhane"
            Caption         =   "„«‘Ì‰ ¬·«  ﬂ«—Œ«‰Â"
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
            DataField       =   "tashararat"
            Caption         =   " «”Ì”«  Õ—«— Ì"
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
            DataField       =   "tasab"
            Caption         =   " «”Ì”«  ¬»—”«‰Ì"
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
            DataField       =   "tascool"
            Caption         =   " «”Ì”«  Œ‰ﬂ ﬂ‰‰œÂ"
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
            DataField       =   "tasbargh"
            Caption         =   " «”Ì”«  »—ﬁ —”«‰Ì"
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
            DataField       =   "lavazemazmayeshgah"
            Caption         =   "·Ê«“„ ¬“„«Ì‘ê«ÂÌ"
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
            DataField       =   "asasedari"
            Caption         =   "«À«ÀÌÂ Ê ·Ê«“„ «œ«—Ì"
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
            DataField       =   "vasayelertebati"
            Caption         =   "Ê”«Ì· «— »«ÿÌ"
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
            DataField       =   "mashinsakhteman"
            Caption         =   "„«‘Ì‰ ¬·«  ”«Œ „«‰Ì"
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
            DataField       =   "tasmovaledhava"
            Caption         =   " «”Ì”«  „Ê·œ ÂÊ«"
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
            DataField       =   "vasayelvalefterak"
            Caption         =   "Ê”«Ìÿ ‰ﬁ·ÌÂ Ê ·Ì› —«ﬂ Â«Ì ﬂ«—Œ«‰Â"
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
            DataField       =   "makhazenDOP"
            Caption         =   "„Œ«“‰ Ê  «‰ﬂ Â«Ì DOP"
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
            DataField       =   "tasisgaz"
            Caption         =   " «”Ì”«  ê«“—”«‰Ì"
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
            DataField       =   "abzarkargah"
            Caption         =   "«»“«— ¬·«  ﬂ«—ê«ÂÌ Ê „ ›—ﬁÂ"
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
            DataField       =   "sumend"
            Caption         =   "Ã„⁄ ‰Â«ÌÌ"
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
               ColumnWidth     =   2160
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
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form24.frx":2D47
         Height          =   7935
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   13996
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
            Caption         =   "„—«Õ·  Ê·Ìœ"
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
            DataField       =   "zirbana"
            Caption         =   "“Ì—»‰«Ì"
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
            DataField       =   "abresani"
            Caption         =   "¬» —”«‰Ì"
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
            DataField       =   "cooler"
            Caption         =   "Œ‰ﬂ ﬂ‰‰œÂ"
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
            DataField       =   "naghleye"
            Caption         =   "Ê”«Ìÿ ‰ﬁ·ÌÂ"
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
            DataField       =   "tasisathararati"
            Caption         =   " «”Ì”«  Õ—«— Ì"
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
            DataField       =   "kilovat"
            Caption         =   "»—ﬁ —”«‰Ì"
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
            DataField       =   "asaskarkhane"
            Caption         =   "«À«À ﬂ«—Œ«‰Â"
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
      RecordSource    =   "Estehlak1"
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
      Left            =   1440
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
      RecordSource    =   "Estehlak2"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   2760
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   4080
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
End
Attribute VB_Name = "Form24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmp(20) As String
Dim tmp0(20) As String, introwforall As Integer, blnasd As Boolean, tmpasd As String

Private Sub DataGrid_sum_help()
  For q = 0 To 20
    tmp0(q) = 0
  Next q
  Adodc3.Refresh
  Adodc3.Recordset.MoveFirst
  Do
    If (Adodc3.Recordset.Fields!rad <> "999") And (Adodc3.Recordset.Fields!zirbana <> "") Then tmp0(0) = Val(tmp0(0)) + Val(Adodc3.Recordset.Fields!zirbana)
    If (Adodc3.Recordset.Fields!rad <> "999") And (Adodc3.Recordset.Fields!abresani <> "") Then tmp0(1) = Val(tmp0(1)) + Val(Adodc3.Recordset.Fields!abresani)
    If (Adodc3.Recordset.Fields!rad <> "999") And (Adodc3.Recordset.Fields!cooler <> "") Then tmp0(2) = Val(tmp0(2)) + Val(Adodc3.Recordset.Fields!cooler)
    If (Adodc3.Recordset.Fields!rad <> "999") And (Adodc3.Recordset.Fields!naghleye <> "") Then tmp0(3) = Val(tmp0(3)) + Val(Adodc3.Recordset.Fields!naghleye)
    If (Adodc3.Recordset.Fields!rad <> "999") And (Adodc3.Recordset.Fields!tasisathararati <> "") Then tmp0(4) = Val(tmp0(4)) + Val(Adodc3.Recordset.Fields!tasisathararati)
    If (Adodc3.Recordset.Fields!rad <> "999") And (Adodc3.Recordset.Fields!kilovat <> "") Then tmp0(5) = Val(tmp0(5)) + Val(Adodc3.Recordset.Fields!kilovat)
    If (Adodc3.Recordset.Fields!rad <> "999") And (Adodc3.Recordset.Fields!asaskarkhane <> "") Then tmp0(6) = Val(tmp0(6)) + Val(Adodc3.Recordset.Fields!asaskarkhane)
    Adodc3.Recordset.MoveNext
  Loop Until Adodc3.Recordset.EOF = True
  For q = 1 To 300
  Next q
  Adodc3.Refresh
  Adodc3.Recordset.Find "rad=999", , adSearchForward, 1
  Adodc3.Recordset.Fields!zirbana = tmp0(0)
  Adodc3.Recordset.Fields!abresani = tmp0(1)
  Adodc3.Recordset.Fields!cooler = tmp0(2)
  Adodc3.Recordset.Fields!naghleye = tmp0(3)
  Adodc3.Recordset.Fields!tasisathararati = tmp0(4)
  Adodc3.Recordset.Fields!kilovat = tmp0(5)
  Adodc3.Recordset.Fields!asaskarkhane = tmp0(6)
  Adodc3.Recordset.Update
  Adodc3.Refresh
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
'  introwforall = Adodc1.Recordset.Fields!rad
'  Call DataGrid_sum_help
'  Adodc1.Refresh
'  Adodc1.Recordset.Find "rad=" + Trim(Str(introwforall)), , adSearchForward, 1

End Sub

Private Sub DataGrid2_AfterColUpdate(ByVal ColIndex As Integer)
If blnasd = True Then
  DataGrid2.Text = tmpasd
  DataGrid2.Refresh
End If
If Adodc2.Recordset.Fields!rad = 50 Then

End If
End Sub

Private Sub DataGrid2_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If Adodc2.Recordset.Fields!rad = 50 Then
  blnasd = False
  tmpasd = DataGrid2.Text
Else
  blnasd = True
  tmpasd = DataGrid2.Text
End If
End Sub

Private Sub Form_Activate()
Adodc3.ConnectionString = Form3.Text10.Text
Adodc3.CommandType = adCmdUnknown
Adodc3.RecordSource = "select * from Estehlak1 ORDER BY rad"
Adodc3.Refresh

Adodc4.ConnectionString = Form3.Text10.Text
Adodc4.CommandType = adCmdUnknown
Adodc4.RecordSource = "select * from Estehlak2 ORDER BY rad"
Adodc4.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
Me.Hide
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
  Adodc3.Recordset.Find "rad=999", , adSearchForward, 1
  tmp(0) = Adodc3.Recordset.Fields!zirbana
  
  Adodc4.Recordset.Find "rad=50", , adSearchForward, 1
  tmp(1) = Adodc4.Recordset.Fields!skarghahi
  
  Adodc4.Recordset.Find "rad=50", , adSearchForward, 1
  tmp(2) = Adodc4.Recordset.Fields!tasisgaz
  
  Adodc3.Recordset.Find "rad=999", , adSearchForward, 1
  tmp(3) = Adodc3.Recordset.Fields!tasisathararati
  
  Adodc4.Recordset.Find "rad=50", , adSearchForward, 1
  tmp(4) = Adodc4.Recordset.Fields!tashararat
  
  Adodc4.Recordset.Find "rad=50", , adSearchForward, 1
  tmp(5) = Adodc4.Recordset.Fields!tasab
  
  Adodc4.Recordset.Find "rad=50", , adSearchForward, 1
  tmp(6) = Adodc4.Recordset.Fields!tascool
  
  Adodc3.Recordset.Find "rad=999", , adSearchForward, 1
  tmp(7) = Adodc3.Recordset.Fields!kilovat
  
  Adodc4.Recordset.Find "rad=50", , adSearchForward, 1
  tmp(8) = Adodc4.Recordset.Fields!tasbargh
  
  Adodc3.Recordset.Find "rad=999", , adSearchForward, 1
  tmp(9) = Adodc3.Recordset.Fields!asaskarkhane
  
  Adodc4.Recordset.Find "rad=50", , adSearchForward, 1
  tmp(10) = Adodc4.Recordset.Fields!asasedari
  
  Adodc4.Recordset.Find "rad=50", , adSearchForward, 1
  tmp(11) = Adodc4.Recordset.Fields!vasayelvalefterak
  
  Adodc4.Recordset.MoveFirst
  Do
    If (Adodc4.Recordset.Fields!rad <> 39) And (Adodc4.Recordset.Fields!rad <> 50) And (Adodc4.Recordset.Fields!rad <> 999) Then
    
      Adodc3.Recordset.Find "rad=" + Trim(Str(Adodc4.Recordset.Fields!rad)), , adSearchForward, 1
      
      r1 = (Val(tmp(1)) / Val(tmp(0))) * Val(Adodc3.Recordset.Fields!zirbana)
      Adodc4.Recordset.Fields!skarghahi = Round(r1)
      
      r1 = (Val(tmp(2)) / Val(tmp(0))) * Val(Adodc3.Recordset.Fields!zirbana)
      Adodc4.Recordset.Fields!tasisgaz = Round(r1)
      
      r1 = (Val(tmp(4)) / Val(tmp(3))) * Val(Adodc3.Recordset.Fields!tasisathararati)
      Adodc4.Recordset.Fields!tashararat = Round(r1)
      
      r1 = (Val(tmp(5)) * Val(Adodc3.Recordset.Fields!abresani))
      Adodc4.Recordset.Fields!tasab = Round(r1)
      
      r1 = (Val(tmp(6)) * Val(Adodc3.Recordset.Fields!cooler))
      Adodc4.Recordset.Fields!tascool = Round(r1)
      
      r1 = (Val(tmp(7)) / Val(tmp(8))) * Val(Adodc3.Recordset.Fields!kilovat)
      Adodc4.Recordset.Fields!tasbargh = Round(r1)
      
      r1 = (Val(tmp(9)) / Val(tmp(10))) * Val(Adodc3.Recordset.Fields!asaskarkhane)
      Adodc4.Recordset.Fields!asasedari = Round(r1)
      
      r1 = (Val(tmp(11)) * Val(Adodc3.Recordset.Fields!naghleye))
      Adodc4.Recordset.Fields!vasayelvalefterak = Round(r1)
      
      Adodc4.Recordset.Fields!sumend = Val(Adodc4.Recordset.Fields!skarghahi) + Val(Adodc4.Recordset.Fields!tasisgaz) + Val(Adodc4.Recordset.Fields!tashararat) + Val(Adodc4.Recordset.Fields!tasab) + Val(Adodc4.Recordset.Fields!tascool) + Val(Adodc4.Recordset.Fields!tasbargh) + Val(Adodc4.Recordset.Fields!asasedari) + Val(Adodc4.Recordset.Fields!vasayelvalefterak)
      Adodc4.Recordset.Update
    End If
    
    Adodc4.Recordset.MoveNext
  Loop Until Adodc4.Recordset.EOF = True
  Adodc2.Refresh
  DataGrid2.Refresh
  Adodc2.Refresh
  DataGrid2.Refresh

End If
End Sub

