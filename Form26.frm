VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form26 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "”—»«—"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13425
   BeginProperty Font 
      Name            =   "B Titr"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form26.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form26"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   13425
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13245
      _ExtentX        =   23363
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "ÃœÊ· „Õ«”»Â «” Â·«ﬂ œ«—«∆ÌÂ«Ì À«»  "
      TabPicture(0)   =   "Form26.frx":2CFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   " Œ’Ì’ Â“Ì‰Â Â«Ì œÊ«Ì— Œœ„« Ì »Â  Ê·ÌœÌ "
      TabPicture(1)   =   "Form26.frx":2D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   " ”ÂÌ„ Â“Ì‰Â Â«Ì  Œœ„« Ì »Â Œœ„« Ì"
      TabPicture(2)   =   "Form26.frx":2D32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DataGrid2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "÷—«Ì»  ”ÂÌ„ ”—»«— "
      TabPicture(3)   =   "Form26.frx":2D4E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DataGrid1"
      Tab(3).ControlCount=   1
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form26.frx":2D6A
         Height          =   7935
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   13996
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   2
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
         ColumnCount     =   8
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
            DataField       =   "nafarat"
            Caption         =   "‰›—«  ÃÂ  ⁄„Ê„Ì ° —” Ê—«‰ Ê „œÌ—Ì "
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
            DataField       =   "kontrol_keyfi"
            Caption         =   "Ê“‰ „” ÃÂ  ﬂ‰ —· ﬂÌ›Ì"
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
            DataField       =   "hazvahedfani"
            Caption         =   "÷—Ì» ›‰Ì ÃÂ   ”ÂÌ„ Â“Ì‰Â Â«Ì Ê«Õœ›‰Ì"
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
            DataField       =   "roghankeshsh"
            Caption         =   "—Ê€‰ ﬂ‘‘"
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
            DataField       =   "masrafab"
            Caption         =   "÷—Ì» „’—› ¬»"
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
            DataField       =   "barghkilowat"
            Caption         =   "÷—Ì» „’—› »—ﬁ"
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
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Form26.frx":2D7F
         Height          =   7935
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   12975
         _ExtentX        =   22886
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
         ColumnCount     =   10
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
            Caption         =   "Ê«ÕœÂ«Ì Œœ„« Ì"
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
            DataField       =   "restoran"
            Caption         =   "—” Ê—«‰"
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
            DataField       =   "edari"
            Caption         =   "«œ«—Ì ﬂ«—Œ«‰Â"
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
            DataField       =   "omomi"
            Caption         =   "⁄„Ê„Ì"
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
            DataField       =   "kargahfani"
            Caption         =   "ﬂ«—ê«Â ›‰Ì  Ê·Ìœ"
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
            DataField       =   "kontrol"
            Caption         =   "ﬂ‰ —· ﬂÌ›Ì "
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
            DataField       =   "barghkilowat"
            Caption         =   "»—ﬁ „’—›Ì"
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
            DataField       =   "estehlak"
            Caption         =   "«” Â·«ﬂ"
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
            DataField       =   "sum"
            Caption         =   "Ã„⁄"
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
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "Form26.frx":2D94
         Height          =   7935
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   12975
         _ExtentX        =   22886
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
         ColumnCount     =   11
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
            DataField       =   "roghankeshsh"
            Caption         =   "—Ê€‰ ﬂ‘‘"
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
            DataField       =   "masterig"
            Caption         =   "„” —ÌÃ"
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
            DataField       =   "edari"
            Caption         =   "«œ«—Ì  Ê·Ìœ"
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
            DataField       =   "restoran"
            Caption         =   "—” Ê—«‰"
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
            DataField       =   "edarikarkhane"
            Caption         =   "«œ«—Ì ﬂ«—Œ«‰Â"
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
            DataField       =   "omomi"
            Caption         =   "⁄„Ê„Ì"
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
            DataField       =   "fani"
            Caption         =   "ﬂ«—ê«Â ›‰Ì"
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
            DataField       =   "kontrolkeyfi"
            Caption         =   "ﬂ‰ —· ﬂÌ›Ì"
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
            DataField       =   "sum"
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
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1800
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1800
            EndProperty
            BeginProperty Column10 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "Form26.frx":2DA9
         Height          =   7935
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   12975
         _ExtentX        =   22886
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
            DataField       =   "dastmozd"
            Caption         =   "œ” „“œ"
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
            DataField       =   "sarbarvahed"
            Caption         =   "”—»«— Ê«Õœ"
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
            DataField       =   "estehlak"
            Caption         =   "«” Â·«ﬂ"
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
            DataField       =   "sarbarjazb"
            Caption         =   "”—»«— Ã–» ‘œÂ"
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
            DataField       =   "sum"
            Caption         =   "Ã„⁄"
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
            DataField       =   "mavadvahed"
            Caption         =   "„Ê«œ «Ê·ÌÂ Ê«Õœ"
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
            DataField       =   "bahayevahed"
            Caption         =   "»Â«Ì  „«„ ‘œÂ Ê«Õœ"
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
            DataField       =   "naghlaz_vahedghabl"
            Caption         =   "Â“Ì‰Â ‰ﬁ· «“ Ê«Õœ ﬁ»·"
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
            DataField       =   "sumbahayetolid"
            Caption         =   "Ã„⁄ »Â«Ì  Ê·Ìœ Ê«Õœ"
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
            DataField       =   "kaladarjaryanavaldore"
            Caption         =   "Â“Ì‰Â ﬂ«·«Ì œ—Ã—Ì«‰ ”«Œ  «Ê· œÊ—Â"
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
            DataField       =   "amadebaraymasraf"
            Caption         =   "¬„«œÂ »—«Ì „’—›"
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
            DataField       =   "sahmhazvahedbad"
            Caption         =   "”Â„ Â“Ì‰Â ‰ﬁ· »Â Ê«Õœ »⁄œ"
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
            DataField       =   "zayeat"
            Caption         =   "÷«Ì⁄«  ÿÌ œÊ—Â"
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
            DataField       =   "hazkalapayandore"
            Caption         =   "Â“Ì‰Â ﬂ«·«Ì œ— Ã—Ì«‰ ”«Œ  Å«Ì«‰ œÊ—Â"
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
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   1440
      Top             =   4920
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
      RecordSource    =   "sarbar_4"
      Caption         =   "Adodc4"
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
      Left            =   1440
      Top             =   4440
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
      RecordSource    =   "sarbar_3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   240
      Top             =   4920
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
      RecordSource    =   "sarbar_2"
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
      Left            =   240
      Top             =   4440
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
      RecordSource    =   "sarbar_1"
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
Attribute VB_Name = "Form26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmp(30) As String, tmp0(30) As String, introwforall As Integer

Private Sub DataGrid_sum_help()
  For q = 0 To 20
    tmp0(q) = 0
  Next q
  
  Adodc4.Recordset.MoveFirst
  Do
    If (Adodc4.Recordset.Fields!rad <> "999") And (Adodc4.Recordset.Fields!rad <> "998") And (Adodc4.Recordset.Fields!rad <> "997") And (Adodc4.Recordset.Fields!dastmozd <> "") Then tmp0(0) = Val(tmp0(0)) + Val(Adodc4.Recordset.Fields!dastmozd)
    If (Adodc4.Recordset.Fields!rad <> "999") And (Adodc4.Recordset.Fields!rad <> "998") And (Adodc4.Recordset.Fields!rad <> "997") And (Adodc4.Recordset.Fields!sarbarvahed <> "") Then tmp0(1) = Val(tmp0(1)) + Val(Adodc4.Recordset.Fields!sarbarvahed)
    If (Adodc4.Recordset.Fields!rad <> "999") And (Adodc4.Recordset.Fields!rad <> "998") And (Adodc4.Recordset.Fields!rad <> "997") And (Adodc4.Recordset.Fields!sarbarjazb <> "") Then tmp0(2) = Val(tmp0(2)) + Val(Adodc4.Recordset.Fields!sarbarjazb)
    If (Adodc4.Recordset.Fields!rad <> "999") And (Adodc4.Recordset.Fields!rad <> "998") And (Adodc4.Recordset.Fields!rad <> "997") And (Adodc4.Recordset.Fields!Sum <> "") Then tmp0(3) = Val(tmp0(3)) + Val(Adodc4.Recordset.Fields!Sum)
    If (Adodc4.Recordset.Fields!rad <> "999") And (Adodc4.Recordset.Fields!rad <> "998") And (Adodc4.Recordset.Fields!rad <> "997") And (Adodc4.Recordset.Fields!mavadvahed <> "") Then tmp0(4) = Val(tmp0(4)) + Val(Adodc4.Recordset.Fields!mavadvahed)
    If (Adodc4.Recordset.Fields!rad <> "999") And (Adodc4.Recordset.Fields!rad <> "998") And (Adodc4.Recordset.Fields!rad <> "997") And (Adodc4.Recordset.Fields!bahayevahed <> "") Then tmp0(5) = Val(tmp0(5)) + Val(Adodc4.Recordset.Fields!bahayevahed)
    If (Adodc4.Recordset.Fields!rad <> "999") And (Adodc4.Recordset.Fields!rad <> "998") And (Adodc4.Recordset.Fields!rad <> "997") And (Adodc4.Recordset.Fields!naghlaz_vahedghabl <> "") Then tmp0(6) = Val(tmp0(6)) + Val(Adodc4.Recordset.Fields!naghlaz_vahedghabl)
    If (Adodc4.Recordset.Fields!rad <> "999") And (Adodc4.Recordset.Fields!rad <> "998") And (Adodc4.Recordset.Fields!rad <> "997") And (Adodc4.Recordset.Fields!sumbahayetolid <> "") Then tmp0(7) = Val(tmp0(7)) + Val(Adodc4.Recordset.Fields!sumbahayetolid)
    If (Adodc4.Recordset.Fields!rad <> "999") And (Adodc4.Recordset.Fields!rad <> "998") And (Adodc4.Recordset.Fields!rad <> "997") And (Adodc4.Recordset.Fields!kaladarjaryanavaldore <> "") Then tmp0(8) = Val(tmp0(8)) + Val(Adodc4.Recordset.Fields!kaladarjaryanavaldore)
    If (Adodc4.Recordset.Fields!rad <> "999") And (Adodc4.Recordset.Fields!rad <> "998") And (Adodc4.Recordset.Fields!rad <> "997") And (Adodc4.Recordset.Fields!amadebaraymasraf <> "") Then tmp0(9) = Val(tmp0(9)) + Val(Adodc4.Recordset.Fields!amadebaraymasraf)
    If (Adodc4.Recordset.Fields!rad <> "999") And (Adodc4.Recordset.Fields!rad <> "998") And (Adodc4.Recordset.Fields!rad <> "997") And (Adodc4.Recordset.Fields!sahmhazvahedbad <> "") Then tmp0(10) = Val(tmp0(10)) + Val(Adodc4.Recordset.Fields!sahmhazvahedbad)
    If (Adodc4.Recordset.Fields!rad <> "999") And (Adodc4.Recordset.Fields!rad <> "998") And (Adodc4.Recordset.Fields!rad <> "997") And (Adodc4.Recordset.Fields!zayeat <> "") Then tmp0(11) = Val(tmp0(11)) + Val(Adodc4.Recordset.Fields!zayeat)
    If (Adodc4.Recordset.Fields!rad <> "999") And (Adodc4.Recordset.Fields!rad <> "998") And (Adodc4.Recordset.Fields!rad <> "997") And (Adodc4.Recordset.Fields!hazkalapayandore <> "") Then tmp0(12) = Val(tmp0(12)) + Val(Adodc4.Recordset.Fields!hazkalapayandore)
    Adodc4.Recordset.MoveNext
  Loop Until Adodc4.Recordset.EOF = True
  
  
  Adodc4.Recordset.Find "rad=997", , adSearchForward, 1
  Adodc4.Recordset.Fields!dastmozd = tmp0(0)
  Adodc4.Recordset.Fields!sarbarvahed = tmp0(1)
  Adodc4.Recordset.Fields!sarbarjazb = tmp0(2)
  Adodc4.Recordset.Fields!Sum = tmp0(3)
  Adodc4.Recordset.Fields!mavadvahed = tmp0(4)
  Adodc4.Recordset.Fields!bahayevahed = tmp0(5)
  Adodc4.Recordset.Fields!naghlaz_vahedghabl = tmp0(6)
  Adodc4.Recordset.Fields!sumbahayetolid = tmp0(7)
  Adodc4.Recordset.Fields!kaladarjaryanavaldore = tmp0(8)
  Adodc4.Recordset.Fields!amadebaraymasraf = tmp0(9)
  Adodc4.Recordset.Fields!sahmhazvahedbad = tmp0(10)
  Adodc4.Recordset.Fields!zayeat = tmp0(11)
  Adodc4.Recordset.Fields!hazkalapayandore = tmp0(12)
  Adodc4.Recordset.Update
  DataGrid4.Refresh
  Adodc4.Recordset.Find "rad=" + Trim(Str(introwforall)), , adSearchForward, 1
End Sub

Private Sub DataGrid4_AfterColUpdate_help()
If (Adodc4.Recordset.Fields!rad <> 999) And (Adodc4.Recordset.Fields!rad <> 998) And (Adodc4.Recordset.Fields!rad <> 997) Then
  Form24.Adodc2.Recordset.Find "rad=" + Trim(Str(Adodc4.Recordset.Fields!rad)), , adSearchForward, 1
  Adodc4.Recordset.Fields!estehlak = Form24.Adodc2.Recordset.Fields!sumend
  
  Adodc3.Recordset.Find "rad=" + Trim(Str(Adodc4.Recordset.Fields!rad)), , adSearchForward, 1
  Adodc4.Recordset.Fields!sarbarjazb = Adodc3.Recordset.Fields!Sum
  
  If Adodc4.Recordset.Fields!dastmozd = "" Then Adodc4.Recordset.Fields!dastmozd = 0
  If Adodc4.Recordset.Fields!sarbarvahed = "" Then Adodc4.Recordset.Fields!sarbarvahed = 0
  
  Adodc4.Recordset.Fields!Sum = Val(Adodc4.Recordset.Fields!dastmozd) + Val(Adodc4.Recordset.Fields!sarbarvahed) + Val(Adodc4.Recordset.Fields!estehlak) + Val(Adodc4.Recordset.Fields!sarbarjazb)
  
  If Adodc4.Recordset.Fields!rad = 1 Then
    Form9.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
    Adodc4.Recordset.Fields!mavadvahed = Form9.Adodc1.Recordset.Fields!mavadaval
  ElseIf Adodc4.Recordset.Fields!rad = 11 Then
'    Form9.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
'    Adodc4.Recordset.Fields!mavadvahed = Form9.Adodc1.Recordset.Fields!mavadaval
  Else
    Adodc4.Recordset.Fields!mavadvahed = 0
  End If
  
  Adodc4.Recordset.Fields!bahayevahed = Val(Adodc4.Recordset.Fields!mavadvahed) + Val(Adodc4.Recordset.Fields!Sum)
  
  Select Case Adodc4.Recordset.Fields!rad
    Case 1
      Form9.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc4.Recordset.Fields!naghlaz_vahedghabl = 0
      Adodc4.Recordset.Fields!kaladarjaryanavaldore = Form9.Adodc1.Recordset.Fields!mojodiavalmemoney
      Adodc4.Recordset.Fields!sahmhazvahedbad = Form9.Adodc1.Recordset.Fields!naghlbebadmoney
      Form9.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc4.Recordset.Fields!zayeat = Form9.Adodc1.Recordset.Fields!naghlbebadmoney
      
    Case 2
      Form10.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc4.Recordset.Fields!naghlaz_vahedghabl = Form10.Adodc1.Recordset.Fields!mavadaval
      Adodc4.Recordset.Fields!kaladarjaryanavaldore = Form10.Adodc1.Recordset.Fields!mojodiavalmemoney
      Adodc4.Recordset.Fields!sahmhazvahedbad = Form10.Adodc1.Recordset.Fields!naghlbebadmoney
      Form10.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc4.Recordset.Fields!zayeat = Form10.Adodc1.Recordset.Fields!naghlbebadmoney
      
    Case 3
      Form11.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc4.Recordset.Fields!naghlaz_vahedghabl = Form11.Adodc1.Recordset.Fields!mavadaval
      Adodc4.Recordset.Fields!kaladarjaryanavaldore = Form11.Adodc1.Recordset.Fields!mojodiavalmemoney
      Adodc4.Recordset.Fields!sahmhazvahedbad = Form11.Adodc1.Recordset.Fields!naghlbebadmoney
      Form11.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc4.Recordset.Fields!zayeat = Form11.Adodc1.Recordset.Fields!naghlbebadmoney

    Case 4
      Form13.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc4.Recordset.Fields!naghlaz_vahedghabl = Form13.Adodc1.Recordset.Fields!mavadaval
      Adodc4.Recordset.Fields!kaladarjaryanavaldore = Form13.Adodc1.Recordset.Fields!mojodiavalmemoney
      Adodc4.Recordset.Fields!sahmhazvahedbad = Form13.Adodc1.Recordset.Fields!naghlbebadmoney
      Form13.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc4.Recordset.Fields!zayeat = Form13.Adodc1.Recordset.Fields!naghlbebadmoney

    Case 5
      Form1.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc4.Recordset.Fields!naghlaz_vahedghabl = Form1.Adodc1.Recordset.Fields!mavadaval
      Adodc4.Recordset.Fields!kaladarjaryanavaldore = Form1.Adodc1.Recordset.Fields!mojodiavalmemoney
      Adodc4.Recordset.Fields!sahmhazvahedbad = Form1.Adodc1.Recordset.Fields!naghlbebadmoney
      Form1.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc4.Recordset.Fields!zayeat = Form1.Adodc1.Recordset.Fields!naghlbebadmoney

    Case 6
      Form14.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc4.Recordset.Fields!naghlaz_vahedghabl = Form14.Adodc1.Recordset.Fields!mavadaval
      Adodc4.Recordset.Fields!kaladarjaryanavaldore = Form14.Adodc1.Recordset.Fields!mojodiavalmemoney
      Adodc4.Recordset.Fields!sahmhazvahedbad = Form14.Adodc1.Recordset.Fields!naghlbebadmoney
      Form14.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc4.Recordset.Fields!zayeat = Form14.Adodc1.Recordset.Fields!naghlbebadmoney

    Case 7
      Form16.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc4.Recordset.Fields!naghlaz_vahedghabl = Form16.Adodc1.Recordset.Fields!mavadaval
      Adodc4.Recordset.Fields!kaladarjaryanavaldore = Form16.Adodc1.Recordset.Fields!mojodiavalmemoney
      Adodc4.Recordset.Fields!sahmhazvahedbad = Form16.Adodc1.Recordset.Fields!naghlbebadmoney
      Form16.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc4.Recordset.Fields!zayeat = Form16.Adodc1.Recordset.Fields!naghlbebadmoney
      
    Case 8
      Form17.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc4.Recordset.Fields!naghlaz_vahedghabl = Form17.Adodc1.Recordset.Fields!mavadaval
      Adodc4.Recordset.Fields!kaladarjaryanavaldore = Form17.Adodc1.Recordset.Fields!mojodiavalmemoney
      Adodc4.Recordset.Fields!sahmhazvahedbad = Form17.Adodc1.Recordset.Fields!naghlbebadmoney
      Form17.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc4.Recordset.Fields!zayeat = Form17.Adodc1.Recordset.Fields!naghlbebadmoney

    Case 9
      Form18.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc4.Recordset.Fields!naghlaz_vahedghabl = Form18.Adodc1.Recordset.Fields!mavadaval
      Adodc4.Recordset.Fields!kaladarjaryanavaldore = Form18.Adodc1.Recordset.Fields!mojodiavalmemoney
      Adodc4.Recordset.Fields!sahmhazvahedbad = Form18.Adodc1.Recordset.Fields!naghlbebadmoney
      Form18.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc4.Recordset.Fields!zayeat = Form18.Adodc1.Recordset.Fields!naghlbebadmoney
      
    Case 10
      Form19.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc4.Recordset.Fields!naghlaz_vahedghabl = Form19.Adodc1.Recordset.Fields!mavadaval
      Adodc4.Recordset.Fields!kaladarjaryanavaldore = Form19.Adodc1.Recordset.Fields!mojodiavalmemoney
      Adodc4.Recordset.Fields!sahmhazvahedbad = Form19.Adodc1.Recordset.Fields!naghlbebadmoney
      Form19.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc4.Recordset.Fields!zayeat = Form19.Adodc1.Recordset.Fields!naghlbebadmoney
      
    Case 11
      Form20.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc4.Recordset.Fields!naghlaz_vahedghabl = Form20.Adodc1.Recordset.Fields!mavadaval
      Adodc4.Recordset.Fields!kaladarjaryanavaldore = Form20.Adodc1.Recordset.Fields!mojodiavalmemoney
      Adodc4.Recordset.Fields!sahmhazvahedbad = Form20.Adodc1.Recordset.Fields!naghlbebadmoney
      Form20.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc4.Recordset.Fields!zayeat = Form20.Adodc1.Recordset.Fields!naghlbebadmoney
      
    Case 12
      Form21.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc4.Recordset.Fields!naghlaz_vahedghabl = Form21.Adodc1.Recordset.Fields!mavadaval
      Adodc4.Recordset.Fields!kaladarjaryanavaldore = Form21.Adodc1.Recordset.Fields!mojodiavalmemoney
      Adodc4.Recordset.Fields!sahmhazvahedbad = Form21.Adodc1.Recordset.Fields!naghlbebadmoney
      Form21.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc4.Recordset.Fields!zayeat = Form21.Adodc1.Recordset.Fields!naghlbebadmoney

    Case 13
      Form22.Adodc1.Recordset.Find "rad=99999", , adSearchForward, 1
      Adodc4.Recordset.Fields!naghlaz_vahedghabl = Form22.Adodc1.Recordset.Fields!mavadaval
      Adodc4.Recordset.Fields!kaladarjaryanavaldore = Form22.Adodc1.Recordset.Fields!mojodiavalmemoney
      Adodc4.Recordset.Fields!sahmhazvahedbad = Form22.Adodc1.Recordset.Fields!naghlbebadmoney
      Form22.Adodc1.Recordset.Find "rad=99998", , adSearchForward, 1
      Adodc4.Recordset.Fields!zayeat = Form22.Adodc1.Recordset.Fields!naghlbebadmoney
      
  End Select
  Adodc4.Recordset.Fields!sumbahayetolid = Val(Adodc4.Recordset.Fields!naghlaz_vahedghabl) + Val(Adodc4.Recordset.Fields!bahayevahed)
  Adodc4.Recordset.Fields!amadebaraymasraf = Val(Adodc4.Recordset.Fields!kaladarjaryanavaldore) + Val(Adodc4.Recordset.Fields!sumbahayetolid)
  Adodc4.Recordset.Fields!hazkalapayandore = Adodc4.Recordset.Fields!amadebaraymasraf - Adodc4.Recordset.Fields!sahmhazvahedbad - Adodc4.Recordset.Fields!zayeat
  Adodc4.Recordset.Update
  
  Call DataGrid_sum_help
End If
End Sub
  
Private Sub DataGrid4_AfterColEdit(ByVal ColIndex As Integer)
  introwforall = Adodc4.Recordset.Fields!rad
  Call DataGrid4_AfterColUpdate_help
End Sub

Private Sub Form_Activate()
Adodc1.Recordset.Sort = "rad"
Adodc2.Recordset.Sort = "rad"
Adodc3.Recordset.Sort = "rad"
Adodc4.Recordset.Sort = "rad"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
Me.Hide
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 2 Then
  Adodc2.Recordset.Find "rad=998", , adSearchForward, 1
  tmp(0) = Adodc2.Recordset.Fields!restoran
  
  Adodc1.Recordset.Find "rad=999", , adSearchForward, 1
  tmp(1) = Adodc1.Recordset.Fields!nafarat
  
  Adodc1.Recordset.Find "rad=30", , adSearchForward, 1
  tmp(2) = Adodc1.Recordset.Fields!nafarat
  
  Adodc2.Recordset.Find "rad=998", , adSearchForward, 1
  tmp(3) = Adodc2.Recordset.Fields!edari
  
  Adodc1.Recordset.Find "rad=31", , adSearchForward, 1
  tmp(4) = Adodc1.Recordset.Fields!nafarat
  
  Adodc2.Recordset.Find "rad=998", , adSearchForward, 1
  tmp(5) = Adodc2.Recordset.Fields!omomi
  
  Adodc1.Recordset.Find "rad=32", , adSearchForward, 1
  tmp(6) = Adodc1.Recordset.Fields!nafarat
  
  Adodc2.Recordset.Find "rad=998", , adSearchForward, 1
  tmp(7) = Adodc2.Recordset.Fields!kargahfani
  
  Adodc1.Recordset.Find "rad=999", , adSearchForward, 1
  tmp(8) = Adodc1.Recordset.Fields!hazvahedfani
  
  Adodc1.Recordset.Find "rad=33", , adSearchForward, 1
  tmp(9) = Adodc1.Recordset.Fields!hazvahedfani
  
  Adodc2.Recordset.MoveFirst
  Do
    If (Adodc2.Recordset.Fields!rad <> 998) And (Adodc2.Recordset.Fields!rad <> 999) Then
      '—” Ê—«‰
      If (Adodc2.Recordset.Fields!rad <> 30) Then
        Adodc1.Recordset.Find "rad=" + Trim(Str(Adodc2.Recordset.Fields!rad)), , adSearchForward, 1
        r1 = (Val(tmp(0)) / (Val(tmp(1)) - Val(tmp(2)))) * Val(Adodc1.Recordset.Fields!nafarat)
        Adodc2.Recordset.Fields!restoran = Round(r1)
      End If
    
      '«œ«—Ì
      If (Adodc2.Recordset.Fields!rad <> 31) Then
        Adodc1.Recordset.Find "rad=" + Trim(Str(Adodc2.Recordset.Fields!rad)), , adSearchForward, 1
        r1 = (Val(tmp(3)) / (Val(tmp(1)) - Val(tmp(4)))) * Val(Adodc1.Recordset.Fields!nafarat)
        Adodc2.Recordset.Fields!edari = Round(r1)
      End If
    
      '⁄„Ê„Ì
      If (Adodc2.Recordset.Fields!rad <> 32) Then
        Adodc1.Recordset.Find "rad=" + Trim(Str(Adodc2.Recordset.Fields!rad)), , adSearchForward, 1
        r1 = (Val(tmp(5)) / (Val(tmp(1)) - Val(tmp(6)))) * Val(Adodc1.Recordset.Fields!nafarat)
        Adodc2.Recordset.Fields!omomi = Round(r1)
      End If
    
      'ﬂ«—ê«Â ›‰Ì
      If (Adodc2.Recordset.Fields!rad <> 33) Then
        Adodc1.Recordset.Find "rad=" + Trim(Str(Adodc2.Recordset.Fields!rad)), , adSearchForward, 1
        r1 = (Val(tmp(7)) / (Val(tmp(8)) - Val(tmp(9)))) * Val(Adodc1.Recordset.Fields!hazvahedfani)
        Adodc2.Recordset.Fields!kargahfani = Round(r1)
      End If
        
      '«” Â·«ﬂ
      If (Adodc2.Recordset.Fields!rad <> 30) And (Adodc2.Recordset.Fields!rad <> 998) And (Adodc2.Recordset.Fields!rad <> 999) Then
        Form24.Adodc2.Recordset.Find "rad=" + Trim(Str(Adodc2.Recordset.Fields!rad)), , adSearchForward, 1
        Adodc2.Recordset.Fields!estehlak = Form24.Adodc2.Recordset.Fields!sumend
      End If
      
      'Ã„⁄
      Adodc2.Recordset.Fields!Sum = Val(Adodc2.Recordset.Fields!restoran) + Val(Adodc2.Recordset.Fields!edari) + Val(Adodc2.Recordset.Fields!omomi) + Val(Adodc2.Recordset.Fields!kargahfani) + Val(Adodc2.Recordset.Fields!estehlak)
      Adodc2.Recordset.Update
    End If
    Adodc2.Recordset.MoveNext
  Loop Until Adodc2.Recordset.EOF = True
End If

If SSTab1.Tab = 1 Then
  Adodc3.Recordset.Find "rad=996", , adSearchForward, 1
  
  '—Ê€‰ ﬂ‘‘
  tmp(12) = Adodc3.Recordset.Fields!roghankeshsh
  
  '«œ«—Ì  Ê·Ìœ
  tmp(13) = Adodc3.Recordset.Fields!edari
  
  '—” Ê—«‰
  Adodc2.Recordset.Find "rad=998", , adSearchForward, 1
  tmp(0) = Adodc2.Recordset.Fields!restoran
  Adodc2.Recordset.Find "rad=30", , adSearchForward, 1
  tmp(1) = Adodc2.Recordset.Fields!Sum
  tmp(2) = Adodc2.Recordset.Fields!Sum
  Adodc3.Recordset.Fields!restoran = tmp(0) - tmp(1)
  tmp(7) = Val(Adodc3.Recordset.Fields!restoran) + Val(tmp(2))
  
  '«œ«—Ì ﬂ«—Œ«‰Â
  Adodc2.Recordset.Find "rad=998", , adSearchForward, 1
  tmp(0) = Adodc2.Recordset.Fields!edari
  Adodc2.Recordset.Find "rad=31", , adSearchForward, 1
  tmp(1) = Adodc2.Recordset.Fields!Sum
  tmp(3) = Adodc2.Recordset.Fields!Sum
  Adodc3.Recordset.Fields!edarikarkhane = tmp(0) - tmp(1)
  tmp(8) = Val(Adodc3.Recordset.Fields!edarikarkhane) + Val(tmp(3))
  
  '⁄„Ê„Ì
  Adodc2.Recordset.Find "rad=998", , adSearchForward, 1
  tmp(0) = Adodc2.Recordset.Fields!omomi
  Adodc2.Recordset.Find "rad=32", , adSearchForward, 1
  tmp(1) = Adodc2.Recordset.Fields!Sum
  tmp(4) = Adodc2.Recordset.Fields!Sum
  Adodc3.Recordset.Fields!omomi = tmp(0) - tmp(1)
  tmp(9) = Val(Adodc3.Recordset.Fields!omomi) + Val(tmp(4))
  
  'ﬂ«—ê«Â ›‰Ì
  Adodc2.Recordset.Find "rad=998", , adSearchForward, 1
  tmp(0) = Adodc2.Recordset.Fields!kargahfani
  Adodc2.Recordset.Find "rad=33", , adSearchForward, 1
  tmp(1) = Adodc2.Recordset.Fields!Sum
  tmp(5) = Adodc2.Recordset.Fields!Sum
  Adodc3.Recordset.Fields!fani = tmp(0) - tmp(1)
  tmp(10) = Val(Adodc3.Recordset.Fields!fani) + Val(tmp(5))
  
  'ﬂ‰ —· ﬂÌ›Ì 
  Adodc2.Recordset.Find "rad=998", , adSearchForward, 1
  tmp(0) = Adodc2.Recordset.Fields!kontrol
  Adodc2.Recordset.Find "rad=34", , adSearchForward, 1
  tmp(1) = Adodc2.Recordset.Fields!Sum
  tmp(6) = Adodc2.Recordset.Fields!Sum
  Adodc3.Recordset.Fields!kontrolkeyfi = tmp(0) - tmp(1)
  tmp(11) = Val(Adodc3.Recordset.Fields!kontrolkeyfi) + Val(tmp(6))
  
  Adodc3.Recordset.Fields!Sum = Val(Adodc3.Recordset.Fields!restoran) + Val(Adodc3.Recordset.Fields!edarikarkhane) + Val(Adodc3.Recordset.Fields!omomi) + Val(Adodc3.Recordset.Fields!fani) + Val(Adodc3.Recordset.Fields!kontrolkeyfi)
  Adodc3.Recordset.Update
  
  Adodc3.Recordset.Find "rad=997", , adSearchForward, 1
  Adodc3.Recordset.Fields!restoran = tmp(2)
  Adodc3.Recordset.Fields!edarikarkhane = tmp(3)
  Adodc3.Recordset.Fields!omomi = tmp(4)
  Adodc3.Recordset.Fields!fani = tmp(5)
  Adodc3.Recordset.Fields!kontrolkeyfi = tmp(6)
  Adodc3.Recordset.Fields!Sum = Val(Adodc3.Recordset.Fields!restoran) + Val(Adodc3.Recordset.Fields!edarikarkhane) + Val(Adodc3.Recordset.Fields!omomi) + Val(Adodc3.Recordset.Fields!fani) + Val(Adodc3.Recordset.Fields!kontrolkeyfi)
  Adodc3.Recordset.Update

  Adodc3.Recordset.Find "rad=998", , adSearchForward, 1
  Adodc3.Recordset.Fields!restoran = tmp(7)
  Adodc3.Recordset.Fields!edarikarkhane = tmp(8)
  Adodc3.Recordset.Fields!omomi = tmp(9)
  Adodc3.Recordset.Fields!fani = tmp(10)
  Adodc3.Recordset.Fields!kontrolkeyfi = tmp(11)
  Adodc3.Recordset.Fields!roghankeshsh = tmp(12)
  Adodc3.Recordset.Fields!edari = tmp(13)
  Adodc3.Recordset.Fields!Sum = Val(Adodc3.Recordset.Fields!restoran) + Val(Adodc3.Recordset.Fields!edarikarkhane) + Val(Adodc3.Recordset.Fields!omomi) + Val(Adodc3.Recordset.Fields!fani) + Val(Adodc3.Recordset.Fields!kontrolkeyfi)
  Adodc3.Recordset.Update

  Adodc1.Recordset.Find "rad=29", , adSearchForward, 1
  tmp(14) = Adodc1.Recordset.Fields!roghankeshsh
  
  Adodc1.Recordset.Find "rad=29", , adSearchForward, 1
  tmp(15) = Adodc1.Recordset.Fields!nafarat
  
  Adodc1.Recordset.Find "rad=29", , adSearchForward, 1
  tmp(16) = Adodc1.Recordset.Fields!hazvahedfani
  
  Adodc1.Recordset.Find "rad=29", , adSearchForward, 1
  tmp(17) = Adodc1.Recordset.Fields!kontrol_keyfi
  
  Adodc3.Recordset.MoveFirst
  Do
    If (Adodc3.Recordset.Fields!rad <> 999) And (Adodc3.Recordset.Fields!rad <> 998) And (Adodc3.Recordset.Fields!rad <> 997) And (Adodc3.Recordset.Fields!rad <> 996) Then
      '—Ê€‰ ﬂ‘‘
      Adodc1.Recordset.Find "rad=" + Trim(Str(Adodc3.Recordset.Fields!rad)), , adSearchForward, 1
      r1 = (Val(tmp(12)) / Val(tmp(14))) * Val(Adodc1.Recordset.Fields!roghankeshsh)
      Adodc3.Recordset.Fields!roghankeshsh = Round(r1)
      
      '«œ«—Ì  Ê·Ìœ
      Adodc1.Recordset.Find "rad=" + Trim(Str(Adodc3.Recordset.Fields!rad)), , adSearchForward, 1
      r1 = (Val(tmp(13)) / Val(tmp(15))) * Val(Adodc1.Recordset.Fields!nafarat)
      Adodc3.Recordset.Fields!edari = Round(r1)
      
      '—” Ê—«‰
      Adodc1.Recordset.Find "rad=" + Trim(Str(Adodc3.Recordset.Fields!rad)), , adSearchForward, 1
      r1 = (Val(tmp(7)) / Val(tmp(15))) * Val(Adodc1.Recordset.Fields!nafarat)
      Adodc3.Recordset.Fields!restoran = Round(r1)
  
      '«œ«—Ì ﬂ«—Œ«‰Â
      Adodc1.Recordset.Find "rad=" + Trim(Str(Adodc3.Recordset.Fields!rad)), , adSearchForward, 1
      r1 = (Val(tmp(8)) / Val(tmp(15))) * Val(Adodc1.Recordset.Fields!nafarat)
      Adodc3.Recordset.Fields!edarikarkhane = Round(r1)

      '⁄„Ê„Ì
      Adodc1.Recordset.Find "rad=" + Trim(Str(Adodc3.Recordset.Fields!rad)), , adSearchForward, 1
      r1 = (Val(tmp(9)) / Val(tmp(15))) * Val(Adodc1.Recordset.Fields!nafarat)
      Adodc3.Recordset.Fields!omomi = Round(r1)
  
      'ﬂ«—ê«Â ›‰Ì
      Adodc1.Recordset.Find "rad=" + Trim(Str(Adodc3.Recordset.Fields!rad)), , adSearchForward, 1
      r1 = (Val(tmp(10)) / Val(tmp(16))) * Val(Adodc1.Recordset.Fields!hazvahedfani)
      Adodc3.Recordset.Fields!fani = Round(r1)
  
      'ﬂ‰ —· ﬂÌ›Ì 
      Adodc1.Recordset.Find "rad=" + Trim(Str(Adodc3.Recordset.Fields!rad)), , adSearchForward, 1
      r1 = (Val(tmp(11)) / Val(tmp(17))) * Val(Adodc1.Recordset.Fields!kontrol_keyfi)
      Adodc3.Recordset.Fields!kontrolkeyfi = Round(r1)
      
      Adodc3.Recordset.Fields!Sum = Val(Adodc3.Recordset.Fields!roghankeshsh) + Val(Adodc3.Recordset.Fields!edari) + Val(Adodc3.Recordset.Fields!restoran) + Val(Adodc3.Recordset.Fields!edarikarkhane) + Val(Adodc3.Recordset.Fields!omomi) + Val(Adodc3.Recordset.Fields!fani) + Val(Adodc3.Recordset.Fields!kontrolkeyfi)
      Adodc3.Recordset.Update
      
    End If
    Adodc3.Recordset.MoveNext
  Loop Until Adodc3.Recordset.EOF = True
End If
End Sub

