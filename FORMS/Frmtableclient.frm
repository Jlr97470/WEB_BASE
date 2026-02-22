VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTableClient 
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Tag             =   "4100"
   Begin VB.Frame fraTableChamp 
      Height          =   4275
      Left            =   0
      TabIndex        =   2
      Tag             =   "4102"
      Top             =   720
      Width           =   9615
      Begin VB.TextBox txtChampValeur 
         DataField       =   "CliMotDePasse"
         DataMember      =   "DEcmdTblClient"
         DataSource      =   "DEWebBase"
         Height          =   285
         Index           =   2
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   30
         Tag             =   "4113"
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtChampValeur 
         DataField       =   "CliEMail"
         DataMember      =   "DEcmdTblClient"
         DataSource      =   "DEWebBase"
         Height          =   285
         Index           =   1
         Left            =   6600
         TabIndex        =   17
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtChampValeur 
         DataField       =   "CliAdresse"
         DataMember      =   "DEcmdTblClient"
         DataSource      =   "DEWebBase"
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   15
         Tag             =   "4117"
         Top             =   2160
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mebChampValeur 
         Bindings        =   "Frmtableclient.frx":0000
         DataField       =   "CliLogin"
         DataMember      =   "DEcmdTblClient"
         DataSource      =   "DEWebBase"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   11
         Tag             =   "4112"
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebChampValeur 
         Bindings        =   "Frmtableclient.frx":003E
         DataField       =   "CliNom"
         DataMember      =   "DEcmdTblClient"
         DataSource      =   "DEWebBase"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   12
         Tag             =   "4114"
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebChampValeur 
         Bindings        =   "Frmtableclient.frx":007C
         DataField       =   "CliPrenom"
         DataMember      =   "DEcmdTblClient"
         DataSource      =   "DEWebBase"
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   13
         Tag             =   "4115"
         Top             =   1440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebChampValeur 
         Bindings        =   "Frmtableclient.frx":00BA
         DataField       =   "CliDateNai"
         DataMember      =   "DEcmdTblClient"
         DataSource      =   "DEWebBase"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   14
         Tag             =   "4116"
         Top             =   1800
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebChampValeur 
         Bindings        =   "Frmtableclient.frx":00F8
         DataField       =   "CliVilNum"
         DataMember      =   "DEcmdTblClient"
         DataSource      =   "DEWebBase"
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   16
         Tag             =   "4118"
         Top             =   2640
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label lblFieldLabel 
         Caption         =   "MAIL"
         Height          =   255
         Index           =   0
         Left            =   5040
         TabIndex        =   10
         Tag             =   "4110"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "VILLE:"
         Height          =   195
         Index           =   7
         Left            =   1470
         TabIndex        =   9
         Tag             =   "4109"
         Top             =   2640
         Width           =   480
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ADRESSE:"
         Height          =   195
         Index           =   6
         Left            =   1140
         TabIndex        =   8
         Tag             =   "4108"
         Top             =   2265
         Width           =   810
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "DATE NAISSANCE:"
         Height          =   195
         Index           =   5
         Left            =   510
         TabIndex        =   7
         Tag             =   "4107"
         Top             =   1890
         Width           =   1440
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PRENOM:"
         Height          =   195
         Index           =   4
         Left            =   1200
         TabIndex        =   6
         Tag             =   "4106"
         Top             =   1500
         Width           =   750
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "NOM:"
         Height          =   195
         Index           =   3
         Left            =   1530
         TabIndex        =   5
         Tag             =   "4105"
         Top             =   1125
         Width           =   420
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "MOT DE PASSE:"
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   4
         Tag             =   "4104"
         Top             =   750
         Width           =   1245
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "LOGIN:"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   3
         Tag             =   "4103"
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame fraTableCommand 
      Height          =   1095
      Left            =   0
      TabIndex        =   18
      Tag             =   "4122"
      Top             =   7440
      Width           =   9615
      Begin VB.CommandButton cmdBouton 
         Caption         =   "VALIDER"
         Height          =   735
         Index           =   7
         Left            =   7440
         TabIndex        =   26
         Tag             =   "4130"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "EDITER"
         Height          =   735
         Index           =   6
         Left            =   6240
         TabIndex        =   25
         Tag             =   "4129"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "-"
         Height          =   735
         Index           =   5
         Left            =   5040
         TabIndex        =   24
         Tag             =   "4128"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "+"
         Height          =   735
         Index           =   4
         Left            =   4080
         TabIndex        =   23
         Tag             =   "4127"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   ">>|"
         Height          =   735
         Index           =   3
         Left            =   3000
         TabIndex        =   22
         Tag             =   "4126"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   ">"
         Height          =   735
         Index           =   2
         Left            =   2040
         TabIndex        =   21
         Tag             =   "4125"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "<"
         Height          =   735
         Index           =   1
         Left            =   1080
         TabIndex        =   20
         Tag             =   "4124"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "|<<"
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Tag             =   "4123"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "ANNULER"
         Height          =   735
         Index           =   8
         Left            =   8400
         TabIndex        =   27
         Tag             =   "4131"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraTableNom 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Tag             =   "4100"
      Top             =   0
      Width           =   9615
      Begin VB.Label lblTableNom 
         Alignment       =   2  'Center
         Caption         =   "CLIENT"
         Height          =   375
         Left            =   60
         TabIndex        =   1
         Tag             =   "4101"
         Top             =   225
         Width           =   9495
      End
   End
   Begin MSDataGridLib.DataGrid dgdTable 
      Bindings        =   "Frmtableclient.frx":0136
      Height          =   2415
      Left            =   15
      TabIndex        =   28
      Tag             =   "3113"
      Top             =   5025
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Enabled         =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "DEcmdTblClient"
      ColumnCount     =   25
      BeginProperty Column00 
         DataField       =   "CliNum"
         Caption         =   "CliNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "CliCliNum"
         Caption         =   "CliCliNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "CliLogin"
         Caption         =   "CliLogin"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "CliMotDePasse"
         Caption         =   "CliMotDePasse"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "CliDateCre"
         Caption         =   "CliDateCre"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "CliDateMaj"
         Caption         =   "CliDateMaj"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "CliDateAcc"
         Caption         =   "CliDateAcc"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "CliActif"
         Caption         =   "CliActif"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "CliNom"
         Caption         =   "CliNom"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "CliPrenom"
         Caption         =   "CliPrenom"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "CliCliSexCode"
         Caption         =   "CliCliSexCode"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "CliCliSitCode"
         Caption         =   "CliCliSitCode"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "CliCliFonNum"
         Caption         =   "CliCliFonNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "CliCliSecNum"
         Caption         =   "CliCliSecNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "CliDateNai"
         Caption         =   "CliDateNai"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "CliTaille"
         Caption         =   "CliTaille"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "CliPoids"
         Caption         =   "CliPoids"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "CliCliGrpSanCode"
         Caption         =   "CliCliGrpSanCode"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "CliAdresse"
         Caption         =   "CliAdresse"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column19 
         DataField       =   "CliVilNum"
         Caption         =   "CliVilNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column20 
         DataField       =   "CliTelephone"
         Caption         =   "CliTelephone"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column21 
         DataField       =   "CliPortable"
         Caption         =   "CliPortable"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column22 
         DataField       =   "CliEMail"
         Caption         =   "CliEMail"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column23 
         DataField       =   "CliSiteWebURL"
         Caption         =   "CliSiteWebURL"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column24 
         DataField       =   "CliNote"
         Caption         =   "CliNote"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowFocus      =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1260,284
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1289,764
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   1289,764
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox mebChampValeur 
      Bindings        =   "Frmtableclient.frx":014E
      DataField       =   "CliLogin"
      DataMember      =   "DEcmdTblClient"
      DataSource      =   "DEWebBase"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   29
      Tag             =   "4112"
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "frmTableClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    Delta Copyright                                                             (31/10/2000)  ***
'******************************************************************************
'***    FORM:                                                                                              ***
'***        frmTableClient                                                                           ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    PROGRAMMEUR:                                                                              ***
'***      Royer Jean-Laurent                                                                         ***
'******************************************************************************

'******************************************************************************
'***    MODIF :                                                                                            ***
'***      Version 1.0 : 30/10/2000 :                                                                ***
'******************************************************************************
Option Explicit                                                                                               ' Je doit etre sur que mes variables on ete declarer

'******************************************************************************
'***    Declaration De Constante Privee                                                       ***
'******************************************************************************

'******************************************************************************
'***    Constante Qui Defini Les Libelles De La feuille En Erreur                   ***
'******************************************************************************
Private Const mconFeuilleType = FEUILLEFORM                                                       ' Le type de feuille
Private Const mconFeuilleNom = "frmTableClient"                                          ' Le nom de la Feuille

'******************************************************************************
'***    Evenement                                                                                       ***
'******************************************************************************

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'***      Form_Load()                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***        - Mise En Form Des Controls De La Feuille                                     ***
'***        - Initialisation Des Controls De La Feuille                                      ***
'***        - Definition Du Mode De Saisie                                                      ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
Private Sub Form_Load()
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo Form_Load_Erreur
    ' Je Centre la feuille.
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    ' J'Initialise Les Controls De La Feuille
    InitialiseControl frmTableClient
   ' Je Desactive Les Controls De Saissie Des Champs De La Frame De Saissie
    EditeControl frmTableClient, False
    ' J'Active La DataGrid Pour La Selection D'un Enregistrement
    dgdTable.Enabled = True
   ' Je Defini Les Boutons De La Feuille En Mode Deplacement
    ValideBouton cmdBouton, VALIDEDEPLACEMENT
    ' Fin
Form_Load_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
Form_Load_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, FEUILLEFORM, mconFeuilleNom, INSTRUCTIONEVENEMENT, "Form_Load", vbNullString, Err
    ' Je Continue
    Resume Form_Load_Exit
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
Private Sub cmdBouton_Click(Index As Integer)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo cmdBouton_Click_Erreur
    ' Je Valide L'Action Sur Les Boutons De Commande
    ValideFormBouton Me, DEWebBase.rsDEcmdTblClient, Index
    ' Fin
cmdBouton_Click_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
cmdBouton_Click_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, FEUILLEFORM, mconFeuilleNom, INSTRUCTIONEVENEMENT, "cmdBouton_Click", CStr(Index), Err
    ' Je Continue
    Resume cmdBouton_Click_Exit
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'***        mebChampValeur_GotFocus(Index As Integer)                                ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***                                                                                                            ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Index  - Numero Du MaskedBox Qui A Pris Le Focus                         ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
Private Sub mebChampValeur_GotFocus(Index As Integer)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo mebChampValeur_GotFocus_Erreur
    ' Je positionne le point d'insertion
    mebChampValeur(Index).SelStart = 0
    ' Fin
mebChampValeur_GotFocus_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
mebChampValeur_GotFocus_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, FEUILLEFORM, mconFeuilleNom, INSTRUCTIONEVENEMENT, "mebChampValeur_GotFocus", CStr(Index), Err
    ' Je Continue
    Resume mebChampValeur_GotFocus_Exit
    ' Fin
End Sub

