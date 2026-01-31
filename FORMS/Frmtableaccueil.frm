VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTableAccueil 
   ClientHeight    =   13590
   ClientLeft      =   4785
   ClientTop       =   2910
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   13590
   ScaleWidth      =   9810
   Tag             =   "3100"
   Begin VB.Frame fraTableCommand 
      Height          =   1095
      Left            =   0
      TabIndex        =   12
      Tag             =   "3114"
      Top             =   12360
      Width           =   9615
      Begin VB.CommandButton cmdBouton 
         Caption         =   "|<<"
         Height          =   735
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Tag             =   "3115"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "<"
         Height          =   735
         Index           =   1
         Left            =   1200
         TabIndex        =   14
         Tag             =   "3116"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   ">"
         Height          =   735
         Index           =   2
         Left            =   2160
         TabIndex        =   15
         Tag             =   "3117"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   ">>|"
         Height          =   735
         Index           =   3
         Left            =   3120
         TabIndex        =   16
         Tag             =   "3118"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "+"
         Height          =   735
         Index           =   4
         Left            =   4200
         TabIndex        =   17
         Tag             =   "3119"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "-"
         Height          =   735
         Index           =   5
         Left            =   5160
         TabIndex        =   18
         Tag             =   "3120"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "EDITER"
         Height          =   735
         Index           =   6
         Left            =   6360
         TabIndex        =   19
         Tag             =   "3121"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "VALIDER"
         Height          =   735
         Index           =   7
         Left            =   7560
         TabIndex        =   20
         Tag             =   "3122"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "ANNULER"
         Height          =   735
         Index           =   8
         Left            =   8520
         TabIndex        =   21
         Tag             =   "3123"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraTableChamp 
      Height          =   9195
      Left            =   -15
      TabIndex        =   2
      Tag             =   "3102"
      Top             =   765
      Width           =   9630
      Begin VB.TextBox txtChampValeur 
         DataField       =   "ArtContenue"
         DataMember      =   "DEcmdTblAccueil"
         DataSource      =   "DEWebBase"
         Height          =   5085
         Index           =   0
         Left            =   3360
         MultiLine       =   -1  'True
         TabIndex        =   24
         Tag             =   "3112"
         Top             =   3840
         Width           =   6135
      End
      Begin MSDataListLib.DataCombo mcbChampValeur 
         Bindings        =   "Frmtableaccueil.frx":0000
         DataField       =   "CliNum"
         DataMember      =   "DEcmdTblClient"
         DataSource      =   "DEWebBase"
         Height          =   315
         Left            =   3375
         TabIndex        =   23
         Top             =   330
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "CLIENT"
         BoundColumn     =   "CLIENTNUMERO"
         Text            =   "DataCombo1"
         Object.DataMember      =   "DECmdVuelClientNumeroNom"
      End
      Begin VB.TextBox txtChampValeur 
         DataField       =   "ArtDescription"
         DataMember      =   "DEcmdTblAccueil"
         DataSource      =   "DEWebBase"
         Height          =   1365
         Index           =   1
         Left            =   3360
         MultiLine       =   -1  'True
         TabIndex        =   10
         Tag             =   "3110"
         Top             =   1800
         Width           =   6135
      End
      Begin MSMask.MaskEdBox mebChampValeur 
         Bindings        =   "Frmtableaccueil.frx":004D
         DataField       =   "ArtDateCre"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   3
         EndProperty
         DataMember      =   "DEcmdTblAccueil"
         DataSource      =   "DEWebBase"
         Height          =   285
         Index           =   0
         Left            =   3360
         TabIndex        =   8
         Tag             =   "3108"
         Top             =   840
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebChampValeur 
         Bindings        =   "Frmtableaccueil.frx":008D
         DataField       =   "ArtTitre"
         DataMember      =   "DEcmdTblAccueil"
         DataSource      =   "DEWebBase"
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   9
         Tag             =   "3109"
         Top             =   1320
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebChampValeur 
         Bindings        =   "Frmtableaccueil.frx":00CE
         DataField       =   "ArtAuteur"
         DataMember      =   "DEcmdTblAccueil"
         DataSource      =   "DEWebBase"
         Height          =   285
         Index           =   3
         Left            =   3360
         TabIndex        =   11
         Tag             =   "3111"
         Top             =   3360
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   "_"
      End
      Begin VB.Label lblChampLibelle 
         AutoSize        =   -1  'True
         Caption         =   "CONTENUE:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Tag             =   "3106"
         Top             =   3840
         Width           =   945
      End
      Begin VB.Label lblChampLibelle 
         AutoSize        =   -1  'True
         Caption         =   "AUTEUR:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Tag             =   "3105"
         Top             =   3360
         Width           =   720
      End
      Begin VB.Label lblChampLibelle 
         AutoSize        =   -1  'True
         Caption         =   "TEXT:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Tag             =   "3104"
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label lblChampLibelle 
         AutoSize        =   -1  'True
         Caption         =   "THEME:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Tag             =   "3103"
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblChampLibelle 
         AutoSize        =   -1  'True
         Caption         =   "DATE:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Tag             =   "3102"
         Top             =   840
         Width           =   480
      End
   End
   Begin VB.Frame fraTableNom 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Tag             =   "3100"
      Top             =   0
      Width           =   9615
      Begin VB.Label lblTableNom 
         Alignment       =   2  'Center
         Caption         =   "ACCUEIL"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Tag             =   "3101"
         Top             =   240
         Width           =   8535
      End
   End
   Begin MSDataGridLib.DataGrid dgdTable 
      Bindings        =   "Frmtableaccueil.frx":0118
      Height          =   2085
      Left            =   0
      TabIndex        =   22
      Tag             =   "3113"
      Top             =   10140
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3678
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
      DataMember      =   "DEcmdTblAccueil"
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "ArtNum"
         Caption         =   "ArtNum"
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
         DataField       =   "ArtCliNum"
         Caption         =   "ArtCliNum"
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
         DataField       =   "ArtArtCatNum"
         Caption         =   "ArtArtCatNum"
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
         DataField       =   "ArtDateCre"
         Caption         =   "ArtDateCre"
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
         DataField       =   "ArtDateMaj"
         Caption         =   "ArtDateMaj"
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
         DataField       =   "ArtDateAcc"
         Caption         =   "ArtDateAcc"
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
         DataField       =   "ArtAuteur"
         Caption         =   "ArtAuteur"
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
         DataField       =   "ArtCode"
         Caption         =   "ArtCode"
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
         DataField       =   "ArtTitre"
         Caption         =   "ArtTitre"
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
         DataField       =   "ArtDescription"
         Caption         =   "ArtDescription"
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
         DataField       =   "ArtContenue"
         Caption         =   "ArtContenue"
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
            ColumnWidth     =   975,118
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
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmTableAccueil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    Delta Copyright                                                             (31/10/2000)  ***
'******************************************************************************
'***    FORM:                                                                                              ***
'***        frmTableAccueil                                                                           ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***        - Pour La Gestion Des Messages D'Accueil                                    ***
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
Private Const mconFeuilleType = FEUILLEFORM                                                      ' Le type de feuille
Private Const mconFeuilleNom = "frmTableAccueil"                                       ' Le nom de la Feuille

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
    InitialiseControl FrmTableAccueil
   ' Je Desactive Les Controls De Saissie Des Champs De La Frame De Saissie
    EditeControl FrmTableAccueil, False
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
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, INSTRUCTIONEVENEMENT, "Form_Load", vbNullString, Err
    ' Je Continue
    Resume Form_Load_Exit
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'***        cmdBouton_Click(Index As Integer)                                                ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***                                                                                                            ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Index  - Numero Du Bouton Appuyer                                                ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
Private Sub cmdBouton_Click(Index As Integer)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo cmdBouton_Click_Erreur
    ' Je Valide L'Action Sur Les Boutons De Commande
    ValideFormBouton Me, DEWebBase.rsDEcmdTblAccueil, Index
    ' Fin
cmdBouton_Click_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
cmdBouton_Click_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, INSTRUCTIONEVENEMENT, "cmdBouton_Click", CStr(Index), Err
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
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, INSTRUCTIONEVENEMENT, "mebChampValeur_GotFocus", CStr(Index), Err
    ' Je Continue
    Resume mebChampValeur_GotFocus_Exit
    ' Fin
End Sub
