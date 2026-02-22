VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTableConnexion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   Tag             =   "100"
   Begin VB.Frame fraTableNom 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Tag             =   "100"
      Top             =   0
      Width           =   9615
      Begin VB.Label lblTableNom 
         Alignment       =   2  'Center
         Caption         =   "CONNEXION"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Tag             =   "101"
         Top             =   240
         Width           =   8535
      End
   End
   Begin VB.Frame fraTableCommand 
      Height          =   1095
      Left            =   0
      TabIndex        =   11
      Tag             =   "114"
      Top             =   7440
      Width           =   9615
      Begin VB.CommandButton cmdBouton 
         Caption         =   "ANNULER"
         Height          =   735
         Index           =   8
         Left            =   8520
         TabIndex        =   20
         Tag             =   "123"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "VALIDER"
         Height          =   735
         Index           =   7
         Left            =   7560
         TabIndex        =   19
         Tag             =   "122"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "EDITER"
         Height          =   735
         Index           =   6
         Left            =   6360
         TabIndex        =   18
         Tag             =   "121"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "-"
         Height          =   735
         Index           =   5
         Left            =   5160
         TabIndex        =   17
         Tag             =   "120"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "+"
         Height          =   735
         Index           =   4
         Left            =   4200
         TabIndex        =   16
         Tag             =   "119"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   ">>|"
         Height          =   735
         Index           =   3
         Left            =   3120
         TabIndex        =   15
         Tag             =   "118"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   ">"
         Height          =   735
         Index           =   2
         Left            =   2160
         TabIndex        =   14
         Tag             =   "117"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "<"
         Height          =   735
         Index           =   1
         Left            =   1200
         TabIndex        =   13
         Tag             =   "116"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "|<<"
         Height          =   735
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Tag             =   "115"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraTableChamp 
      Height          =   3960
      Left            =   0
      TabIndex        =   2
      Tag             =   "102"
      Top             =   735
      Width           =   9615
      Begin VB.TextBox txtChampValeur 
         DataField       =   "MacConInterface"
         DataMember      =   "DEcmdTblConnexion"
         DataSource      =   "DEWebBase"
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   9
         Tag             =   "110"
         Top             =   1320
         Width           =   6135
      End
      Begin MSMask.MaskEdBox mebChampValeur 
         Bindings        =   "Frmtableconnexion.frx":0000
         DataField       =   "MacConDate"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   3
         EndProperty
         DataMember      =   "DEcmdTblConnexion"
         DataSource      =   "DEWebBase"
         Height          =   285
         Index           =   0
         Left            =   3360
         TabIndex        =   7
         Tag             =   "108"
         Top             =   360
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebChampValeur 
         Bindings        =   "Frmtableconnexion.frx":0046
         DataField       =   "MacConHeure"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   4
         EndProperty
         DataMember      =   "DEcmdTblConnexion"
         DataSource      =   "DEWebBase"
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   8
         Tag             =   "109"
         Top             =   840
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebChampValeur 
         Bindings        =   "Frmtableconnexion.frx":008D
         DataField       =   "MacConIP"
         DataMember      =   "DEcmdTblConnexion"
         DataSource      =   "DEWebBase"
         Height          =   285
         Index           =   3
         Left            =   3360
         TabIndex        =   10
         Tag             =   "111"
         Top             =   1800
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   "_"
      End
      Begin VB.Label lblChampLibelle 
         AutoSize        =   -1  'True
         Caption         =   "DATE:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Tag             =   "103"
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblChampLibelle 
         AutoSize        =   -1  'True
         Caption         =   "HEURE:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Tag             =   "104"
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblChampLibelle 
         AutoSize        =   -1  'True
         Caption         =   "NAVIGATEUR:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Tag             =   "105"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblChampLibelle 
         AutoSize        =   -1  'True
         Caption         =   "REMOTE HOST:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Tag             =   "106"
         Top             =   1800
         Width           =   1230
      End
   End
   Begin MSDataGridLib.DataGrid dgdTable 
      Bindings        =   "Frmtableconnexion.frx":00D9
      Height          =   2625
      Left            =   15
      TabIndex        =   21
      Tag             =   "113"
      Top             =   4770
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   4630
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   0   'False
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
      DataMember      =   "DEcmdTblConnexion"
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "MacConNum"
         Caption         =   "MacConNum"
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
         DataField       =   "MacConDate"
         Caption         =   "MacConDate"
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
         DataField       =   "MacConHeure"
         Caption         =   "MacConHeure"
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
         DataField       =   "MacConDateFin"
         Caption         =   "MacConDateFin"
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
         DataField       =   "MacConHeureFin"
         Caption         =   "MacConHeureFin"
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
         DataField       =   "MacConIP"
         Caption         =   "MacConIP"
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
         DataField       =   "MacConSessionID"
         Caption         =   "MacConSessionID"
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
         DataField       =   "MacConInterface"
         Caption         =   "MacConInterface"
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
         DataField       =   "MacConLargeur"
         Caption         =   "MacConLargeur"
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
         DataField       =   "MacConHauteur"
         Caption         =   "MacConHauteur"
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
         DataField       =   "MacConRefere"
         Caption         =   "MacConRefere"
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
         DataField       =   "MacConCliNum"
         Caption         =   "MacConCliNum"
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
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
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
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1184,882
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1110,047
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTableConnexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    Delta Copyright                                                             (31/10/2000)  ***
'******************************************************************************
'***    FORM:                                                                                              ***
'***        frmTableConnexion                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     - Pour Gestion De La Table Des Connexions                                      ***
'***     - Affiche la liste des enregistrement                                                 ***
'***     - Affiche l'enregistrement courant et les champs correpondant          ***
'******************************************************************************
'***    PROGRAMMEUR:                                                                              ***
'***      Royer Jean-Laurent                                                                         ***
'******************************************************************************

'******************************************************************************
'***    MODIF :                                                                                            ***
'***      Version 1.0 : 30/10/2000 :                                                                ***
'***      - Creation initial de la classe                                                          ***
'******************************************************************************
Option Explicit                                                                                               ' Je doit etre sur que mes variables on ete declarer

'******************************************************************************
'***    Declaration De Constante Privee                                                       ***
'******************************************************************************

'******************************************************************************
'***    Constante Qui Defini Les Libelles De La feuille En Erreur                   ***
'******************************************************************************
Private Const mconFeuilleType = FEUILLEFORM                                                       ' Le type de feuille
Private Const mconFeuilleNom = "frmTableConnexion"                                   ' Le nom de la Feuille

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
    InitialiseControl Me
   ' Je Desactive Les Controls De Saissie Des Champs De La Frame De Saissie
    EditeControl Me, False
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
    ValideFormBouton Me, DEWebBase.rsDEcmdTblConnexion, Index
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


