VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTableLiensSite 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   Tag             =   "1100"
   Begin VB.Frame fraTableNom 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Tag             =   "1100"
      Top             =   0
      Width           =   9615
      Begin VB.Label lblTableNom 
         Alignment       =   2  'Center
         Caption         =   "LIENS SITE"
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Tag             =   "1101"
         Top             =   240
         Width           =   8535
      End
   End
   Begin VB.Frame fraTableCommand 
      Height          =   1095
      Left            =   0
      TabIndex        =   13
      Tag             =   "1117"
      Top             =   7440
      Width           =   9615
      Begin VB.CommandButton cmdBouton 
         Caption         =   "ANNULER"
         Height          =   735
         Index           =   8
         Left            =   8400
         TabIndex        =   22
         Tag             =   "1126"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "|<<"
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Tag             =   "1118"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "<"
         Height          =   735
         Index           =   1
         Left            =   1080
         TabIndex        =   15
         Tag             =   "1119"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   ">"
         Height          =   735
         Index           =   2
         Left            =   2040
         TabIndex        =   16
         Tag             =   "1120"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   ">>|"
         Height          =   735
         Index           =   3
         Left            =   3000
         TabIndex        =   17
         Tag             =   "1121"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "+"
         Height          =   735
         Index           =   4
         Left            =   4080
         TabIndex        =   18
         Tag             =   "1122"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "-"
         Height          =   735
         Index           =   5
         Left            =   5040
         TabIndex        =   19
         Tag             =   "1123"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "EDITER"
         Height          =   735
         Index           =   6
         Left            =   6240
         TabIndex        =   20
         Tag             =   "1124"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBouton 
         Caption         =   "VALIDER"
         Height          =   735
         Index           =   7
         Left            =   7440
         TabIndex        =   21
         Tag             =   "1125"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraTableChamp 
      Height          =   3360
      Left            =   0
      TabIndex        =   2
      Tag             =   "1102"
      Top             =   690
      Width           =   9615
      Begin VB.TextBox txtChampValeur 
         DataField       =   "LieDescription"
         DataMember      =   "DEcmdTblLiensSite"
         DataSource      =   "DEWebBase"
         Height          =   285
         Index           =   2
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   25
         Tag             =   "1114"
         Top             =   1320
         Width           =   6105
      End
      Begin VB.TextBox txtChampValeur 
         DataField       =   "LieImageMin"
         DataMember      =   "DEcmdTblLiensSite"
         DataSource      =   "DEWebBase"
         Height          =   285
         Index           =   1
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   24
         Tag             =   "1114"
         Top             =   1680
         Width           =   6105
      End
      Begin MSDataListLib.DataList dblChampValeur 
         Bindings        =   "Frmtablelienssite.frx":0000
         DataField       =   "LieLieCatNum"
         DataMember      =   "DEcmdTblLiensSite"
         DataSource      =   "DEWebBase"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Tag             =   "1109"
         Top             =   240
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   450
         _Version        =   393216
         ListField       =   "CATEGORIENOM"
         BoundColumn     =   "LieCatNum"
         Object.DataMember      =   "DEcmdTblLiensCategorie"
      End
      Begin VB.TextBox txtChampValeur 
         DataField       =   "LieDescription"
         DataMember      =   "DEcmdTblLiensSite"
         DataSource      =   "DEWebBase"
         Height          =   1095
         Index           =   0
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   12
         Tag             =   "1114"
         Top             =   2040
         Width           =   6105
      End
      Begin MSMask.MaskEdBox mebChampValeur 
         Bindings        =   "Frmtablelienssite.frx":002A
         DataField       =   "LieSite"
         DataMember      =   "DEcmdTblLiensSite"
         DataSource      =   "DEWebBase"
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   10
         Tag             =   "1110"
         Top             =   600
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebChampValeur 
         Bindings        =   "Frmtablelienssite.frx":0066
         DataField       =   "LieDateCre"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   3
         EndProperty
         DataMember      =   "DEcmdTblLiensSite"
         DataSource      =   "DEWebBase"
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   11
         Tag             =   "1111"
         Top             =   960
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         PromptChar      =   "_"
      End
      Begin VB.Label lblChampLibelle 
         AutoSize        =   -1  'True
         Caption         =   "PHOTO:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Tag             =   "1107"
         Top             =   1680
         Width           =   615
      End
      Begin VB.Image imgPhotoImage 
         Height          =   1395
         Left            =   7680
         Stretch         =   -1  'True
         Tag             =   "115"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblChampLibelle 
         AutoSize        =   -1  'True
         Caption         =   "DATE:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Tag             =   "1105"
         Top             =   960
         Width           =   480
      End
      Begin VB.Label lblChampLibelle 
         AutoSize        =   -1  'True
         Caption         =   "CATEGORIE:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Tag             =   "1103"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblChampLibelle 
         AutoSize        =   -1  'True
         Caption         =   "SITE:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Tag             =   "1104"
         Top             =   600
         Width           =   405
      End
      Begin VB.Label lblChampLibelle 
         AutoSize        =   -1  'True
         Caption         =   "URL:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Tag             =   "1106"
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblChampLibelle 
         AutoSize        =   -1  'True
         Caption         =   "DESCRIPTION:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Tag             =   "1108"
         Top             =   2040
         Width           =   1140
      End
   End
   Begin MSDataGridLib.DataGrid dgdTable 
      Bindings        =   "Frmtablelienssite.frx":00A2
      Height          =   3330
      Left            =   15
      TabIndex        =   23
      Tag             =   "1115"
      Top             =   4065
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   5874
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
      DataMember      =   "DEcmdTblLiensSite"
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "LieNum"
         Caption         =   "LieNum"
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
         DataField       =   "LieCliNum"
         Caption         =   "LieCliNum"
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
         DataField       =   "LieDateCre"
         Caption         =   "LieDateCre"
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
         DataField       =   "LieDateMaj"
         Caption         =   "LieDateMaj"
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
         DataField       =   "LieDateAcc"
         Caption         =   "LieDateAcc"
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
         DataField       =   "LieLieCatNum"
         Caption         =   "LieLieCatNum"
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
         DataField       =   "LieSite"
         Caption         =   "LieSite"
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
         DataField       =   "LieURL"
         Caption         =   "LieURL"
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
         DataField       =   "LieDescription"
         Caption         =   "LieDescription"
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
         DataField       =   "LieImageMax"
         Caption         =   "LieImageMax"
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
         DataField       =   "LieImageMin"
         Caption         =   "LieImageMin"
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
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739,906
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
Attribute VB_Name = "frmTableLiensSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    Delta Copyright                                                             (31/10/2000)  ***
'******************************************************************************
'***    FORM:                                                                                              ***
'***        frmTableLiens                                                                               ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***     Pour Gestion De La Table Des Liens                                                 ***
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
Private Const LOGFEUILLENOM = "frmTableLiens"                                           ' Le nom de la Feuille

'******************************************************************************
'***    PROCEDURE:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
Private Sub ChangePhoto()
    Dim intListeIndex As Integer
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo ChangePhoto_Erreur
    ' Je recherche la photo actuel
    imgPhotoImage.Picture = LoadPicture(txtChampValeur(CHAMPLIENSPHOTO))
    ' Je sort de la boucle
ChangePhoto_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
ChangePhoto_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, FEUILLEFORM, LOGFEUILLENOM, INSTRUCTIONEVENEMENT, "ChangePhoto", vbNullString, Err
    ' Je Continue
    Resume ChangePhoto_Exit
    ' Fin
End Sub


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
      
    ChangePhoto
    ' Fin
Form_Load_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
Form_Load_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, FEUILLEFORM, LOGFEUILLENOM, INSTRUCTIONEVENEMENT, "Form_Load", vbNullString, Err
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
    Dim intListeIndex As Integer
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo cmdBouton_Click_Erreur
    
    ValideFormBouton Me, DEWebBase.rsDEcmdTblLiensSite, Index
    
    ChangePhoto
    ' Fin
cmdBouton_Click_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
cmdBouton_Click_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, FEUILLEFORM, LOGFEUILLENOM, INSTRUCTIONEVENEMENT, "cmdBouton_Click", CStr(Index), Err
    ' Je Continue
    Resume cmdBouton_Click_Exit
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
Private Sub mebChampValeur_Validate(Index As Integer, Cancel As Boolean)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo mebChampValeur_Change_Erreur
    
    Select Case Index
        Case CHAMPLIENSPHOTO
        
            ChangePhoto
        Case Else
        
    End Select
    ' Fin
mebChampValeur_Change_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
mebChampValeur_Change_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, FEUILLEFORM, LOGFEUILLENOM, INSTRUCTIONEVENEMENT, "mebChampValeur_Change", CStr(Index), Err
    ' Je Continue
    Resume mebChampValeur_Change_Exit
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
    gfloLogWebBase.AjouteErreur App, FEUILLEFORM, LOGFEUILLENOM, INSTRUCTIONEVENEMENT, "mebChampValeur_GotFocus", CStr(Index), Err
    ' Je Continue
    Resume mebChampValeur_GotFocus_Exit
    ' Fin
End Sub


