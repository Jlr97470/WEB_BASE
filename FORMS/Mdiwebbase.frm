VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm mdiWebBase 
   BackColor       =   &H8000000C&
   ClientHeight    =   5280
   ClientLeft      =   375
   ClientTop       =   1050
   ClientWidth     =   9540
   LockControls    =   -1  'True
   Begin MSComDlg.CommonDialog dlgWebBase 
      Left            =   0
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbrDeltaWebBase 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4905
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17639
            MinWidth        =   17639
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrDeltaWebBase 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   1429
      _CBWidth        =   9540
      _CBHeight       =   810
      _Version        =   "6.0.8169"
      MinHeight1      =   360
      Width1          =   2880
      NewRow1         =   0   'False
      MinHeight2      =   360
      Width2          =   1440
      NewRow2         =   -1  'True
      MinHeight3      =   360
      Width3          =   1440
      NewRow3         =   0   'False
   End
   Begin VB.Menu mnuDeltaWebBase 
      Caption         =   "&Fichier"
      Index           =   0
      Begin VB.Menu mnuFichier 
         Caption         =   "&Importer Favoris Internet Explorer"
         Index           =   0
      End
      Begin VB.Menu mnuFichier 
         Caption         =   "&Exporter Favoris Internet Explorer"
         Index           =   1
      End
      Begin VB.Menu mnuFichier 
         Caption         =   "&Exportation Outlook"
         Index           =   2
      End
      Begin VB.Menu mnuFichier 
         Caption         =   "&Quitter"
         Index           =   3
      End
   End
   Begin VB.Menu mnuDeltaWebBase 
      Caption         =   "&Information"
      Index           =   1
      Begin VB.Menu mnuInformation 
         Caption         =   "&Information Connexion"
         Index           =   0
      End
      Begin VB.Menu mnuInformation 
         Caption         =   "&Information Client"
         Index           =   1
      End
      Begin VB.Menu mnuInformation 
         Caption         =   "&Information Accueil"
         Index           =   2
      End
      Begin VB.Menu mnuInformation 
         Caption         =   "&Information Liens Categorie"
         Index           =   3
      End
      Begin VB.Menu mnuInformation 
         Caption         =   "&Information Liens Site"
         Index           =   4
      End
   End
   Begin VB.Menu mnuDeltaWebBase 
      Caption         =   "&Parametre"
      Index           =   2
      Begin VB.Menu mnuParametre 
         Caption         =   "&General"
         Index           =   0
      End
   End
End
Attribute VB_Name = "mdiWebBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    Delta Copyright                                                             (31/05/2001)  ***
'******************************************************************************
'***    MDIFORM:                                                                                        ***
'***        mdiWebBase                                                                                ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
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
Private Const mconFeuilleType = FEUILLEMDIFORM                                                   ' Le type de feuille
Private Const mconFeuilleNom = "mdiWebBase"                                               ' Le nom de la Feuille

'******************************************************************************
'***    Evenement                                                                                       ***
'******************************************************************************

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
Private Sub MDIForm_Load()
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo MDIForm_Load_Erreur
   ' Cntre la feuille.
   Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
   ' Je defini le nom de l'application
'FIXIT: App.Revision property n'a pas d'équivalent Visual Basic .NET et ne peut pas être mis à niveau.     FixIT90210ae-R7593-R67265
   Me.Caption = App.ProductName & " V " & App.Major & "." & App.Minor & "." & App.Revision & " Copyright " + App.LegalCopyright
   ' Fin
MDIForm_Load_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
MDIForm_Load_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, INSTRUCTIONEVENEMENT, "MDIForm_Load", vbNullString, Err
    ' Je Continue
    Resume MDIForm_Load_Exit
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
Private Sub MDIForm_Unload(Cancel As Integer)
    Dim frmForm As Form
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo MDIForm_Unload_Erreur
    
    EcritureParametre

'FIXIT: La collection Forms n'a pas été mise à niveau vers Visual Basic .NET par l'Assistant Mise à niveau.     FixIT90210ae-R6616-H1984
    For Each frmForm In Forms
    
        Unload frmForm
    Next
    ' Fin
MDIForm_Unload_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
MDIForm_Unload_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, INSTRUCTIONEVENEMENT, "MDIForm_Unload", vbNullString, Err
    ' Je Continue
    Resume MDIForm_Unload_Exit
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
Private Sub mnuFichier_Click(Index As Integer)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo mnuFichier_Click_Erreur
        
    Select Case Index
      Case FICHIERIMPORTERFAVORIS, FICHIEREXPORTERFAVORIS
      
         mnuFichier(FICHIERIMPORTERFAVORIS).Enabled = False
                            
         mnuFichier(FICHIEREXPORTERFAVORIS).Enabled = False
                            
         frmTableLiensCategorie.fraTableCommand.Enabled = False
                            
         frmTableLiensSite.fraTableCommand.Enabled = False
      
         Select Case Index
            Case FICHIERIMPORTERFAVORIS
                              
               ImporteFavorisInternet
            Case FICHIEREXPORTERFAVORIS
                               
               ExporteFavorisInternet
         End Select
         
         frmTableLiensSite.fraTableCommand.Enabled = True
                
         frmTableLiensCategorie.fraTableCommand.Enabled = True
                
         mnuFichier(FICHIERIMPORTERFAVORIS).Enabled = True
    
         mnuFichier(FICHIEREXPORTERFAVORIS).Enabled = True
      Case FICHIEREXPORTATIONOUTLOOK
      
         frmExportationOutlook.Show
      Case FICHIERQUITTER
         
            Unload Me
    End Select
    ' Fin
mnuFichier_Click_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
mnuFichier_Click_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, INSTRUCTIONEVENEMENT, "mnuFichier_Click", CStr(Index), Err
    ' Je Continue
    Resume mnuFichier_Click_Exit
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
Private Sub mnuInformation_Click(Index As Integer)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo mnuInformation_Click_Erreur
    Select Case Index
        Case INFORMATIONCONNEXION
        
            frmTableConnexion.Show
        Case INFORMATIONCLIENT
                    
            frmTableClient.Show
        Case INFORMATIONACCUEIL
        
            FrmTableAccueil.Show
            
        Case INFORMATIONLIENSCATEGORIE
                    
            frmTableLiensCategorie.Show
            
        Case INFORMATIONLIENSSITE
                    
            frmTableLiensSite.Show
    End Select
    ' Fin
mnuInformation_Click_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
mnuInformation_Click_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, INSTRUCTIONEVENEMENT, "mnuInformation_Click", CStr(Index), Err
    ' Je Continue
    Resume mnuInformation_Click_Exit
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
Private Sub mnuParametre_Click(Index As Integer)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo mnuParametre_Click_Erreur
    
    Select Case Index
        Case PARAMETREGENERAL
        
            frmParametre.Show
    End Select
    ' Fin
mnuParametre_Click_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
mnuParametre_Click_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, INSTRUCTIONEVENEMENT, "mnuParametre_Click", CStr(Index), Err
    ' Je Continue
    Resume mnuParametre_Click_Exit
    ' Fin
End Sub
