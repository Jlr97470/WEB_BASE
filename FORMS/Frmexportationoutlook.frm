VERSION 5.00
Begin VB.Form frmExportationOutlook 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmExportationOutlook"
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

Private WithEvents OutlookApp As Outlook.Application
Attribute OutlookApp.VB_VarHelpID = -1

'******************************************************************************
'***    Declaration De Constante Privee                                                       ***
'******************************************************************************

'******************************************************************************
'***    Constante Qui Defini Les Libelles De La feuille En Erreur                   ***
'******************************************************************************
Private Const mconFeuilleType = FEUILLEFORM                                                     ' Le type de feuille
Private Const mconFeuilleNom = "frmExportationOutLook"                             ' Le nom de la Feuille

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
'FIXIT: Déclarer 'myItem' avec un type de données à liaison anticipée                      FixIT90210ae-R1672-R1B8ZE
   Dim myItem As Object
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo Form_Load_Erreur
   ' Je Centre la feuille.
   Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
   
   Set OutlookApp = CreateObject("Outlook.Application")
   
   Set myItem = OutlookApp.CreateItem(olContactItem)
   
   myItem.Display
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

