VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmParametre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametre"
   ClientHeight    =   5880
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6150
   Icon            =   "frmParametre.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraParametre 
      Height          =   4665
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   5655
      Begin VB.TextBox txtParametre 
         Height          =   285
         Index           =   5
         Left            =   2880
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   3480
         Width           =   2415
      End
      Begin VB.TextBox txtParametre 
         Height          =   285
         Index           =   4
         Left            =   2880
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox txtParametre 
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtParametre 
         Height          =   285
         Index           =   2
         Left            =   2880
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtParametre 
         Height          =   270
         Index           =   1
         Left            =   2880
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   825
         Width           =   2415
      End
      Begin VB.TextBox txtParametre 
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblParametre 
         Caption         =   "Repertoire Des Photos"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   16
         Top             =   3480
         Width           =   2535
      End
      Begin VB.Label lblParametre 
         Caption         =   "Repertoire Des Photos Des Liens"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label lblParametre 
         Caption         =   "Nom De La Base"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label lblParametre 
         Caption         =   "Repertoire De La Base"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblParametre 
         Caption         =   "Repertoire Du site"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   945
         Width           =   2535
      End
      Begin VB.Label lblParametre 
         Caption         =   "Domaine"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdBouton 
      Caption         =   "Appliquer"
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   3
      Top             =   5415
      Width           =   1095
   End
   Begin VB.CommandButton cmdBouton 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   2
      Top             =   5415
      Width           =   1095
   End
   Begin VB.CommandButton cmdBouton 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   2490
      TabIndex        =   1
      Top             =   5415
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tbsParametre 
      Height          =   5205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9181
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "GENERAL"
            ImageVarType    =   2
         EndProperty
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
   End
End
Attribute VB_Name = "frmParametre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    Delta Copyright                                                             (31/05/2001)  ***
'******************************************************************************
'***    MODULE:                                                                                          ***
'***        modWebBase                                                                                ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***       - Pour la gestion des deplacements dans les enregistrements          ***
'***       - Pour la gestion des controls et des champs de saissie                  ***
'***       - Pour la gestion des importations et exportations de donners         ***
'******************************************************************************
'***    PROGRAMMEUR:                                                                              ***
'***      Royer Jean-Laurent                                                                         ***
'******************************************************************************

'******************************************************************************
'***    MODIF :                                                                                            ***
'***      Version 1.0 : 30/10/2000 :                                                                ***
'******************************************************************************
Option Explicit

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
Private Sub Form_Load()
    ' Centre la feuille.
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    With guwbpParametre
    
      txtParametre(0) = .strSiteDomaine
      
      txtParametre(1) = .strSiteRepertoire
      
      txtParametre(2) = .strBaseRepertoire
      
      txtParametre(3) = .strBaseFichier
      
      txtParametre(4) = .strBaseLiens
      
      txtParametre(5) = .strBasePhoto
    End With
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

   Select Case Index
      Case 0
      
         With guwbpParametre
         
            .strSiteDomaine = txtParametre(0)
         
            .strSiteRepertoire = txtParametre(1)
            
            .strBaseRepertoire = txtParametre(2)
            
            .strBaseFichier = txtParametre(3)
            
            .strBaseLiens = txtParametre(4)
            
            .strBasePhoto = txtParametre(5)
         End With
      
         EcritureParametre
         
         Unload Me
      Case 1
      
         Unload Me
      Case 2
      
         With guwbpParametre
         
            .strSiteDomaine = txtParametre(0)
         
            .strSiteRepertoire = txtParametre(1)
            
            .strBaseRepertoire = txtParametre(2)
            
            .strBaseFichier = txtParametre(3)
            
            .strBaseLiens = txtParametre(4)
            
            .strBasePhoto = txtParametre(5)
         End With
      
         EcritureParametre
   End Select
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
Private Sub txtParametre_Validate(Index As Integer, Cancel As Boolean)
    Dim fsoObjectFichier As New FileSystemObject                                          ' L'object de system de fichier

   Select Case Index
      Case 0
      
      Case 1
         Select Case fsoObjectFichier.FolderExists(txtParametre(1))
            Case True
            
               Cancel = False
            Case False
                           
               MsgBox "Repertoire Non Valide"
               
               txtParametre(Index).SetFocus
               
               Cancel = True
         End Select
      Case 2
         Select Case fsoObjectFichier.FolderExists(txtParametre(1) + "\" + txtParametre(2))
            Case True
            
               Cancel = False
            Case False
                           
               MsgBox "Repertoire Non Valide"
               
               txtParametre(Index).SetFocus
               
               Cancel = True
         End Select
      Case 3
         Select Case fsoObjectFichier.FileExists(txtParametre(1) + "\" + txtParametre(2) + "\" + txtParametre(3))
            Case True
            
               Cancel = False
            Case False
                           
               MsgBox "Fichier Non Valide"
               
               txtParametre(Index).SetFocus
               
               Cancel = True
         End Select
      Case 4
         Select Case fsoObjectFichier.FolderExists(txtParametre(1) + "\" + txtParametre(4))
            Case True
            
               Cancel = False
            Case False
                           
               MsgBox "Repertoire Non Valide"
               
               txtParametre(Index).SetFocus
                           
               Cancel = True
         End Select
      End Select
End Sub
