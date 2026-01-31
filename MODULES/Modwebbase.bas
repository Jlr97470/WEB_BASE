Attribute VB_Name = "modWebBase"
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
Option Explicit                                                                                               ' Je doit etre sur que mes variables on ete declarer

'******************************************************************************
'***    Declaration De Constante Privee                                                       ***
'******************************************************************************

'******************************************************************************
'***    Constante Qui Defini Les Libelles De La feuille En Erreur                   ***
'******************************************************************************
Private Const mconFeuilleType = FEUILLEMODULE                                                 ' Le type de feuille
Private Const mconFeuilleNom = "modWebBase"                                             ' Le nom de la Feuille

'******************************************************************************
'***    Declaration De Constante Public                                                       ***
'******************************************************************************

'******************************************************************************
'***    Constante Qui Defini Les Libelles Des Erreur                                     ***
'******************************************************************************
Public Const LIBELLEFONCTION = "FONCTION"                                                ' Le libelle fonction
Public Const LIBELLEPROCEDURE = "PROCEDURE"                                         ' Le libelle procedure

'******************************************************************************
'***    Constante Qui defini les numeros des boutons de command               ***
'******************************************************************************
Public Enum BOUTONCOMMANDCONSTANTES
    COMMANDPREMIER = 0
    COMMANDPRECEDENT = 1
    COMMANDSUIVANT = 2
    COMMANDDERNIER = 3
    COMMANDAJOUTER = 4
    COMMANDSUPPRIMER = 5
    COMMANDEDITER = 6
    COMMANDVALIDER = 7
    COMMANDANNULER = 8
End Enum

'******************************************************************************
'***    Constante Qui defini les boutons valides                                            ***
'******************************************************************************
Public Enum BOUTONVALIDECONSTANTES
    VALIDEOUI = 0
    VALIDENON = 1
    VALIDEDEPLACEMENT = 2
End Enum

'******************************************************************************
'***    Constante Qui defini les numeros du menu parametre                        ***
'******************************************************************************
Public Enum TABLELIENSCHAMPCONSTANTES
    CHAMPLIENSPHOTO = 1
End Enum

'******************************************************************************
'***    Constante Qui defini les numeros du menu fichier                              ***
'******************************************************************************
Public Enum MENUFICHIERCONSTANTES
    FICHIERIMPORTERFAVORIS = 0
    FICHIEREXPORTERFAVORIS = 1
    FICHIEREXPORTATIONOUTLOOK = 2
    FICHIERQUITTER = 3
End Enum

'******************************************************************************
'***    Constante Qui defini les numeros du menu information                       ***
'******************************************************************************
Public Enum MENUINFORMATIONCONSTANTES
    INFORMATIONCONNEXION = 0
    INFORMATIONCLIENT = 1
    INFORMATIONACCUEIL = 2
    INFORMATIONLIENSCATEGORIE = 3
    INFORMATIONLIENSSITE = 4
End Enum

'******************************************************************************
'***    Constante Qui defini les numeros du menu parametre                        ***
'******************************************************************************
Public Enum MENUPARAMETRECONSTANTES
    PARAMETREGENERAL = 0
End Enum

'******************************************************************************
'***    Declaration De Type Public                                                               ***
'******************************************************************************
Public Type WEBBASEPARAMETRE
    strSiteDomaine As String
    strSiteRepertoire As String                                                                         ' Le repertoire du site
    strBaseRepertoire  As String                                                                       ' Le repertoire de la base
    strBaseFichier As String                                                                             ' Le fichier de la base
    strBaseLiens As String
    strBasePhoto As String
    strFavorisRepertoire As String
End Type

'******************************************************************************
'***    Declaration De Variable Priver                                                          ***
'******************************************************************************

'******************************************************************************
'***    Declaration De Object Public                                                             ***
'******************************************************************************

'******************************************************************************
'***    Object Pour La Gestion Des Parametres                                            ***
'******************************************************************************
Public guwbpParametre As WEBBASEPARAMETRE

'******************************************************************************
'***    Object Pour La Gestion D'Un Fichier Profiler                                      ***
'******************************************************************************
Public gfprFicProfiler As New ClsFicPrf

'******************************************************************************
'***    Object Pour La Gestion D'un Fichier Journal                                      ***
'******************************************************************************
Public gfloLogWebBase As New ClsFicLog                                                           ' L'object pour la gestion d'un fichier journal

'******************************************************************************
'***    Declaration De Procedure Priver                                                        ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'***        Main()                                                                                           ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Private Sub Main()
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo Main_Erreur
    ' Je creer l'object journal
    Set gfloLogWebBase = New ClsFicLog
    ' J'initialise Les Parametres
    InitialiseParametre
    ' Je Lit Les Parametres Du Fichier Ini
    LectureParametre
    ' Je vais beaucoup utiliser
    With guwbpParametre
      ' Je Defini La Chaine De Connexion A La Base De Donnee
      DEWebBase.DEconWebBase.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & .strSiteRepertoire & "\" & .strBaseRepertoire & "\" & .strBaseFichier & ";Persist Security Info=False"
    End With
    ' J'Ouvre La Feuille Principale
    mdiWebBase.Show
    ' Fin
Main_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
Main_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, LIBELLEPROCEDURE, "Main", vbNullString, Err
    ' Je Continue
    Resume Main_Exit
    ' Fin
End Sub

'******************************************************************************
'***    Declaration De Procedure Public                                                       ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'***        InitialiseControl(ByRef frmInitialise As Form)                                ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      - Initialiser Les Controls De La Feuille A Partir Du Tag Du Control    ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***        FrmInitialise - Form a initialiser                                                    ***
'***    SORTIE:                                                                                           ***
'***        Neant                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***        InitialiseControl(FrmName)                                                           ***
'******************************************************************************
Public Sub InitialiseControl(ByRef frmInitialise As Form)
    Dim ctlControl As Control
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo InitialiseControl_Erreur
    ' J'initialise le nom de la feuille
    frmInitialise.Caption = LoadResString(frmInitialise.Tag)
    ' Je fait defiler les controls de la feuilles
    For Each ctlControl In frmInitialise.Controls
        ' Je regarde Le Type De Control
        Select Case TypeName(ctlControl)
            Case "Label"
                ' C'est control label
                ' Je defini sont texte
'FIXIT: 'Caption' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'Caption', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                ctlControl.Caption = LoadResString(ctlControl.Tag)
            Case "Frame"
               ' C'est une frame
               ' Je defini sont texte
            Case "CommandBouton"
               ' C'est un bouton de commande
               ' Je defini sont texte
'FIXIT: 'Caption' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'Caption', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                ctlControl.Caption = LoadResString(ctlControl.Tag)
            Case "TextBox"
               ' C'est un TextBox
            Case "MaskEdBox"
                'C'est un TextBox
                ' Je defini si le caractere inviter est inclut dans le champ texte
'FIXIT: 'PromptInclude' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'PromptInclude', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                ctlControl.PromptInclude = CBool(LoadResString(ctlControl.Tag + 300))
                ' Je defini le caractere d'invite
'FIXIT: 'PromptChar' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'PromptChar', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                ctlControl.PromptChar = LoadResString(ctlControl.Tag + 200)
                'Je defini sont mask
'FIXIT: 'Mask' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'Mask', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                ctlControl.Mask = LoadResString(ctlControl.Tag + 100)
                ' Je defini le format
'FIXIT: 'DataFormat.Format' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataFormat.Format', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
'FIXIT: 'Format' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'Format', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                ctlControl.Format = ctlControl.DataFormat.Format
            Case "DataCombo"
               ' C'est un comboBox
            Case "DataList"
               ' C'est un datalist
            Case "DataGrid"
              ' C'est une datagrid
        End Select
    Next
    ' Fin
InitialiseControl_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
InitialiseControl_Erreur:
    ' Je l'ecrit dans le journal
'FIXIT: 'Name)' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'Name)', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, INSTRUCTIONPROCEDURE, "InitialiseControl", CStr(ctlControl.Name), Err
    ' Je Continue
    Resume InitialiseControl_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'***        DesactiveControlData(ByRef frmDesactive As Form)                      ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      - Desactive Les Controls Lies A Une Source De Donnee                   ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***        frmDesactive - Form Ou On Veut Desactiver Les Controls Lies       ***
'***    SORTIE:                                                                                           ***
'***        Neant                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***        DesactiveControlData(FrmName)                                                  ***
'******************************************************************************
Public Sub DesactiveControlData(ByRef frmDesactive As Form)
    Dim ctlControl As Control
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo DesactiveControlData_Erreur
    ' Je fait defiler les controls de la feuilles
    For Each ctlControl In frmDesactive.Controls
        ' Je regarde Le Type De Control
        Select Case TypeName(ctlControl)
            Case "Label"
                ' C'est control label
            Case "Frame"
               ' C'est une frame
            Case "CommandBouton"
               ' C'est un bouton de commande
            Case "TextBox"
               ' C'est un TextBox
               ' Je Regarde Si Une Menbre De Source De Donnee A Ete Defini
'FIXIT: 'DataMember' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataMember', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
               Select Case ctlControl.DataMember <> vbNullString
                  Case True
                     ' Un Menbre De Source De Donne A Ete Defini
                     ' Je Desactive La Source De Donne
'FIXIT: 'DataSource' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataSource', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                     Set ctlControl.DataSource = Nothing
                  Case False
                     ' Aucun Menbre De Source De Donnee A Ete Defini
               End Select
            Case "MaskEdBox"
                'C'est un TextBox
               ' Je Regarde Si Une Menbre De Source De Donnee A Ete Defini
'FIXIT: 'DataMember' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataMember', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
               Select Case ctlControl.DataMember <> vbNullString
                  Case True
                     ' Un Menbre De Source De Donne A Ete Defini
                     ' Je Desactive La Source De Donne
'FIXIT: 'DataSource' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataSource', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                     Set ctlControl.DataSource = Nothing
                  Case False
                     ' Aucun Menbre De Source De Donnee A Ete Defini
               End Select
            Case "DataCombo"
               ' C'est un comboBox
               ' Je Regarde Si Une Menbre De Source De Donnee A Ete Defini
'FIXIT: 'DataMember' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataMember', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
               Select Case ctlControl.DataMember <> vbNullString
                  Case True
                     ' Un Menbre De Source De Donne A Ete Defini
                     ' Je Desactive La Source De Donne
'FIXIT: 'DataSource' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataSource', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                     Set ctlControl.DataSource = Nothing
                  Case False
                     ' Aucun Menbre De Source De Donnee A Ete Defini
               End Select
            Case "DataList"
               ' C'est un datalist
               ' Je Regarde Si Une Menbre De Source De Donnee A Ete Defini
'FIXIT: 'DataMember' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataMember', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
               Select Case ctlControl.DataMember <> vbNullString
                  Case True
                     ' Un Menbre De Source De Donne A Ete Defini
                     ' Je Desactive La Source De Donne
'FIXIT: 'DataSource' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataSource', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                     Set ctlControl.DataSource = Nothing
                  Case False
                     ' Aucun Menbre De Source De Donnee A Ete Defini
               End Select
            Case "DataGrid"
              ' C'est une datagrid
               ' Je Regarde Si Une Menbre De Source De Donnee A Ete Defini
'FIXIT: 'DataMember' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataMember', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
               Select Case ctlControl.DataMember <> vbNullString
                  Case True
                     ' Un Menbre De Source De Donne A Ete Defini
                     ' Je Desactive La Source De Donne
'FIXIT: 'DataSource' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataSource', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                     Set ctlControl.DataSource = Nothing
                  Case False
                     ' Aucun Menbre De Source De Donnee A Ete Defini
               End Select
        End Select
    Next
    ' Fin
DesactiveControlData_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
DesactiveControlData_Erreur:
    ' Je l'ecrit dans le journal
'FIXIT: 'Name)' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'Name)', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, LIBELLEPROCEDURE, "DesactiveControlData", CStr(ctlControl.Name), Err
    ' Je Continue
    Resume DesactiveControlData_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'***        ReactiveControlData(ByRef frmReactive As Form)                         ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      - Desactive Les Controls Lies A Une Source De Donnee                   ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***        frmDesactive - Form Ou On Veut Reactiver Les Controls Lies        ***
'***    SORTIE:                                                                                           ***
'***        Neant                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***        DesactiveControlData(FrmName)                                                  ***
'******************************************************************************
Public Sub ReactiveControlData(ByRef frmReactive As Form)
    Dim ctlControl As Control
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo ReactiveControlData_Erreur
    ' Je fait defiler les controls de la feuilles
    For Each ctlControl In frmReactive.Controls
        ' Je regarde Le Type De Control
        Select Case TypeName(ctlControl)
            Case "Label"
                ' C'est control label
            Case "Frame"
               ' C'est une frame
            Case "CommandBouton"
               ' C'est un bouton de commande
            Case "TextBox"
               ' C'est un TextBox
               ' Je Regarde Si Une Menbre De Source De Donnee A Ete Defini
'FIXIT: 'DataMember' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataMember', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
               Select Case ctlControl.DataMember <> vbNullString
                  Case True
                     ' Un Menbre De Source De Donne A Ete Defini
                     ' Je Desactive La Source De Donne
'FIXIT: 'DataSource' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataSource', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                     Set ctlControl.DataSource = DEWebBase
                  Case False
                     ' Aucun Menbre De Source De Donnee A Ete Defini
               End Select
            Case "MaskEdBox"
                'C'est un TextBox
               ' Je Regarde Si Une Menbre De Source De Donnee A Ete Defini
'FIXIT: 'DataMember' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataMember', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
               Select Case ctlControl.DataMember <> vbNullString
                  Case True
                     ' Un Menbre De Source De Donne A Ete Defini
                     ' Je Desactive La Source De Donne
'FIXIT: 'DataSource' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataSource', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                     Set ctlControl.DataSource = DEWebBase
                  Case False
                     ' Aucun Menbre De Source De Donnee A Ete Defini
               End Select
            Case "DataCombo"
               ' C'est un comboBox
               ' Je Regarde Si Une Menbre De Source De Donnee A Ete Defini
'FIXIT: 'DataMember' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataMember', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
               Select Case ctlControl.DataMember <> vbNullString
                  Case True
                     ' Un Menbre De Source De Donne A Ete Defini
                     ' Je Desactive La Source De Donne
'FIXIT: 'DataSource' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataSource', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                     Set ctlControl.DataSource = DEWebBase
                  Case False
                     ' Aucun Menbre De Source De Donnee A Ete Defini
               End Select
            Case "DataList"
               ' C'est un datalist
               ' Je Regarde Si Une Menbre De Source De Donnee A Ete Defini
'FIXIT: 'DataMember' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataMember', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
               Select Case ctlControl.DataMember <> vbNullString
                  Case True
                     ' Un Menbre De Source De Donne A Ete Defini
                     ' Je Desactive La Source De Donne
'FIXIT: 'DataSource' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataSource', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                     Set ctlControl.DataSource = DEWebBase
                  Case False
                     ' Aucun Menbre De Source De Donnee A Ete Defini
               End Select
            Case "DataGrid"
              ' C'est une datagrid
               ' Je Regarde Si Une Menbre De Source De Donnee A Ete Defini
'FIXIT: 'DataMember' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataMember', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
               Select Case ctlControl.DataMember <> vbNullString
                  Case True
                     ' Un Menbre De Source De Donne A Ete Defini
                     ' Je Desactive La Source De Donne
'FIXIT: 'DataSource' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataSource', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                     Set ctlControl.DataSource = DEWebBase
                  Case False
                     ' Aucun Menbre De Source De Donnee A Ete Defini
               End Select
        End Select
    Next
    ' Fin
ReactiveControlData_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
ReactiveControlData_Erreur:
    ' Je l'ecrit dans le journal
'FIXIT: 'Name)' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'Name)', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, LIBELLEPROCEDURE, "ReactiveControlData", CStr(ctlControl.Name), Err
    ' Je Continue
    Resume ReactiveControlData_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'***        ()                                                                                                 ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Sub InitialiseParametre()
    Dim rbwRepertoireFavoris As New RepBurWin                                             ' Un object des repertoires du bureau
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo InitialiseParametre_Erreur
    ' Je vais beaucoup utiliser
    With guwbpParametre
        
        .strSiteDomaine = ""
        
        .strSiteRepertoire = App.Path
        
        .strBaseRepertoire = "DATABASE\BASE"
        
        .strBaseFichier = "Delta-InfoFich_2010.accdb"

        .strBaseLiens = "IMAGES\LIENS"
        
        .strBasePhoto = "PHOTO"
        ' Je recupere le repertoire favoris
        rbwRepertoireFavoris.LectureRepertoire CSIDL_FAVORITES, .strFavorisRepertoire
    End With
    ' Fin
InitialiseParametre_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
InitialiseParametre_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, LIBELLEPROCEDURE, "InitialiseParametre", vbNullString, Err
    ' Je Continue
    Resume InitialiseParametre_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'***        ()                                                                                                 ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Sub LectureParametre()
    Dim fsoObjectFichier As New FileSystemObject                                          ' L'object de system de fichier
    Dim strSiteDomaine As String
    Dim strSiteRepertoire As String                                                                  ' Le repertoire du site
    Dim strBaseRepertoire  As String                                                                ' Le repertoire de la base
    Dim strBaseFichier As String                                                                      ' Le fichier de la base
    Dim strBaseLiens As String
    Dim strBasePhoto As String
    Dim strFavorisRepertoire As String
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo LectureParametre_Erreur
    ' Je vais beaucoup utiliser
    With guwbpParametre

        gfprFicProfiler.LectureValeur App.Path & "\" & App.EXEName & ".INI", "SITE", "DOMAINE", strSiteDomaine, vbNullString
           
        Select Case strSiteDomaine
            Case vbNullString
                    
            Case Else
            
                .strSiteDomaine = strSiteDomaine
        End Select
    
        gfprFicProfiler.LectureValeur App.Path & "\" & App.EXEName & ".INI", "SITE", "REPERTOIRE", strSiteRepertoire, vbNullString
           
        Select Case strSiteRepertoire
            Case vbNullString
                    
            Case Else
            
               Select Case fsoObjectFichier.FolderExists(strSiteRepertoire)
                  Case True
                  
                      .strSiteRepertoire = strSiteRepertoire
                  Case False
                  
               End Select
        End Select
           
        gfprFicProfiler.LectureValeur App.Path & "\" & App.EXEName & ".INI", "BASE", "REPERTOIRE", strBaseRepertoire, vbNullString
        
        Select Case strBaseRepertoire
            Case vbNullString
            
            Case Else
            
               Select Case fsoObjectFichier.FolderExists(strSiteRepertoire & "\" & strBaseRepertoire)
                  Case True
                  
                      .strBaseRepertoire = strBaseRepertoire
                  Case False
                  
               End Select
        End Select
            
        gfprFicProfiler.LectureValeur App.Path & "\" & App.EXEName & ".INI", "BASE", "FICHIER", strBaseFichier, vbNullString
            
         Select Case strBaseFichier
            Case vbNullString
                                    
            Case Else
            
               Select Case fsoObjectFichier.FileExists(strSiteRepertoire & "\" & strBaseRepertoire & "\" & strBaseFichier)
                  Case True
                  
                      .strBaseFichier = strBaseFichier
                  Case False
                  
               End Select
        End Select
        
        gfprFicProfiler.LectureValeur App.Path & "\" & App.EXEName & ".INI", "BASE", "LIENS", strBaseLiens, vbNullString
            
         Select Case strBaseLiens
            Case vbNullString
                                    
            Case Else
               Select Case fsoObjectFichier.FolderExists(strSiteRepertoire & "\" & strBaseLiens)
                  Case True
                  
                    .strBaseLiens = strBaseLiens
                Case False
                  
               End Select
        End Select
        
         gfprFicProfiler.LectureValeur App.Path & "\" & App.EXEName & ".INI", "BASE", "PHOTO", strBasePhoto, vbNullString
            
         Select Case strBasePhoto
            Case vbNullString
                                    
            Case Else
               Select Case fsoObjectFichier.FolderExists(strSiteRepertoire & "\" & strBasePhoto)
                  Case True
                  
                    .strBasePhoto = strBasePhoto
                Case False
                  
               End Select
        End Select
       
        gfprFicProfiler.LectureValeur App.Path & "\" & App.EXEName & ".INI", "FAVORIS", "REPERTOIRE", strFavorisRepertoire, vbNullString
            
         Select Case strFavorisRepertoire
            Case vbNullString
                                    
            Case Else
            
               Select Case fsoObjectFichier.FolderExists(strFavorisRepertoire)
                  Case True
                  
                      .strFavorisRepertoire = strFavorisRepertoire
                  Case False
                  
               End Select
        End Select
    End With
    ' Fin
LectureParametre_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
LectureParametre_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, LIBELLEPROCEDURE, "LectureParametre", vbNullString, Err
    ' Je Continue
    Resume LectureParametre_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'***        ()                                                                                                 ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Sub EcritureParametre()
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo EcritureParametre_Erreur
    ' Je vais beaucoup utiliser
    With guwbpParametre

        gfprFicProfiler.EcritureValeur App.Path & "\" & App.EXEName & ".INI", "SITE", "DOMAINE", .strSiteDomaine
               
        gfprFicProfiler.EcritureValeur App.Path & "\" & App.EXEName & ".INI", "SITE", "REPERTOIRE", .strSiteRepertoire
                      
        gfprFicProfiler.EcritureValeur App.Path & "\" & App.EXEName & ".INI", "BASE", "REPERTOIRE", .strBaseRepertoire
                    
        gfprFicProfiler.EcritureValeur App.Path & "\" & App.EXEName & ".INI", "BASE", "FICHIER", .strBaseFichier
                    
        gfprFicProfiler.EcritureValeur App.Path & "\" & App.EXEName & ".INI", "BASE", "LIENS", .strBaseLiens
                    
        gfprFicProfiler.EcritureValeur App.Path & "\" & App.EXEName & ".INI", "BASE", "PHOTO", .strBasePhoto
                    
        gfprFicProfiler.EcritureValeur App.Path & "\" & App.EXEName & ".INI", "FAVORIS", "REPERTOIRE", .strFavorisRepertoire
            
    End With
    ' Fin
EcritureParametre_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
EcritureParametre_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, LIBELLEPROCEDURE, "EcritureParametre", vbNullString, Err
    ' Je Continue
    Resume EcritureParametre_Exit
    ' Fin
End Sub


'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'***        EditeControl(ByRef frmEdite As Form, ByVal blnEdite As Boolean)***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***        - Modifier l'etat des controls d'une feuille                                      ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***        frmEdite - La form dont on veut changer l'etat des controls            ***
'***        blnEdit   - L'etat dans le quelle on veut mettre les controls            ***
'***    SORTIE:                                                                                           ***
'***        Neant                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***        EditControl(FrmName,True)                                                           ***
'******************************************************************************
Public Sub EditeControl(ByRef frmEdite As Form, ByVal blnEdite As Boolean)
    Dim ctlControl As Control
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo EditeControl_Erreur
    ' Je fait defiler les controls
    For Each ctlControl In frmEdite.Controls
      ' Je regarde sur quel object est contenu le control
'FIXIT: 'Container.Name' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'Container.Name', déclarez 'ctlControl' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
        Select Case ctlControl.Container.Name
            Case "fraTableChamp"
               ' C'est un control de champ de la frame de saissie
               ' Je change sont etat
               ctlControl.Enabled = blnEdite
            Case Else
               ' C'est un control qui n'est pas contenu dans la frame de saissie
        End Select
    Next
    ' Fin
EditeControl_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
EditeControl_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, LIBELLEPROCEDURE, "EditeControl", CStr(ctlControl.Name) & "," & CStr(blnEdite), Err
    ' Je Continue
    Resume EditeControl_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Sub ValideFormBouton(ByRef frmBouton As Form, ByRef DErstCmd As Recordset, ByVal intIndex As Integer)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo ValideFormBouton_Erreur
        
    Select Case (DErstCmd.BOF = True Or DErstCmd.EOF = True)
        Case True
        
        Case False
        
            Select Case intIndex
                Case COMMANDEDITER
                
                    ValideBouton frmBouton.cmdBouton, VALIDEOUI
                    
                    EditeControl frmBouton, True
                    
                    frmBouton.dgdTable.Enabled = False
                Case COMMANDVALIDER
                
                    DErstCmd.Update
                    
                    ValideBouton frmBouton.cmdBouton, VALIDEDEPLACEMENT
                    
                    EditeControl frmBouton, False
                    
                    frmBouton.dgdTable.Enabled = True
                    
                    ReactiveControlData frmBouton
                Case COMMANDANNULER
                    Select Case DErstCmd.EditMode
                        Case adEditNone
                        
                        Case adEditInProgress, adEditAdd
                        
                           DErstCmd.CancelUpdate
                    End Select
                    
                    ValideBouton frmBouton.cmdBouton, VALIDEDEPLACEMENT
                    
                    EditeControl frmBouton, False
                    
                    frmBouton.dgdTable.Enabled = True
                Case COMMANDSUPPRIMER
                                                                            
                    DErstCmd.Delete
            End Select
    End Select
    Select Case DErstCmd.BOF
        Case True
        
        Case False
            Select Case intIndex
                Case COMMANDPREMIER
                
                    DErstCmd.MoveFirst
                Case COMMANDPRECEDENT
                
                    Select Case DErstCmd.AbsolutePosition
                        Case Is > 1
                        
                            DErstCmd.MovePrevious
                            
                        Case Else
                        
                    End Select
            End Select
    End Select
    Select Case DErstCmd.EOF
        Case True
        
        Case False
        
            Select Case intIndex
                Case COMMANDDERNIER
                
                    DErstCmd.MoveLast
                Case COMMANDSUIVANT
                
                    Select Case DErstCmd.AbsolutePosition
                        Case Is < DErstCmd.RecordCount
                        
                            DErstCmd.MoveNext
                        Case Else
                        
                    End Select
            End Select
    End Select
    Select Case intIndex
        Case COMMANDAJOUTER
                
            DErstCmd.AddNew
            
            AjouteChampValeur frmBouton.mebChampValeur
            
            ValideBouton frmBouton.cmdBouton, VALIDEOUI
            
            EditeControl frmBouton, True
            
            frmBouton.dgdTable.Enabled = False
    End Select
    ' Fin
ValideFormBouton_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
ValideFormBouton_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, LIBELLEPROCEDURE, "ValideFormBouton", CStr(frmBouton.Name) & "," & CStr(DErstCmd.DataMember) & "," & CStr(intIndex), Err
    ' Je Continue
    Resume ValideFormBouton_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
'FIXIT: Déclarer 'colcmdBouton' avec un type de données à liaison anticipée                FixIT90210ae-R1672-R1B8ZE
Public Sub ValideBouton(ByRef colcmdBouton As Object, ByVal bvaValidation As BOUTONVALIDECONSTANTES)
    Dim cmdBouton As CommandButton
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo ValideBouton_Erreur
    
    For Each cmdBouton In colcmdBouton
    
'FIXIT: Select Case cmdBouton.Index property n'a pas d'équivalent Visual Basic .NET et ne peut pas être mis à niveau.     FixIT90210ae-R7593-R67265
        Select Case cmdBouton.Index
            Case COMMANDPREMIER, COMMANDPRECEDENT, COMMANDSUIVANT, COMMANDDERNIER, COMMANDAJOUTER, COMMANDSUPPRIMER, COMMANDEDITER
            
                Select Case bvaValidation
                    Case VALIDENON, VALIDEOUI
                    
                        cmdBouton.Enabled = False
                    Case VALIDEDEPLACEMENT
                    
                        cmdBouton.Enabled = True
                End Select
            Case COMMANDVALIDER, COMMANDANNULER
            
                Select Case bvaValidation
                    Case VALIDENON, VALIDEDEPLACEMENT
                    
                        cmdBouton.Enabled = False
                    Case VALIDEOUI
                    
                        cmdBouton.Enabled = True
                End Select
        End Select
    Next
    ' Fin
ValideBouton_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
ValideBouton_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, LIBELLEPROCEDURE, "ValideBouton", CStr(cmdBouton.Name) & "," & CStr(bvaValidation), Err
    ' Je Continue
    Resume ValideBouton_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
'FIXIT: Déclarer 'colctlChampValeur' avec un type de données à liaison anticipée           FixIT90210ae-R1672-R1B8ZE
Public Sub AjouteChampValeur(ByRef colctlChampValeur As Object)
    Dim ctlValeur As Control
    Dim DErstCmd As Recordset
    Dim DEfldValeur As Field
    Dim strValeur As String
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo AjouteChampValeur_Erreur
    
    For Each ctlValeur In colctlChampValeur
    
'FIXIT: 'DataMember' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataMember', déclarez 'ctlValeur' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
        Set DErstCmd = DEWebBase.Recordsets(ctlValeur.DataMember)
        
'FIXIT: 'DataField' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'DataField', déclarez 'ctlValeur' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
        Set DEfldValeur = DErstCmd(ctlValeur.DataField)
        
        Select Case DEfldValeur.Type
            Case adDBDate, adDate, adDBTimeStamp
'FIXIT: 'Format' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'Format', déclarez 'ctlValeur' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                Select Case ctlValeur.Format
                    Case "hh:mm:ss", "HH:mm:ss"
                    
'FIXIT: 'Format' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'Format', déclarez 'ctlValeur' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                        DEfldValeur = Format(Time, ctlValeur.Format)
                    Case "dd/mm/yyyy", "dd/MM/yyyy"
                    
'FIXIT: 'Format' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'Format', déclarez 'ctlValeur' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                        DEfldValeur = Format(Date, ctlValeur.Format)
                End Select
            Case adVarWChar
            
'FIXIT: 'Format' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'Format', déclarez 'ctlValeur' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                Select Case ctlValeur.Format
                    Case vbNullString
                    
'FIXIT: 'Mask' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'Mask', déclarez 'ctlValeur' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                        Select Case ctlValeur.Mask
                            Case vbNullString
                            
                                DEfldValeur = vbNullString
                            Case Else
                            
'FIXIT: 'Mask' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'Mask', déclarez 'ctlValeur' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                                strValeur = ctlValeur.Mask
                                
'FIXIT: 'PromptInclude' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'PromptInclude', déclarez 'ctlValeur' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                                Select Case ctlValeur.PromptInclude
                                    Case True
                                    
                                        strValeur = Replace(strValeur, "#", "0")
                                        strValeur = Replace(strValeur, "9", "0")
'FIXIT: 'PromptChar' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'PromptChar', déclarez 'ctlValeur' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                                        strValeur = Replace(strValeur, "&", ctlValeur.PromptChar)
'FIXIT: 'PromptChar' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'PromptChar', déclarez 'ctlValeur' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                                        strValeur = Replace(strValeur, "?", ctlValeur.PromptChar)
'FIXIT: 'PromptChar' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'PromptChar', déclarez 'ctlValeur' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                                        strValeur = Replace(strValeur, "C", ctlValeur.PromptChar)
                                        strValeur = Replace(strValeur, "<", vbNullString)
                                        strValeur = Replace(strValeur, ">", vbNullString)
                                    Case False
                                    
                                        strValeur = Replace(strValeur, "#", "0")
                                        strValeur = Replace(strValeur, "9", "0")
                                        strValeur = Replace(strValeur, "&", vbNullString)
                                        strValeur = Replace(strValeur, "?", vbNullString)
                                        strValeur = Replace(strValeur, "C", vbNullString)
                                        strValeur = Replace(strValeur, "<", vbNullString)
                                        strValeur = Replace(strValeur, ">", vbNullString)
                                End Select
                                
                                DEfldValeur = strValeur
                        End Select
                    Case Else
                    
'FIXIT: 'Format' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'Format', déclarez 'ctlValeur' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
                        DEfldValeur = Format(vbNullString, ctlValeur.Format)
                End Select
        End Select
    Next
    ' Fin
AjouteChampValeur_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
AjouteChampValeur_Erreur:
    ' Je l'ecrit dans le journal
'FIXIT: 'Name)' n'est pas une propriété de l'objet 'Control' générique dans Visual Basic .NET. Pour accéder à 'Name)', déclarez 'ctlValeur' en utilisant son type effectif au lieu de 'Control'     FixIT90210ae-R1460-RCFE85
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, LIBELLEPROCEDURE, "AjouteChampValeur", CStr(ctlValeur.Name), Err
    ' Je Continue
    Resume AjouteChampValeur_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Sub ImporteFavorisInternet()
    Dim fsoObjectFichier As New FileSystemObject                                          ' L'object de system de fichier
    Dim fldRepertoireFavoris As Folder                                                             ' Un object repertoire
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo ImporteFavorisInternet_Erreur
        
    mdiWebBase.sbrDeltaWebBase.SimpleText = "EXECUTE:=ImporteFavoris:Debut"
    
    Set fldRepertoireFavoris = fsoObjectFichier.GetFolder(guwbpParametre.strFavorisRepertoire)

    ImporteFavorisRepertoire fldRepertoireFavoris
    
    With DEWebBase
        
        .rsDEcmdTblLiensSite.Filter = vbNullString
                        
        .rsDEcmdTblLiensCategorie.Filter = vbNullString
    End With
    
    mdiWebBase.sbrDeltaWebBase.SimpleText = "EXECUTE:=ImporteFavoris:OK"
    ' Fin
ImporteFavorisInternet_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
ImporteFavorisInternet_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, LIBELLEPROCEDURE, "ImporteFavorisInternet", vbNullString, Err
    ' Je Continue
    Resume ImporteFavorisInternet_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Private Sub ImporteFavorisRepertoire(ByRef fldRepertoireCourant As Folder)
    Dim fldRepertoireSous As Folder
    Dim strCategorieNom As String * 60
    Dim strCategorieTitre As String * 60
    Dim strCategorieDescription As String * 200
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo ImporteFavorisRepertoire_Erreur
    
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
    mdiWebBase.sbrDeltaWebBase.SimpleText = "REPERTOIRE:=" & Trim(fldRepertoireCourant.Path)
    
    With DEWebBase
                               
        For Each fldRepertoireSous In fldRepertoireCourant.SubFolders
    
'FIXIT: Remplacer la fonction 'UCase' par la fonction 'UCase$'                             FixIT90210ae-R9757-R1B8ZE
'FIXIT: Remplacer la fonction 'Mid' par la fonction 'Mid$'                                 FixIT90210ae-R9757-R1B8ZE
            strCategorieNom = Replace(UCase(Mid(fldRepertoireSous.Path, Len(guwbpParametre.strFavorisRepertoire) + 2) & "\"), Chr$(39), " ")
            
'FIXIT: Remplacer la fonction 'UCase' par la fonction 'UCase$'                             FixIT90210ae-R9757-R1B8ZE
            strCategorieTitre = UCase(fldRepertoireSous.Name)
    
            gfprFicProfiler.LectureValeur fldRepertoireSous & "\Desktop.ini", "DELTA", "DESCRIPTION", strCategorieDescription, "Aucunne Description"
                                        
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
            .rsDEcmdTblLiensCategorie.Filter = "[CATEGORIENOM] = " & Chr$(39) & Trim(strCategorieNom) & Chr$(39)
            
            Select Case .rsDEcmdTblLiensCategorie.EOF
                Case True
                
                    .rsDEcmdTblLiensCategorie.AddNew
                Case False
                
            End Select
                                                                              
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
            .rsDEcmdTblLiensCategorie("CATEGORIENOM") = Trim(strCategorieNom)
            
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
            .rsDEcmdTblLiensCategorie("CATEGORIETITRE") = Trim(strCategorieTitre)
            
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
            .rsDEcmdTblLiensCategorie("CATEGORIEDESCRIPTION") = Trim(Replace(strCategorieDescription, "%10%13", vbCrLf))
                            
            .rsDEcmdTblLiensCategorie.Update
            
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
            mdiWebBase.sbrDeltaWebBase.SimpleText = "AJOUTE:=CATEGORIE:-NOM:" & Trim(strCategorieNom) & "-TITRE:" & Trim(strCategorieTitre) & "-DESCRIPTION:" & Trim(strCategorieDescription)

            ImporteFavorisFichier fldRepertoireSous
            
            ImporteFavorisRepertoire fldRepertoireSous
        Next
    End With
    
    ' Fin
ImporteFavorisRepertoire_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
ImporteFavorisRepertoire_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, LIBELLEPROCEDURE, "ImporteFavorisRepertoire", fldRepertoireCourant.Name, Err
    ' Je Continue
    Resume ImporteFavorisRepertoire_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Private Sub ImporteFavorisFichier(ByRef fldRepertoireCourant As Folder)
    Dim filFichierCourant As File
    Dim strLiensSite As String * 60
    Dim strLiensDate As String * 10
    Dim strLiensUrl As String * 60
    Dim strLiensPhoto As String * 60
    Dim strLiensDescription As String * 200
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo ImporteFavorisFichier_Erreur
        
    With DEWebBase
                   
        For Each filFichierCourant In fldRepertoireCourant.Files
        
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
            mdiWebBase.sbrDeltaWebBase.SimpleText = "FICHIER:=" & Trim(filFichierCourant.Path)

            DoEvents
            
            Select Case filFichierCourant.Type
                Case "Raccourci Internet", "Internet Shortcut"
                                       
'FIXIT: Remplacer la fonction 'UCase' par la fonction 'UCase$'                             FixIT90210ae-R9757-R1B8ZE
'FIXIT: Remplacer la fonction 'Left' par la fonction 'Left$'                               FixIT90210ae-R9757-R1B8ZE
                    strLiensSite = Replace(UCase(Left(filFichierCourant.Name, Len(filFichierCourant.Name) - 4)), Chr$(39), " ")
                                        
                    gfprFicProfiler.LectureValeur filFichierCourant.Path, "InternetShortcut", "URL", strLiensUrl, vbNullString
                    
'FIXIT: Remplacer la fonction 'LCase' par la fonction 'LCase$'                             FixIT90210ae-R9757-R1B8ZE
                    strLiensUrl = Replace(LCase(strLiensUrl), Chr$(39), " ")
                    
                    Select Case strLiensUrl
                        Case vbNullString
                        
                        Case Else
                                                
                            strLiensDate = Date
                            
                            gfprFicProfiler.LectureValeur filFichierCourant.Path, "DELTA", "PHOTO", strLiensPhoto, "default.gif"
                            
                            gfprFicProfiler.LectureValeur filFichierCourant.Path, "DELTA", "DESCRIPTION", strLiensDescription, "Internet Favoris"
                                                                            
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
                            .rsDEcmdTblLiensSite.Filter = "[CATEGORIENUMERO] = " & Chr$(39) & .rsDEcmdTblLiensCategorie("CATEGORIENUMERO") & Chr$(39) & " AND [LIENSSITE] = " & Chr$(39) & Trim(strLiensSite) & Chr$(39)
                                    
                            Select Case .rsDEcmdTblLiensSite.EOF
                                Case True
                                
                                    .rsDEcmdTblLiensSite.AddNew
                                Case False
                                
                            End Select
                            
                           .rsDEcmdTblLiensSite("CATEGORIENUMERO") = .rsDEcmdTblLiensCategorie("CATEGORIENUMERO")
                           
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
                           .rsDEcmdTblLiensSite("LIENSSITE") = Trim(strLiensSite)
                           
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
                           .rsDEcmdTblLiensSite("LIENSDATE") = Trim(strLiensDate)
                                                       
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
                           .rsDEcmdTblLiensSite("LIENSURL") = Trim(strLiensUrl)
                           
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
                           .rsDEcmdTblLiensSite("LIENSPHOTO") = Trim(strLiensPhoto)
                           
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
                           .rsDEcmdTblLiensSite("LIENSDESCRIPTION") = Trim(Replace(strLiensDescription, "%10%13", vbCrLf))
                           
                           .rsDEcmdTblLiensSite.Update
                           
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
                           mdiWebBase.sbrDeltaWebBase.SimpleText = "AJOUTE:=LIENS:-SITE:" & Trim(strLiensSite) & "-URL:" & Trim(strLiensUrl) & "-DESCRIPTION:" & Trim(strLiensDescription)
                    End Select
                Case Else
                
            End Select
        Next
        
    End With
    ' Fin
ImporteFavorisFichier_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
ImporteFavorisFichier_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, LIBELLEPROCEDURE, "ImporteFavorisFichier", fldRepertoireCourant.Name, Err
    ' Je Continue
    Resume ImporteFavorisFichier_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Sub ExporteFavorisInternet()
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo ExporteFavorisInternet_Erreur
        
    mdiWebBase.sbrDeltaWebBase.SimpleText = "EXECUTE:=ExporteFavoris:Debut"
    
    ExporteFavorisRepertoire
    
    With DEWebBase
        
        .rsDEcmdTblLiensSite.Filter = vbNullString
                        
        .rsDEcmdTblLiensCategorie.Filter = vbNullString
    End With
    
    mdiWebBase.sbrDeltaWebBase.SimpleText = "EXECUTE:=ExporteFavoris:OK"
    ' Fin
ExporteFavorisInternet_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
ExporteFavorisInternet_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, LIBELLEPROCEDURE, "ExporteFavorisInternet", vbNullString, Err
    ' Je Continue
    Resume ExporteFavorisInternet_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Private Sub ExporteFavorisRepertoire()
    Dim fsoObjectFichier As New FileSystemObject                                          ' L'object de system de fichier
    Dim intNomIndex As Integer
    Dim strCategorieNom As String
    Dim strCategorieTitre As String
    Dim strCategorieDescription As String
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo ExporteFavorisRepertoire_Erreur
        
    With DEWebBase
    
        Select Case .rsDEcmdTblLiensCategorie.EOF
            Case True
            
            Case False
            
                .rsDEcmdTblLiensCategorie.MoveFirst
                    
                Do Until .rsDEcmdTblLiensCategorie.EOF = True
                
                    strCategorieNom = .rsDEcmdTblLiensCategorie("CATEGORIENOM")
                    
                    strCategorieTitre = .rsDEcmdTblLiensCategorie("CATEGORIETITRE")
                    
                    strCategorieDescription = Replace(.rsDEcmdTblLiensCategorie("CATEGORIEDESCRIPTION"), vbCrLf, "%10%13")
                    
                    mdiWebBase.sbrDeltaWebBase.SimpleText = "CATEGORIE:=" & strCategorieNom
                    
                    intNomIndex = 1
                                        
                    Do Until InStr(intNomIndex + 1, strCategorieNom, "\") = 0
                    
                        intNomIndex = InStr(intNomIndex + 1, strCategorieNom, "\")
                                                                
'FIXIT: Remplacer la fonction 'Left' par la fonction 'Left$'                               FixIT90210ae-R9757-R1B8ZE
                        Select Case fsoObjectFichier.FolderExists(guwbpParametre.strFavorisRepertoire & "\" & Left(strCategorieNom, intNomIndex))
                            Case True
                            
                            Case False
                                                    
'FIXIT: Remplacer la fonction 'Left' par la fonction 'Left$'                               FixIT90210ae-R9757-R1B8ZE
                                fsoObjectFichier.CreateFolder (guwbpParametre.strFavorisRepertoire & "\" & Left(strCategorieNom, intNomIndex))
                                                    
'FIXIT: Remplacer la fonction 'Left' par la fonction 'Left$'                               FixIT90210ae-R9757-R1B8ZE
                                mdiWebBase.sbrDeltaWebBase.SimpleText = "AJOUTE:=REPERTOIRE:" & (guwbpParametre.strFavorisRepertoire & "\" & Left(strCategorieNom, intNomIndex))
                        End Select
                        
                    Loop
                                                                                                                                                             
                    gfprFicProfiler.EcritureValeur guwbpParametre.strFavorisRepertoire & "\" & strCategorieNom & "Desktop.ini", "DELTA", "DESCRIPTION", strCategorieDescription
                                                                                                                     
                    ExporteFavorisFichier strCategorieNom
                    
                    .rsDEcmdTblLiensCategorie.MoveNext
                Loop
        End Select
    End With
    ' Fin
ExporteFavorisRepertoire_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
ExporteFavorisRepertoire_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, LIBELLEFONCTION, "ExporteFavorisRepertoire", vbNullString, Err
    ' Je Continue
    Resume ExporteFavorisRepertoire_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Private Sub ExporteFavorisFichier(ByRef strCategorieNom As String)
    Dim strLiensSite As String
    Dim strLiensDate As String
    Dim strLiensUrl As String
    Dim strLiensPhoto As String
    Dim strLiensDescription As String
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo ExporteFavorisFichier_Erreur
        
    With DEWebBase
                           
            DoEvents
                                                                                        
            .rsDEcmdTblLiensSite.Filter = "[CATEGORIENUMERO] = " & Chr$(39) & .rsDEcmdTblLiensCategorie("CATEGORIENUMERO") & Chr$(39)
                    
            Select Case .rsDEcmdTblLiensSite.EOF
                Case True
                
                Case False
                
                    .rsDEcmdTblLiensSite.MoveFirst
                        
                    Do Until .rsDEcmdTblLiensSite.EOF = True
                    
                        strLiensSite = .rsDEcmdTblLiensSite("LIENSSITE")
                                                                            
                        strLiensUrl = .rsDEcmdTblLiensSite("LIENSURL")
                        
                        strLiensPhoto = .rsDEcmdTblLiensSite("LIENSPHOTO")
                        
                        strLiensDescription = Replace(.rsDEcmdTblLiensSite("LIENSDESCRIPTION"), vbCrLf, "%10%13")
                        
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
                        mdiWebBase.sbrDeltaWebBase.SimpleText = "LIENS:=" & Trim(strLiensUrl)
                        
                        gfprFicProfiler.EcritureValeur guwbpParametre.strFavorisRepertoire & "\" & strCategorieNom & strLiensSite & ".URL", "InternetShortcut", "URL", strLiensUrl
                        
                        gfprFicProfiler.EcritureValeur guwbpParametre.strFavorisRepertoire & "\" & strCategorieNom & strLiensSite & ".URL", "DELTA", "PHOTO", strLiensPhoto
                                                        
                        gfprFicProfiler.EcritureValeur guwbpParametre.strFavorisRepertoire & "\" & strCategorieNom & strLiensSite & ".URL", "DELTA", "DESCRIPTION", strLiensDescription
                                                        
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
'FIXIT: Remplacer la fonction 'Trim' par la fonction 'Trim$'                               FixIT90210ae-R9757-R1B8ZE
                        mdiWebBase.sbrDeltaWebBase.SimpleText = "AJOUTE:=FICHIER:-SITE:" & Trim(strLiensSite) & "-URL:" & Trim(strLiensUrl)
                        
                        .rsDEcmdTblLiensSite.MoveNext
                    Loop
            End Select
        
    End With
    ' Fin
ExporteFavorisFichier_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
ExporteFavorisFichier_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogWebBase.AjouteErreur App, mconFeuilleType, mconFeuilleNom, LIBELLEPROCEDURE, "ExporteFavorisFichier", strCategorieNom, Err
    ' Je Continue
    Resume ExporteFavorisFichier_Exit
    ' Fin
End Sub

