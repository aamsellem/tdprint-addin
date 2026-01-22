' ============================================================================
' TD PRINT VARIABLES - INSTALLATION AUTOMATIQUE
' ============================================================================
' INSTRUCTIONS:
' 1. Ouvrir Word
' 2. Alt+F11 pour ouvrir l'editeur VBA
' 3. Menu Insertion > Module
' 4. Copier-coller TOUT ce code
' 5. Appuyer sur F5 ou executer "InstallTDPrint"
' 6. C'est termine ! Le bouton apparait dans le ruban.
' ============================================================================

Option Explicit

' Variables globales pour les donnees
Public Type TDVar
    P As String  ' Placeholder
    T As String  ' Type
    D As String  ' Description
End Type

Public Type TDCat
    N As String  ' Name
    V() As TDVar
    C As Long    ' Count
End Type

Public Cats(1 To 14) As TDCat
Public Favs As Collection
Private Const FAV_KEY As String = "TDPrintFav"

' ============================================================================
' POINT D'ENTREE - INSTALLER L'ADD-IN
' ============================================================================
Public Sub InstallTDPrint()
    On Error GoTo ErrHandler

    MsgBox "Installation de TD Print Variables..." & vbCrLf & vbCrLf & _
           "Un nouveau document va etre cree avec la macro.", vbInformation

    ' Creer un nouveau document base sur Normal
    Dim doc As Document
    Set doc = Documents.Add

    ' Obtenir le projet VBA
    Dim vbProj As Object
    Set vbProj = doc.VBProject

    ' Ajouter le module principal
    Dim modMain As Object
    Set modMain = vbProj.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
    modMain.Name = "ModTDPrint"
    modMain.CodeModule.AddFromString GetMainCode()

    ' Ajouter le UserForm
    Dim frm As Object
    Set frm = vbProj.VBComponents.Add(3) ' 3 = vbext_ct_MSForm
    frm.Name = "FrmTDPrint"

    ' Configurer le formulaire
    With frm.Properties
        .Item("Caption") = "TD Print Variables"
        .Item("Width") = 400
        .Item("Height") = 480
    End With

    ' Ajouter les controles
    CreateFormControls frm

    ' Ajouter le code du formulaire
    frm.CodeModule.AddFromString GetFormCode()

    ' Sauvegarder comme .dotm
    Dim savePath As String
    savePath = Environ("APPDATA") & "\Microsoft\Word\STARTUP\TDPrintVariables.dotm"

    ' Supprimer l'ancien fichier s'il existe
    On Error Resume Next
    Kill savePath
    On Error GoTo ErrHandler

    doc.SaveAs2 savePath, wdFormatXMLTemplateMacroEnabled

    MsgBox "Installation terminee !" & vbCrLf & vbCrLf & _
           "Fichier cree : " & savePath & vbCrLf & vbCrLf & _
           "Redemarrez Word, puis utilisez Alt+F8 > TDPrint", vbInformation

    Exit Sub

ErrHandler:
    MsgBox "Erreur: " & Err.Description & vbCrLf & vbCrLf & _
           "Verifiez que l'acces au VBA est autorise:" & vbCrLf & _
           "Fichier > Options > Centre de gestion de la confidentialite > " & vbCrLf & _
           "Parametres > Parametres des macros > " & vbCrLf & _
           "Cocher 'Accorder l'acces au modele objet des projets VBA'", vbCritical
End Sub

' ============================================================================
' CREATION DES CONTROLES DU FORMULAIRE
' ============================================================================
Private Sub CreateFormControls(frm As Object)
    Dim ctrl As Object

    ' --- RECHERCHE ---
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblSearch": .Caption = "Rechercher:": .Left = 10: .Top = 10: .Width = 60: .Height = 15
    End With

    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl
        .Name = "txtSearch": .Left = 75: .Top = 8: .Width = 250: .Height = 20
    End With

    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "btnClear": .Caption = "X": .Left = 330: .Top = 8: .Width = 25: .Height = 20
    End With

    ' --- FAVORIS ---
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblFav": .Caption = "* FAVORIS": .Left = 10: .Top = 38: .Width = 100: .Height = 15: .Font.Bold = True
    End With

    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblFavCount": .Caption = "": .Left = 300: .Top = 38: .Width = 80: .Height = 15
    End With

    Set ctrl = frm.Designer.Controls.Add("Forms.ListBox.1")
    With ctrl
        .Name = "lstFav": .Left = 10: .Top = 55: .Width = 375: .Height = 55: .ColumnCount = 2: .ColumnWidths = "200;170"
    End With

    ' --- CATEGORIES ---
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblCat": .Caption = "CATEGORIES": .Left = 10: .Top = 118: .Width = 100: .Height = 15: .Font.Bold = True
    End With

    Set ctrl = frm.Designer.Controls.Add("Forms.ListBox.1")
    With ctrl
        .Name = "lstCat": .Left = 10: .Top = 135: .Width = 375: .Height = 90: .ColumnCount = 2: .ColumnWidths = "320;50"
    End With

    ' --- VARIABLES ---
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblVar": .Caption = "VARIABLES": .Left = 10: .Top = 233: .Width = 100: .Height = 15: .Font.Bold = True
    End With

    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblVarCount": .Caption = "": .Left = 300: .Top = 233: .Width = 80: .Height = 15
    End With

    Set ctrl = frm.Designer.Controls.Add("Forms.ListBox.1")
    With ctrl
        .Name = "lstVar": .Left = 10: .Top = 250: .Width = 375: .Height = 130: .ColumnCount = 2: .ColumnWidths = "200;170"
    End With

    ' --- BOUTONS ---
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "btnAddFav": .Caption = "* Ajouter favori": .Left = 10: .Top = 388: .Width = 100: .Height = 25
    End With

    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "btnRemFav": .Caption = "Retirer favori": .Left = 115: .Top = 388: .Width = 100: .Height = 25
    End With

    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "btnInsert": .Caption = "INSERER": .Left = 10: .Top = 420: .Width = 180: .Height = 30: .Font.Bold = True
    End With

    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "btnClose": .Caption = "Fermer": .Left = 285: .Top = 420: .Width = 100: .Height = 30
    End With
End Sub

' ============================================================================
' CODE DU MODULE PRINCIPAL
' ============================================================================
Private Function GetMainCode() As String
    Dim s As String
    s = "Option Explicit" & vbCrLf & vbCrLf
    s = s & "Public Type TDVar" & vbCrLf
    s = s & "    P As String" & vbCrLf
    s = s & "    T As String" & vbCrLf
    s = s & "    D As String" & vbCrLf
    s = s & "End Type" & vbCrLf & vbCrLf
    s = s & "Public Type TDCat" & vbCrLf
    s = s & "    N As String" & vbCrLf
    s = s & "    V() As TDVar" & vbCrLf
    s = s & "    C As Long" & vbCrLf
    s = s & "End Type" & vbCrLf & vbCrLf
    s = s & "Public Cats(1 To 14) As TDCat" & vbCrLf
    s = s & "Public Favs As Collection" & vbCrLf
    s = s & "Private Const FAV_KEY As String = ""TDPrintFav""" & vbCrLf & vbCrLf
    s = s & "Public Sub TDPrint()" & vbCrLf
    s = s & "    InitData" & vbCrLf
    s = s & "    LoadFavs" & vbCrLf
    s = s & "    FrmTDPrint.Show vbModeless" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & GetDataInitCode()
    s = s & GetFavoritesCode()
    GetMainCode = s
End Function

Private Function GetDataInitCode() As String
    Dim s As String
    s = "Private Sub AddV(i As Long, p As String, t As String, d As String)" & vbCrLf
    s = s & "    Cats(i).C = Cats(i).C + 1" & vbCrLf
    s = s & "    ReDim Preserve Cats(i).V(1 To Cats(i).C)" & vbCrLf
    s = s & "    Cats(i).V(Cats(i).C).P = p" & vbCrLf
    s = s & "    Cats(i).V(Cats(i).C).T = t" & vbCrLf
    s = s & "    Cats(i).V(Cats(i).C).D = d" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Public Sub InitData()" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    For i = 1 To 14: Cats(i).C = 0: Next" & vbCrLf & vbCrLf
    ' Categorie 1: Dossier
    s = s & "    Cats(1).N = ""Dossier""" & vbCrLf
    s = s & "    AddV 1, ""${C_dossier_id}"", ""C"", ""Identifiant du dossier""" & vbCrLf
    s = s & "    AddV 1, ""${C_dossier_reference}"", ""C"", ""Reference du dossier""" & vbCrLf
    s = s & "    AddV 1, ""${C_dossier_date_creation}"", ""C"", ""Date de creation""" & vbCrLf
    s = s & "    AddV 1, ""${C_dossier_date_signature}"", ""C"", ""Date de signature""" & vbCrLf
    s = s & "    AddV 1, ""${C_dossier_statut}"", ""C"", ""Statut du dossier""" & vbCrLf
    s = s & "    AddV 1, ""${C_dossier_montant_finance}"", ""C"", ""Montant finance""" & vbCrLf
    s = s & "    AddV 1, ""${C_dossier_duree}"", ""C"", ""Duree en mois""" & vbCrLf
    s = s & "    AddV 1, ""${C_dossier_taux}"", ""C"", ""Taux d'interet""" & vbCrLf
    s = s & "    AddV 1, ""${C_dossier_loyer}"", ""C"", ""Loyer mensuel""" & vbCrLf
    s = s & "    AddV 1, ""${C_dossier_valeur_residuelle}"", ""C"", ""Valeur residuelle""" & vbCrLf
    s = s & "    AddV 1, ""${B_dossier_avec_assurance}"", ""B"", ""Avec assurance""" & vbCrLf
    s = s & "    AddV 1, ""${B_dossier_avec_maintenance}"", ""B"", ""Avec maintenance""" & vbCrLf
    ' Categorie 2: Dossier Fournisseur
    s = s & "    Cats(2).N = ""Dossier Fournisseur""" & vbCrLf
    s = s & "    AddV 2, ""${C_dossier_fournisseur_id}"", ""C"", ""ID dossier fournisseur""" & vbCrLf
    s = s & "    AddV 2, ""${C_dossier_fournisseur_reference}"", ""C"", ""Reference fournisseur""" & vbCrLf
    s = s & "    AddV 2, ""${C_dossier_fournisseur_date_commande}"", ""C"", ""Date commande""" & vbCrLf
    s = s & "    AddV 2, ""${C_dossier_fournisseur_date_livraison}"", ""C"", ""Date livraison""" & vbCrLf
    s = s & "    AddV 2, ""${C_dossier_fournisseur_montant_ht}"", ""C"", ""Montant HT""" & vbCrLf
    s = s & "    AddV 2, ""${C_dossier_fournisseur_montant_ttc}"", ""C"", ""Montant TTC""" & vbCrLf
    s = s & "    AddV 2, ""${B_dossier_fournisseur_livre}"", ""B"", ""Livre""" & vbCrLf
    ' Categorie 3: Client
    s = s & "    Cats(3).N = ""Client""" & vbCrLf
    s = s & "    AddV 3, ""${C_client_id}"", ""C"", ""ID client""" & vbCrLf
    s = s & "    AddV 3, ""${C_client_raison_sociale}"", ""C"", ""Raison sociale""" & vbCrLf
    s = s & "    AddV 3, ""${C_client_siret}"", ""C"", ""SIRET""" & vbCrLf
    s = s & "    AddV 3, ""${C_client_siren}"", ""C"", ""SIREN""" & vbCrLf
    s = s & "    AddV 3, ""${C_client_adresse}"", ""C"", ""Adresse""" & vbCrLf
    s = s & "    AddV 3, ""${C_client_code_postal}"", ""C"", ""Code postal""" & vbCrLf
    s = s & "    AddV 3, ""${C_client_ville}"", ""C"", ""Ville""" & vbCrLf
    s = s & "    AddV 3, ""${C_client_pays}"", ""C"", ""Pays""" & vbCrLf
    s = s & "    AddV 3, ""${C_client_telephone}"", ""C"", ""Telephone""" & vbCrLf
    s = s & "    AddV 3, ""${C_client_email}"", ""C"", ""Email""" & vbCrLf
    s = s & "    AddV 3, ""${C_client_contact_nom}"", ""C"", ""Nom du contact""" & vbCrLf
    s = s & "    AddV 3, ""${C_client_contact_prenom}"", ""C"", ""Prenom du contact""" & vbCrLf
    s = s & "    AddV 3, ""${C_client_forme_juridique}"", ""C"", ""Forme juridique""" & vbCrLf
    s = s & "    AddV 3, ""${C_client_capital}"", ""C"", ""Capital""" & vbCrLf
    s = s & "    AddV 3, ""${C_client_rcs}"", ""C"", ""RCS""" & vbCrLf
    s = s & "    AddV 3, ""${I_client_logo}"", ""I"", ""Logo client""" & vbCrLf
    ' Categorie 4: Commercial
    s = s & "    Cats(4).N = ""Commercial""" & vbCrLf
    s = s & "    AddV 4, ""${C_commercial_id}"", ""C"", ""ID commercial""" & vbCrLf
    s = s & "    AddV 4, ""${C_commercial_nom}"", ""C"", ""Nom""" & vbCrLf
    s = s & "    AddV 4, ""${C_commercial_prenom}"", ""C"", ""Prenom""" & vbCrLf
    s = s & "    AddV 4, ""${C_commercial_email}"", ""C"", ""Email""" & vbCrLf
    s = s & "    AddV 4, ""${C_commercial_telephone}"", ""C"", ""Telephone""" & vbCrLf
    s = s & "    AddV 4, ""${C_commercial_agence}"", ""C"", ""Agence""" & vbCrLf
    s = s & "    AddV 4, ""${I_commercial_signature}"", ""I"", ""Signature""" & vbCrLf
    ' Categorie 5: Fournisseur
    s = s & "    Cats(5).N = ""Fournisseur""" & vbCrLf
    s = s & "    AddV 5, ""${C_fournisseur_id}"", ""C"", ""ID fournisseur""" & vbCrLf
    s = s & "    AddV 5, ""${C_fournisseur_raison_sociale}"", ""C"", ""Raison sociale""" & vbCrLf
    s = s & "    AddV 5, ""${C_fournisseur_siret}"", ""C"", ""SIRET""" & vbCrLf
    s = s & "    AddV 5, ""${C_fournisseur_adresse}"", ""C"", ""Adresse""" & vbCrLf
    s = s & "    AddV 5, ""${C_fournisseur_code_postal}"", ""C"", ""Code postal""" & vbCrLf
    s = s & "    AddV 5, ""${C_fournisseur_ville}"", ""C"", ""Ville""" & vbCrLf
    s = s & "    AddV 5, ""${C_fournisseur_telephone}"", ""C"", ""Telephone""" & vbCrLf
    s = s & "    AddV 5, ""${C_fournisseur_email}"", ""C"", ""Email""" & vbCrLf
    s = s & "    AddV 5, ""${C_fournisseur_iban}"", ""C"", ""IBAN""" & vbCrLf
    s = s & "    AddV 5, ""${C_fournisseur_bic}"", ""C"", ""BIC""" & vbCrLf
    ' Categorie 6: SPV
    s = s & "    Cats(6).N = ""Societe Portage (SPV)""" & vbCrLf
    s = s & "    AddV 6, ""${C_spv_id}"", ""C"", ""ID SPV""" & vbCrLf
    s = s & "    AddV 6, ""${C_spv_raison_sociale}"", ""C"", ""Raison sociale""" & vbCrLf
    s = s & "    AddV 6, ""${C_spv_siret}"", ""C"", ""SIRET""" & vbCrLf
    s = s & "    AddV 6, ""${C_spv_adresse}"", ""C"", ""Adresse""" & vbCrLf
    s = s & "    AddV 6, ""${C_spv_capital}"", ""C"", ""Capital""" & vbCrLf
    s = s & "    AddV 6, ""${C_spv_rcs}"", ""C"", ""RCS""" & vbCrLf
    s = s & "    AddV 6, ""${C_spv_representant_nom}"", ""C"", ""Representant""" & vbCrLf
    ' Categorie 7: Produit
    s = s & "    Cats(7).N = ""Produit""" & vbCrLf
    s = s & "    AddV 7, ""${C_produit_id}"", ""C"", ""ID produit""" & vbCrLf
    s = s & "    AddV 7, ""${C_produit_designation}"", ""C"", ""Designation""" & vbCrLf
    s = s & "    AddV 7, ""${C_produit_reference}"", ""C"", ""Reference""" & vbCrLf
    s = s & "    AddV 7, ""${C_produit_categorie}"", ""C"", ""Categorie""" & vbCrLf
    s = s & "    AddV 7, ""${C_produit_marque}"", ""C"", ""Marque""" & vbCrLf
    s = s & "    AddV 7, ""${C_produit_quantite}"", ""C"", ""Quantite""" & vbCrLf
    s = s & "    AddV 7, ""${C_produit_prix_unitaire_ht}"", ""C"", ""Prix unitaire HT""" & vbCrLf
    s = s & "    AddV 7, ""${C_produit_prix_total_ht}"", ""C"", ""Prix total HT""" & vbCrLf
    s = s & "    AddV 7, ""${B_produit_neuf}"", ""B"", ""Produit neuf""" & vbCrLf
    s = s & "    AddV 7, ""${T_produits_liste}"", ""T"", ""Liste des produits""" & vbCrLf
    ' Categorie 8: Assurance Produit
    s = s & "    Cats(8).N = ""Produit Assurance""" & vbCrLf
    s = s & "    AddV 8, ""${C_assurance_produit_id}"", ""C"", ""ID assurance""" & vbCrLf
    s = s & "    AddV 8, ""${C_assurance_produit_type}"", ""C"", ""Type assurance""" & vbCrLf
    s = s & "    AddV 8, ""${C_assurance_produit_montant}"", ""C"", ""Montant prime""" & vbCrLf
    s = s & "    AddV 8, ""${C_assurance_produit_franchise}"", ""C"", ""Franchise""" & vbCrLf
    s = s & "    AddV 8, ""${B_assurance_produit_active}"", ""B"", ""Active""" & vbCrLf
    ' Categorie 9: Caution
    s = s & "    Cats(9).N = ""Caution""" & vbCrLf
    s = s & "    AddV 9, ""${C_caution_id}"", ""C"", ""ID caution""" & vbCrLf
    s = s & "    AddV 9, ""${C_caution_type}"", ""C"", ""Type""" & vbCrLf
    s = s & "    AddV 9, ""${C_caution_nom}"", ""C"", ""Nom""" & vbCrLf
    s = s & "    AddV 9, ""${C_caution_prenom}"", ""C"", ""Prenom""" & vbCrLf
    s = s & "    AddV 9, ""${C_caution_adresse}"", ""C"", ""Adresse""" & vbCrLf
    s = s & "    AddV 9, ""${C_caution_montant}"", ""C"", ""Montant""" & vbCrLf
    s = s & "    AddV 9, ""${B_caution_solidaire}"", ""B"", ""Solidaire""" & vbCrLf
    s = s & "    AddV 9, ""${T_cautions_liste}"", ""T"", ""Liste cautions""" & vbCrLf
    ' Categorie 10: Mandat Caution
    s = s & "    Cats(10).N = ""Mandat Caution""" & vbCrLf
    s = s & "    AddV 10, ""${C_mandat_caution_id}"", ""C"", ""ID mandat""" & vbCrLf
    s = s & "    AddV 10, ""${C_mandat_caution_reference}"", ""C"", ""Reference""" & vbCrLf
    s = s & "    AddV 10, ""${C_mandat_caution_montant}"", ""C"", ""Montant""" & vbCrLf
    s = s & "    AddV 10, ""${C_mandat_caution_duree}"", ""C"", ""Duree""" & vbCrLf
    s = s & "    AddV 10, ""${B_mandat_caution_actif}"", ""B"", ""Actif""" & vbCrLf
    ' Categorie 11: Credit Fournisseur
    s = s & "    Cats(11).N = ""Credit Fournisseur""" & vbCrLf
    s = s & "    AddV 11, ""${C_credit_fournisseur_id}"", ""C"", ""ID credit""" & vbCrLf
    s = s & "    AddV 11, ""${C_credit_fournisseur_montant}"", ""C"", ""Montant""" & vbCrLf
    s = s & "    AddV 11, ""${C_credit_fournisseur_taux}"", ""C"", ""Taux""" & vbCrLf
    s = s & "    AddV 11, ""${B_credit_fournisseur_actif}"", ""B"", ""Actif""" & vbCrLf
    ' Categorie 12: Attestation Prix Net
    s = s & "    Cats(12).N = ""Attestation Prix Net""" & vbCrLf
    s = s & "    AddV 12, ""${C_attestation_prix_net_id}"", ""C"", ""ID attestation""" & vbCrLf
    s = s & "    AddV 12, ""${C_attestation_prix_net_montant_ht}"", ""C"", ""Montant HT""" & vbCrLf
    s = s & "    AddV 12, ""${C_attestation_prix_net_montant_remise}"", ""C"", ""Remise""" & vbCrLf
    s = s & "    AddV 12, ""${B_attestation_prix_net_valide}"", ""B"", ""Valide""" & vbCrLf
    ' Categorie 13: Societe Assurance
    s = s & "    Cats(13).N = ""Societe Assurance""" & vbCrLf
    s = s & "    AddV 13, ""${C_societe_assurance_id}"", ""C"", ""ID societe""" & vbCrLf
    s = s & "    AddV 13, ""${C_societe_assurance_nom}"", ""C"", ""Nom""" & vbCrLf
    s = s & "    AddV 13, ""${C_societe_assurance_adresse}"", ""C"", ""Adresse""" & vbCrLf
    s = s & "    AddV 13, ""${C_societe_assurance_telephone}"", ""C"", ""Telephone""" & vbCrLf
    s = s & "    AddV 13, ""${I_societe_assurance_logo}"", ""I"", ""Logo""" & vbCrLf
    ' Categorie 14: Simulation
    s = s & "    Cats(14).N = ""Simulation""" & vbCrLf
    s = s & "    AddV 14, ""${C_simulation_id}"", ""C"", ""ID simulation""" & vbCrLf
    s = s & "    AddV 14, ""${C_simulation_montant}"", ""C"", ""Montant""" & vbCrLf
    s = s & "    AddV 14, ""${C_simulation_duree}"", ""C"", ""Duree""" & vbCrLf
    s = s & "    AddV 14, ""${C_simulation_taux}"", ""C"", ""Taux""" & vbCrLf
    s = s & "    AddV 14, ""${C_simulation_loyer}"", ""C"", ""Loyer""" & vbCrLf
    s = s & "    AddV 14, ""${C_simulation_cout_total}"", ""C"", ""Cout total""" & vbCrLf
    s = s & "    AddV 14, ""${B_simulation_avec_assurance}"", ""B"", ""Avec assurance""" & vbCrLf
    s = s & "    AddV 14, ""${T_simulation_echeancier}"", ""T"", ""Echeancier""" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    GetDataInitCode = s
End Function

Private Function GetFavoritesCode() As String
    Dim s As String
    s = "Public Sub LoadFavs()" & vbCrLf
    s = s & "    Set Favs = New Collection" & vbCrLf
    s = s & "    Dim fs As String, fa() As String, i As Long" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    fs = GetSetting(""TDPrint"", ""S"", FAV_KEY, """")" & vbCrLf
    s = s & "    If fs <> """" Then" & vbCrLf
    s = s & "        fa = Split(fs, ""|"")" & vbCrLf
    s = s & "        For i = LBound(fa) To UBound(fa)" & vbCrLf
    s = s & "            If fa(i) <> """" Then Favs.Add fa(i), fa(i)" & vbCrLf
    s = s & "        Next" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Public Sub SaveFavs()" & vbCrLf
    s = s & "    Dim fs As String, v As Variant" & vbCrLf
    s = s & "    For Each v In Favs" & vbCrLf
    s = s & "        If fs <> """" Then fs = fs & ""|""" & vbCrLf
    s = s & "        fs = fs & v" & vbCrLf
    s = s & "    Next" & vbCrLf
    s = s & "    SaveSetting ""TDPrint"", ""S"", FAV_KEY, fs" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Public Function IsFav(p As String) As Boolean" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Dim x: x = Favs(p)" & vbCrLf
    s = s & "    IsFav = (Err.Number = 0)" & vbCrLf
    s = s & "End Function" & vbCrLf
    GetFavoritesCode = s
End Function

' ============================================================================
' CODE DU USERFORM
' ============================================================================
Private Function GetFormCode() As String
    Dim s As String
    s = "Option Explicit" & vbCrLf
    s = s & "Private flt As String" & vbCrLf & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    flt = """": LoadCats: LoadFavList" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub LoadCats()" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    lstCat.Clear" & vbCrLf
    s = s & "    For i = 1 To 14" & vbCrLf
    s = s & "        lstCat.AddItem Cats(i).N" & vbCrLf
    s = s & "        lstCat.List(lstCat.ListCount - 1, 1) = Cats(i).C" & vbCrLf
    s = s & "    Next" & vbCrLf
    s = s & "    If lstCat.ListCount > 0 Then lstCat.ListIndex = 0" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub lstCat_Click()" & vbCrLf
    s = s & "    If lstCat.ListIndex < 0 Then Exit Sub" & vbCrLf
    s = s & "    Dim ci As Long, i As Long, v As TDVar" & vbCrLf
    s = s & "    ci = lstCat.ListIndex + 1" & vbCrLf
    s = s & "    lstVar.Clear" & vbCrLf
    s = s & "    For i = 1 To Cats(ci).C" & vbCrLf
    s = s & "        v = Cats(ci).V(i)" & vbCrLf
    s = s & "        If flt = """" Or Match(v.P, v.D, Cats(ci).N, flt) Then" & vbCrLf
    s = s & "            lstVar.AddItem v.P" & vbCrLf
    s = s & "            lstVar.List(lstVar.ListCount - 1, 1) = v.D" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next" & vbCrLf
    s = s & "    lblVarCount.Caption = lstVar.ListCount & "" var.""" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Function Match(p As String, d As String, c As String, f As String) As Boolean" & vbCrLf
    s = s & "    Dim lf As String: lf = LCase(f)" & vbCrLf
    s = s & "    Match = InStr(1, LCase(p), lf) > 0 Or InStr(1, LCase(d), lf) > 0 Or InStr(1, LCase(c), lf) > 0" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf
    s = s & "Private Sub txtSearch_Change()" & vbCrLf
    s = s & "    flt = Trim(txtSearch.Text)" & vbCrLf
    s = s & "    If lstCat.ListIndex >= 0 Then lstCat_Click" & vbCrLf
    s = s & "    LoadFavList" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub btnClear_Click()" & vbCrLf
    s = s & "    txtSearch.Text = """": txtSearch.SetFocus" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub LoadFavList()" & vbCrLf
    s = s & "    Dim p As Variant, i As Long, j As Long, v As TDVar" & vbCrLf
    s = s & "    lstFav.Clear" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    For Each p In Favs" & vbCrLf
    s = s & "        For i = 1 To 14" & vbCrLf
    s = s & "            For j = 1 To Cats(i).C" & vbCrLf
    s = s & "                If Cats(i).V(j).P = p Then" & vbCrLf
    s = s & "                    v = Cats(i).V(j)" & vbCrLf
    s = s & "                    If flt = """" Or Match(v.P, v.D, Cats(i).N, flt) Then" & vbCrLf
    s = s & "                        lstFav.AddItem v.P" & vbCrLf
    s = s & "                        lstFav.List(lstFav.ListCount - 1, 1) = v.D" & vbCrLf
    s = s & "                    End If" & vbCrLf
    s = s & "                    Exit For" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "            Next" & vbCrLf
    s = s & "        Next" & vbCrLf
    s = s & "    Next" & vbCrLf
    s = s & "    lblFavCount.Caption = lstFav.ListCount & "" fav.""" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub lstVar_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    InsertVar" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub lstFav_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    InsertFav" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub btnInsert_Click()" & vbCrLf
    s = s & "    If lstVar.ListIndex >= 0 Then InsertVar" & vbCrLf
    s = s & "    If lstFav.ListIndex >= 0 Then InsertFav" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub InsertVar()" & vbCrLf
    s = s & "    If lstVar.ListIndex < 0 Then Exit Sub" & vbCrLf
    s = s & "    Selection.TypeText lstVar.List(lstVar.ListIndex, 0)" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub InsertFav()" & vbCrLf
    s = s & "    If lstFav.ListIndex < 0 Then Exit Sub" & vbCrLf
    s = s & "    Selection.TypeText lstFav.List(lstFav.ListIndex, 0)" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub btnAddFav_Click()" & vbCrLf
    s = s & "    If lstVar.ListIndex < 0 Then Exit Sub" & vbCrLf
    s = s & "    Dim p As String: p = lstVar.List(lstVar.ListIndex, 0)" & vbCrLf
    s = s & "    If IsFav(p) Then MsgBox ""Deja en favori"", vbInformation: Exit Sub" & vbCrLf
    s = s & "    Favs.Add p, p: SaveFavs: LoadFavList" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub btnRemFav_Click()" & vbCrLf
    s = s & "    If lstFav.ListIndex < 0 Then Exit Sub" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Favs.Remove lstFav.List(lstFav.ListIndex, 0)" & vbCrLf
    s = s & "    SaveFavs: LoadFavList" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub lstVar_Click(): lstFav.ListIndex = -1: End Sub" & vbCrLf
    s = s & "Private Sub lstFav_Click(): lstVar.ListIndex = -1: End Sub" & vbCrLf
    GetFormCode = s
End Function
