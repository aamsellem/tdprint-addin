Attribute VB_Name = "ModVariables"
Option Explicit

' ============================================================================
' Module: ModVariables
' Description: Données des variables TD Print organisées par catégories
' ============================================================================

' Structure pour une variable
Public Type TDVariable
    Placeholder As String
    VarType As String       ' C=Champ, B=Booléen, I=Image, T=Tableau
    Description As String
End Type

' Structure pour une catégorie
Public Type TDCategory
    Id As String
    Name As String
    Variables() As TDVariable
    Count As Long
End Type

' Tableau global des catégories
Public Categories(1 To 14) As TDCategory

' Initialisation des données
Public Sub InitializeVariables()
    InitDossier
    InitDossierFournisseur
    InitClient
    InitCommercial
    InitFournisseur
    InitSocietePortage
    InitProduit
    InitProduitAssurance
    InitCaution
    InitMandatCaution
    InitCreditFournisseur
    InitAttestationPrixNet
    InitSocieteAssurance
    InitSimulation
End Sub

Private Sub AddVariable(ByRef cat As TDCategory, placeholder As String, varType As String, description As String)
    cat.Count = cat.Count + 1
    ReDim Preserve cat.Variables(1 To cat.Count)
    cat.Variables(cat.Count).Placeholder = placeholder
    cat.Variables(cat.Count).VarType = varType
    cat.Variables(cat.Count).Description = description
End Sub

Private Sub InitDossier()
    With Categories(1)
        .Id = "dossier"
        .Name = "Dossier"
        .Count = 0
        AddVariable Categories(1), "${C_dossier_id}", "C", "Identifiant unique du dossier"
        AddVariable Categories(1), "${C_dossier_reference}", "C", "Référence du dossier"
        AddVariable Categories(1), "${C_dossier_date_creation}", "C", "Date de création du dossier"
        AddVariable Categories(1), "${C_dossier_date_signature}", "C", "Date de signature du contrat"
        AddVariable Categories(1), "${C_dossier_statut}", "C", "Statut actuel du dossier"
        AddVariable Categories(1), "${C_dossier_montant_finance}", "C", "Montant total financé"
        AddVariable Categories(1), "${C_dossier_duree}", "C", "Durée du financement en mois"
        AddVariable Categories(1), "${C_dossier_taux}", "C", "Taux d'intérêt appliqué"
        AddVariable Categories(1), "${C_dossier_loyer}", "C", "Montant du loyer mensuel"
        AddVariable Categories(1), "${C_dossier_valeur_residuelle}", "C", "Valeur résiduelle en fin de contrat"
        AddVariable Categories(1), "${B_dossier_avec_assurance}", "B", "Indique si le dossier inclut une assurance"
        AddVariable Categories(1), "${B_dossier_avec_maintenance}", "B", "Indique si le dossier inclut la maintenance"
        AddVariable Categories(1), "${C_dossier_commentaire}", "C", "Commentaires sur le dossier"
    End With
End Sub

Private Sub InitDossierFournisseur()
    With Categories(2)
        .Id = "dossier_fournisseur"
        .Name = "Dossier Fournisseur"
        .Count = 0
        AddVariable Categories(2), "${C_dossier_fournisseur_id}", "C", "Identifiant du dossier fournisseur"
        AddVariable Categories(2), "${C_dossier_fournisseur_reference}", "C", "Référence fournisseur du dossier"
        AddVariable Categories(2), "${C_dossier_fournisseur_date_commande}", "C", "Date de la commande fournisseur"
        AddVariable Categories(2), "${C_dossier_fournisseur_date_livraison}", "C", "Date de livraison prévue"
        AddVariable Categories(2), "${C_dossier_fournisseur_montant_ht}", "C", "Montant HT de la commande"
        AddVariable Categories(2), "${C_dossier_fournisseur_montant_tva}", "C", "Montant de la TVA"
        AddVariable Categories(2), "${C_dossier_fournisseur_montant_ttc}", "C", "Montant TTC de la commande"
        AddVariable Categories(2), "${B_dossier_fournisseur_livre}", "B", "Indique si la commande est livrée"
        AddVariable Categories(2), "${B_dossier_fournisseur_facture}", "B", "Indique si la facture est reçue"
    End With
End Sub

Private Sub InitClient()
    With Categories(3)
        .Id = "client"
        .Name = "Client"
        .Count = 0
        AddVariable Categories(3), "${C_client_id}", "C", "Identifiant unique du client"
        AddVariable Categories(3), "${C_client_raison_sociale}", "C", "Raison sociale du client"
        AddVariable Categories(3), "${C_client_siret}", "C", "Numéro SIRET du client"
        AddVariable Categories(3), "${C_client_siren}", "C", "Numéro SIREN du client"
        AddVariable Categories(3), "${C_client_adresse}", "C", "Adresse complète du client"
        AddVariable Categories(3), "${C_client_code_postal}", "C", "Code postal du client"
        AddVariable Categories(3), "${C_client_ville}", "C", "Ville du client"
        AddVariable Categories(3), "${C_client_pays}", "C", "Pays du client"
        AddVariable Categories(3), "${C_client_telephone}", "C", "Téléphone du client"
        AddVariable Categories(3), "${C_client_email}", "C", "Email du client"
        AddVariable Categories(3), "${C_client_contact_nom}", "C", "Nom du contact principal"
        AddVariable Categories(3), "${C_client_contact_prenom}", "C", "Prénom du contact principal"
        AddVariable Categories(3), "${C_client_contact_fonction}", "C", "Fonction du contact principal"
        AddVariable Categories(3), "${C_client_forme_juridique}", "C", "Forme juridique de l'entreprise"
        AddVariable Categories(3), "${C_client_capital}", "C", "Capital social de l'entreprise"
        AddVariable Categories(3), "${C_client_naf}", "C", "Code NAF de l'entreprise"
        AddVariable Categories(3), "${C_client_rcs}", "C", "RCS du client"
        AddVariable Categories(3), "${C_client_tva_intra}", "C", "Numéro de TVA intracommunautaire"
        AddVariable Categories(3), "${I_client_logo}", "I", "Logo du client"
    End With
End Sub

Private Sub InitCommercial()
    With Categories(4)
        .Id = "commercial"
        .Name = "Commercial"
        .Count = 0
        AddVariable Categories(4), "${C_commercial_id}", "C", "Identifiant du commercial"
        AddVariable Categories(4), "${C_commercial_nom}", "C", "Nom du commercial"
        AddVariable Categories(4), "${C_commercial_prenom}", "C", "Prénom du commercial"
        AddVariable Categories(4), "${C_commercial_email}", "C", "Email du commercial"
        AddVariable Categories(4), "${C_commercial_telephone}", "C", "Téléphone du commercial"
        AddVariable Categories(4), "${C_commercial_agence}", "C", "Agence du commercial"
        AddVariable Categories(4), "${C_commercial_region}", "C", "Région du commercial"
        AddVariable Categories(4), "${I_commercial_signature}", "I", "Signature du commercial"
    End With
End Sub

Private Sub InitFournisseur()
    With Categories(5)
        .Id = "fournisseur"
        .Name = "Fournisseur"
        .Count = 0
        AddVariable Categories(5), "${C_fournisseur_id}", "C", "Identifiant du fournisseur"
        AddVariable Categories(5), "${C_fournisseur_raison_sociale}", "C", "Raison sociale du fournisseur"
        AddVariable Categories(5), "${C_fournisseur_siret}", "C", "SIRET du fournisseur"
        AddVariable Categories(5), "${C_fournisseur_adresse}", "C", "Adresse du fournisseur"
        AddVariable Categories(5), "${C_fournisseur_code_postal}", "C", "Code postal du fournisseur"
        AddVariable Categories(5), "${C_fournisseur_ville}", "C", "Ville du fournisseur"
        AddVariable Categories(5), "${C_fournisseur_pays}", "C", "Pays du fournisseur"
        AddVariable Categories(5), "${C_fournisseur_telephone}", "C", "Téléphone du fournisseur"
        AddVariable Categories(5), "${C_fournisseur_email}", "C", "Email du fournisseur"
        AddVariable Categories(5), "${C_fournisseur_contact_nom}", "C", "Nom du contact fournisseur"
        AddVariable Categories(5), "${C_fournisseur_contact_prenom}", "C", "Prénom du contact fournisseur"
        AddVariable Categories(5), "${C_fournisseur_iban}", "C", "IBAN du fournisseur"
        AddVariable Categories(5), "${C_fournisseur_bic}", "C", "BIC du fournisseur"
        AddVariable Categories(5), "${I_fournisseur_logo}", "I", "Logo du fournisseur"
    End With
End Sub

Private Sub InitSocietePortage()
    With Categories(6)
        .Id = "societe_portage"
        .Name = "Société de Portage (SPV)"
        .Count = 0
        AddVariable Categories(6), "${C_spv_id}", "C", "Identifiant de la SPV"
        AddVariable Categories(6), "${C_spv_raison_sociale}", "C", "Raison sociale de la SPV"
        AddVariable Categories(6), "${C_spv_siret}", "C", "SIRET de la SPV"
        AddVariable Categories(6), "${C_spv_adresse}", "C", "Adresse de la SPV"
        AddVariable Categories(6), "${C_spv_code_postal}", "C", "Code postal de la SPV"
        AddVariable Categories(6), "${C_spv_ville}", "C", "Ville de la SPV"
        AddVariable Categories(6), "${C_spv_capital}", "C", "Capital social de la SPV"
        AddVariable Categories(6), "${C_spv_rcs}", "C", "RCS de la SPV"
        AddVariable Categories(6), "${C_spv_representant_nom}", "C", "Nom du représentant légal"
        AddVariable Categories(6), "${C_spv_representant_fonction}", "C", "Fonction du représentant"
        AddVariable Categories(6), "${I_spv_logo}", "I", "Logo de la SPV"
    End With
End Sub

Private Sub InitProduit()
    With Categories(7)
        .Id = "produit"
        .Name = "Produit"
        .Count = 0
        AddVariable Categories(7), "${C_produit_id}", "C", "Identifiant du produit"
        AddVariable Categories(7), "${C_produit_designation}", "C", "Désignation du produit"
        AddVariable Categories(7), "${C_produit_reference}", "C", "Référence du produit"
        AddVariable Categories(7), "${C_produit_categorie}", "C", "Catégorie du produit"
        AddVariable Categories(7), "${C_produit_marque}", "C", "Marque du produit"
        AddVariable Categories(7), "${C_produit_quantite}", "C", "Quantité commandée"
        AddVariable Categories(7), "${C_produit_prix_unitaire_ht}", "C", "Prix unitaire HT"
        AddVariable Categories(7), "${C_produit_prix_total_ht}", "C", "Prix total HT"
        AddVariable Categories(7), "${C_produit_taux_tva}", "C", "Taux de TVA applicable"
        AddVariable Categories(7), "${C_produit_numero_serie}", "C", "Numéro de série du produit"
        AddVariable Categories(7), "${C_produit_date_livraison}", "C", "Date de livraison du produit"
        AddVariable Categories(7), "${B_produit_neuf}", "B", "Indique si le produit est neuf"
        AddVariable Categories(7), "${T_produits_liste}", "T", "Tableau de tous les produits"
    End With
End Sub

Private Sub InitProduitAssurance()
    With Categories(8)
        .Id = "produit_assurance"
        .Name = "Produit Assurance"
        .Count = 0
        AddVariable Categories(8), "${C_assurance_produit_id}", "C", "Identifiant de l'assurance produit"
        AddVariable Categories(8), "${C_assurance_produit_type}", "C", "Type d'assurance"
        AddVariable Categories(8), "${C_assurance_produit_garantie}", "C", "Nature de la garantie"
        AddVariable Categories(8), "${C_assurance_produit_montant}", "C", "Montant de la prime"
        AddVariable Categories(8), "${C_assurance_produit_franchise}", "C", "Montant de la franchise"
        AddVariable Categories(8), "${C_assurance_produit_date_debut}", "C", "Date de début de couverture"
        AddVariable Categories(8), "${C_assurance_produit_date_fin}", "C", "Date de fin de couverture"
        AddVariable Categories(8), "${B_assurance_produit_active}", "B", "Indique si l'assurance est active"
    End With
End Sub

Private Sub InitCaution()
    With Categories(9)
        .Id = "caution"
        .Name = "Caution"
        .Count = 0
        AddVariable Categories(9), "${C_caution_id}", "C", "Identifiant de la caution"
        AddVariable Categories(9), "${C_caution_type}", "C", "Type de caution"
        AddVariable Categories(9), "${C_caution_nom}", "C", "Nom de la caution"
        AddVariable Categories(9), "${C_caution_prenom}", "C", "Prénom de la caution"
        AddVariable Categories(9), "${C_caution_adresse}", "C", "Adresse de la caution"
        AddVariable Categories(9), "${C_caution_code_postal}", "C", "Code postal de la caution"
        AddVariable Categories(9), "${C_caution_ville}", "C", "Ville de la caution"
        AddVariable Categories(9), "${C_caution_montant}", "C", "Montant de la caution"
        AddVariable Categories(9), "${C_caution_date_naissance}", "C", "Date de naissance"
        AddVariable Categories(9), "${C_caution_lieu_naissance}", "C", "Lieu de naissance"
        AddVariable Categories(9), "${C_caution_nationalite}", "C", "Nationalité"
        AddVariable Categories(9), "${B_caution_solidaire}", "B", "Indique si la caution est solidaire"
        AddVariable Categories(9), "${T_cautions_liste}", "T", "Tableau de toutes les cautions"
    End With
End Sub

Private Sub InitMandatCaution()
    With Categories(10)
        .Id = "mandat_caution"
        .Name = "Mandat Caution"
        .Count = 0
        AddVariable Categories(10), "${C_mandat_caution_id}", "C", "Identifiant du mandat de caution"
        AddVariable Categories(10), "${C_mandat_caution_reference}", "C", "Référence du mandat"
        AddVariable Categories(10), "${C_mandat_caution_date}", "C", "Date du mandat"
        AddVariable Categories(10), "${C_mandat_caution_montant}", "C", "Montant garanti"
        AddVariable Categories(10), "${C_mandat_caution_duree}", "C", "Durée du mandat"
        AddVariable Categories(10), "${C_mandat_caution_beneficiaire}", "C", "Bénéficiaire du mandat"
        AddVariable Categories(10), "${B_mandat_caution_actif}", "B", "Indique si le mandat est actif"
    End With
End Sub

Private Sub InitCreditFournisseur()
    With Categories(11)
        .Id = "credit_fournisseur"
        .Name = "Crédit Fournisseur"
        .Count = 0
        AddVariable Categories(11), "${C_credit_fournisseur_id}", "C", "Identifiant du crédit fournisseur"
        AddVariable Categories(11), "${C_credit_fournisseur_montant}", "C", "Montant du crédit fournisseur"
        AddVariable Categories(11), "${C_credit_fournisseur_date_debut}", "C", "Date de début du crédit"
        AddVariable Categories(11), "${C_credit_fournisseur_date_fin}", "C", "Date de fin du crédit"
        AddVariable Categories(11), "${C_credit_fournisseur_taux}", "C", "Taux appliqué"
        AddVariable Categories(11), "${C_credit_fournisseur_echeances}", "C", "Nombre d'échéances"
        AddVariable Categories(11), "${B_credit_fournisseur_actif}", "B", "Indique si le crédit est actif"
    End With
End Sub

Private Sub InitAttestationPrixNet()
    With Categories(12)
        .Id = "attestation_prix_net"
        .Name = "Attestation Prix Net"
        .Count = 0
        AddVariable Categories(12), "${C_attestation_prix_net_id}", "C", "Identifiant de l'attestation"
        AddVariable Categories(12), "${C_attestation_prix_net_date}", "C", "Date de l'attestation"
        AddVariable Categories(12), "${C_attestation_prix_net_montant_ht}", "C", "Montant HT attesté"
        AddVariable Categories(12), "${C_attestation_prix_net_montant_remise}", "C", "Montant de la remise"
        AddVariable Categories(12), "${C_attestation_prix_net_montant_final}", "C", "Montant final après remise"
        AddVariable Categories(12), "${C_attestation_prix_net_validite}", "C", "Durée de validité"
        AddVariable Categories(12), "${B_attestation_prix_net_valide}", "B", "Indique si l'attestation est valide"
    End With
End Sub

Private Sub InitSocieteAssurance()
    With Categories(13)
        .Id = "societe_assurance"
        .Name = "Société Assurance"
        .Count = 0
        AddVariable Categories(13), "${C_societe_assurance_id}", "C", "Identifiant de la société d'assurance"
        AddVariable Categories(13), "${C_societe_assurance_nom}", "C", "Nom de la société d'assurance"
        AddVariable Categories(13), "${C_societe_assurance_adresse}", "C", "Adresse de la société"
        AddVariable Categories(13), "${C_societe_assurance_code_postal}", "C", "Code postal"
        AddVariable Categories(13), "${C_societe_assurance_ville}", "C", "Ville"
        AddVariable Categories(13), "${C_societe_assurance_telephone}", "C", "Téléphone"
        AddVariable Categories(13), "${C_societe_assurance_email}", "C", "Email de contact"
        AddVariable Categories(13), "${C_societe_assurance_numero_agrement}", "C", "Numéro d'agrément"
        AddVariable Categories(13), "${I_societe_assurance_logo}", "I", "Logo de la société d'assurance"
    End With
End Sub

Private Sub InitSimulation()
    With Categories(14)
        .Id = "simulation"
        .Name = "Simulation"
        .Count = 0
        AddVariable Categories(14), "${C_simulation_id}", "C", "Identifiant de la simulation"
        AddVariable Categories(14), "${C_simulation_date}", "C", "Date de la simulation"
        AddVariable Categories(14), "${C_simulation_montant}", "C", "Montant simulé"
        AddVariable Categories(14), "${C_simulation_duree}", "C", "Durée de financement simulée"
        AddVariable Categories(14), "${C_simulation_taux}", "C", "Taux simulé"
        AddVariable Categories(14), "${C_simulation_loyer}", "C", "Loyer mensuel simulé"
        AddVariable Categories(14), "${C_simulation_premier_loyer}", "C", "Premier loyer majoré"
        AddVariable Categories(14), "${C_simulation_valeur_residuelle}", "C", "Valeur résiduelle simulée"
        AddVariable Categories(14), "${C_simulation_cout_total}", "C", "Coût total du financement"
        AddVariable Categories(14), "${C_simulation_periodicite}", "C", "Périodicité des loyers"
        AddVariable Categories(14), "${B_simulation_avec_assurance}", "B", "Simulation avec assurance"
        AddVariable Categories(14), "${T_simulation_echeancier}", "T", "Tableau de l'échéancier simulé"
    End With
End Sub

' Fonction pour obtenir le nom complet du type
Public Function GetTypeName(varType As String) As String
    Select Case varType
        Case "C": GetTypeName = "Champ"
        Case "B": GetTypeName = "Booléen"
        Case "I": GetTypeName = "Image"
        Case "T": GetTypeName = "Tableau"
        Case Else: GetTypeName = varType
    End Select
End Function
