/**
 * Données des variables TD Print
 * Types: C = Champ, B = Booléen, I = Image, T = Tableau
 */

export const categories = [
  {
    id: "dossier",
    name: "Dossier",
    description: "Variables liées au dossier de financement",
    variables: [
      { placeholder: "${C_dossier_id}", type: "C", description: "Identifiant unique du dossier", example: "DOS-2024-001" },
      { placeholder: "${C_dossier_reference}", type: "C", description: "Référence du dossier", example: "REF-2024-ABC" },
      { placeholder: "${C_dossier_date_creation}", type: "C", description: "Date de création du dossier", example: "15/01/2024" },
      { placeholder: "${C_dossier_date_signature}", type: "C", description: "Date de signature du contrat", example: "20/01/2024" },
      { placeholder: "${C_dossier_statut}", type: "C", description: "Statut actuel du dossier", example: "En cours" },
      { placeholder: "${C_dossier_montant_finance}", type: "C", description: "Montant total financé", example: "150 000,00 €" },
      { placeholder: "${C_dossier_duree}", type: "C", description: "Durée du financement en mois", example: "48" },
      { placeholder: "${C_dossier_taux}", type: "C", description: "Taux d'intérêt appliqué", example: "3,50 %" },
      { placeholder: "${C_dossier_loyer}", type: "C", description: "Montant du loyer mensuel", example: "3 500,00 €" },
      { placeholder: "${C_dossier_valeur_residuelle}", type: "C", description: "Valeur résiduelle en fin de contrat", example: "15 000,00 €" },
      { placeholder: "${B_dossier_avec_assurance}", type: "B", description: "Indique si le dossier inclut une assurance", example: "Oui/Non" },
      { placeholder: "${B_dossier_avec_maintenance}", type: "B", description: "Indique si le dossier inclut la maintenance", example: "Oui/Non" },
      { placeholder: "${C_dossier_commentaire}", type: "C", description: "Commentaires sur le dossier", example: "Dossier prioritaire" }
    ]
  },
  {
    id: "dossier_fournisseur",
    name: "Dossier Fournisseur",
    description: "Variables liées au dossier côté fournisseur",
    variables: [
      { placeholder: "${C_dossier_fournisseur_id}", type: "C", description: "Identifiant du dossier fournisseur", example: "DF-2024-001" },
      { placeholder: "${C_dossier_fournisseur_reference}", type: "C", description: "Référence fournisseur du dossier", example: "FOUR-REF-001" },
      { placeholder: "${C_dossier_fournisseur_date_commande}", type: "C", description: "Date de la commande fournisseur", example: "10/01/2024" },
      { placeholder: "${C_dossier_fournisseur_date_livraison}", type: "C", description: "Date de livraison prévue", example: "25/01/2024" },
      { placeholder: "${C_dossier_fournisseur_montant_ht}", type: "C", description: "Montant HT de la commande", example: "125 000,00 €" },
      { placeholder: "${C_dossier_fournisseur_montant_tva}", type: "C", description: "Montant de la TVA", example: "25 000,00 €" },
      { placeholder: "${C_dossier_fournisseur_montant_ttc}", type: "C", description: "Montant TTC de la commande", example: "150 000,00 €" },
      { placeholder: "${B_dossier_fournisseur_livre}", type: "B", description: "Indique si la commande est livrée", example: "Oui/Non" },
      { placeholder: "${B_dossier_fournisseur_facture}", type: "B", description: "Indique si la facture est reçue", example: "Oui/Non" }
    ]
  },
  {
    id: "client",
    name: "Client",
    description: "Variables liées aux informations du client",
    variables: [
      { placeholder: "${C_client_id}", type: "C", description: "Identifiant unique du client", example: "CLI-001" },
      { placeholder: "${C_client_raison_sociale}", type: "C", description: "Raison sociale du client", example: "ACME SAS" },
      { placeholder: "${C_client_siret}", type: "C", description: "Numéro SIRET du client", example: "123 456 789 00012" },
      { placeholder: "${C_client_siren}", type: "C", description: "Numéro SIREN du client", example: "123 456 789" },
      { placeholder: "${C_client_adresse}", type: "C", description: "Adresse complète du client", example: "10 rue de la Paix" },
      { placeholder: "${C_client_code_postal}", type: "C", description: "Code postal du client", example: "75001" },
      { placeholder: "${C_client_ville}", type: "C", description: "Ville du client", example: "Paris" },
      { placeholder: "${C_client_pays}", type: "C", description: "Pays du client", example: "France" },
      { placeholder: "${C_client_telephone}", type: "C", description: "Téléphone du client", example: "01 23 45 67 89" },
      { placeholder: "${C_client_email}", type: "C", description: "Email du client", example: "contact@acme.fr" },
      { placeholder: "${C_client_contact_nom}", type: "C", description: "Nom du contact principal", example: "Dupont" },
      { placeholder: "${C_client_contact_prenom}", type: "C", description: "Prénom du contact principal", example: "Jean" },
      { placeholder: "${C_client_contact_fonction}", type: "C", description: "Fonction du contact principal", example: "Directeur Financier" },
      { placeholder: "${C_client_forme_juridique}", type: "C", description: "Forme juridique de l'entreprise", example: "SAS" },
      { placeholder: "${C_client_capital}", type: "C", description: "Capital social de l'entreprise", example: "100 000,00 €" },
      { placeholder: "${C_client_naf}", type: "C", description: "Code NAF de l'entreprise", example: "6201Z" },
      { placeholder: "${C_client_rcs}", type: "C", description: "RCS du client", example: "Paris B 123 456 789" },
      { placeholder: "${C_client_tva_intra}", type: "C", description: "Numéro de TVA intracommunautaire", example: "FR12345678901" },
      { placeholder: "${I_client_logo}", type: "I", description: "Logo du client", example: "[Image]" }
    ]
  },
  {
    id: "commercial",
    name: "Commercial",
    description: "Variables liées au commercial en charge du dossier",
    variables: [
      { placeholder: "${C_commercial_id}", type: "C", description: "Identifiant du commercial", example: "COM-001" },
      { placeholder: "${C_commercial_nom}", type: "C", description: "Nom du commercial", example: "Martin" },
      { placeholder: "${C_commercial_prenom}", type: "C", description: "Prénom du commercial", example: "Pierre" },
      { placeholder: "${C_commercial_email}", type: "C", description: "Email du commercial", example: "p.martin@societe.fr" },
      { placeholder: "${C_commercial_telephone}", type: "C", description: "Téléphone du commercial", example: "06 12 34 56 78" },
      { placeholder: "${C_commercial_agence}", type: "C", description: "Agence du commercial", example: "Paris Centre" },
      { placeholder: "${C_commercial_region}", type: "C", description: "Région du commercial", example: "Île-de-France" },
      { placeholder: "${I_commercial_signature}", type: "I", description: "Signature du commercial", example: "[Image]" }
    ]
  },
  {
    id: "fournisseur",
    name: "Fournisseur",
    description: "Variables liées au fournisseur du matériel",
    variables: [
      { placeholder: "${C_fournisseur_id}", type: "C", description: "Identifiant du fournisseur", example: "FOU-001" },
      { placeholder: "${C_fournisseur_raison_sociale}", type: "C", description: "Raison sociale du fournisseur", example: "Tech Solutions SARL" },
      { placeholder: "${C_fournisseur_siret}", type: "C", description: "SIRET du fournisseur", example: "987 654 321 00012" },
      { placeholder: "${C_fournisseur_adresse}", type: "C", description: "Adresse du fournisseur", example: "25 avenue des Technologies" },
      { placeholder: "${C_fournisseur_code_postal}", type: "C", description: "Code postal du fournisseur", example: "69001" },
      { placeholder: "${C_fournisseur_ville}", type: "C", description: "Ville du fournisseur", example: "Lyon" },
      { placeholder: "${C_fournisseur_pays}", type: "C", description: "Pays du fournisseur", example: "France" },
      { placeholder: "${C_fournisseur_telephone}", type: "C", description: "Téléphone du fournisseur", example: "04 56 78 90 12" },
      { placeholder: "${C_fournisseur_email}", type: "C", description: "Email du fournisseur", example: "contact@techsolutions.fr" },
      { placeholder: "${C_fournisseur_contact_nom}", type: "C", description: "Nom du contact fournisseur", example: "Durand" },
      { placeholder: "${C_fournisseur_contact_prenom}", type: "C", description: "Prénom du contact fournisseur", example: "Marie" },
      { placeholder: "${C_fournisseur_iban}", type: "C", description: "IBAN du fournisseur", example: "FR76 1234 5678 9012 3456 7890 123" },
      { placeholder: "${C_fournisseur_bic}", type: "C", description: "BIC du fournisseur", example: "BNPAFRPP" },
      { placeholder: "${I_fournisseur_logo}", type: "I", description: "Logo du fournisseur", example: "[Image]" }
    ]
  },
  {
    id: "societe_portage",
    name: "Société de Portage (SPV)",
    description: "Variables liées à la société de portage",
    variables: [
      { placeholder: "${C_spv_id}", type: "C", description: "Identifiant de la SPV", example: "SPV-001" },
      { placeholder: "${C_spv_raison_sociale}", type: "C", description: "Raison sociale de la SPV", example: "SPV Finance SAS" },
      { placeholder: "${C_spv_siret}", type: "C", description: "SIRET de la SPV", example: "111 222 333 00012" },
      { placeholder: "${C_spv_adresse}", type: "C", description: "Adresse de la SPV", example: "5 rue du Portage" },
      { placeholder: "${C_spv_code_postal}", type: "C", description: "Code postal de la SPV", example: "75008" },
      { placeholder: "${C_spv_ville}", type: "C", description: "Ville de la SPV", example: "Paris" },
      { placeholder: "${C_spv_capital}", type: "C", description: "Capital social de la SPV", example: "50 000,00 €" },
      { placeholder: "${C_spv_rcs}", type: "C", description: "RCS de la SPV", example: "Paris B 111 222 333" },
      { placeholder: "${C_spv_representant_nom}", type: "C", description: "Nom du représentant légal", example: "Bernard" },
      { placeholder: "${C_spv_representant_fonction}", type: "C", description: "Fonction du représentant", example: "Président" },
      { placeholder: "${I_spv_logo}", type: "I", description: "Logo de la SPV", example: "[Image]" }
    ]
  },
  {
    id: "produit",
    name: "Produit",
    description: "Variables liées aux produits financés",
    variables: [
      { placeholder: "${C_produit_id}", type: "C", description: "Identifiant du produit", example: "PRO-001" },
      { placeholder: "${C_produit_designation}", type: "C", description: "Désignation du produit", example: "Serveur Dell PowerEdge R750" },
      { placeholder: "${C_produit_reference}", type: "C", description: "Référence du produit", example: "DELL-R750-XS" },
      { placeholder: "${C_produit_categorie}", type: "C", description: "Catégorie du produit", example: "Matériel informatique" },
      { placeholder: "${C_produit_marque}", type: "C", description: "Marque du produit", example: "Dell" },
      { placeholder: "${C_produit_quantite}", type: "C", description: "Quantité commandée", example: "5" },
      { placeholder: "${C_produit_prix_unitaire_ht}", type: "C", description: "Prix unitaire HT", example: "8 500,00 €" },
      { placeholder: "${C_produit_prix_total_ht}", type: "C", description: "Prix total HT", example: "42 500,00 €" },
      { placeholder: "${C_produit_taux_tva}", type: "C", description: "Taux de TVA applicable", example: "20 %" },
      { placeholder: "${C_produit_numero_serie}", type: "C", description: "Numéro de série du produit", example: "SN-2024-ABC123" },
      { placeholder: "${C_produit_date_livraison}", type: "C", description: "Date de livraison du produit", example: "25/01/2024" },
      { placeholder: "${B_produit_neuf}", type: "B", description: "Indique si le produit est neuf", example: "Oui/Non" },
      { placeholder: "${T_produits_liste}", type: "T", description: "Tableau de tous les produits du dossier", example: "[Tableau]" }
    ]
  },
  {
    id: "produit_assurance",
    name: "Produit Assurance",
    description: "Variables liées à l'assurance des produits",
    variables: [
      { placeholder: "${C_assurance_produit_id}", type: "C", description: "Identifiant de l'assurance produit", example: "ASS-PRO-001" },
      { placeholder: "${C_assurance_produit_type}", type: "C", description: "Type d'assurance", example: "Tous risques" },
      { placeholder: "${C_assurance_produit_garantie}", type: "C", description: "Nature de la garantie", example: "Dommages matériels" },
      { placeholder: "${C_assurance_produit_montant}", type: "C", description: "Montant de la prime", example: "250,00 €" },
      { placeholder: "${C_assurance_produit_franchise}", type: "C", description: "Montant de la franchise", example: "500,00 €" },
      { placeholder: "${C_assurance_produit_date_debut}", type: "C", description: "Date de début de couverture", example: "01/02/2024" },
      { placeholder: "${C_assurance_produit_date_fin}", type: "C", description: "Date de fin de couverture", example: "31/01/2028" },
      { placeholder: "${B_assurance_produit_active}", type: "B", description: "Indique si l'assurance est active", example: "Oui/Non" }
    ]
  },
  {
    id: "caution",
    name: "Caution",
    description: "Variables liées aux cautions",
    variables: [
      { placeholder: "${C_caution_id}", type: "C", description: "Identifiant de la caution", example: "CAU-001" },
      { placeholder: "${C_caution_type}", type: "C", description: "Type de caution", example: "Caution personnelle" },
      { placeholder: "${C_caution_nom}", type: "C", description: "Nom de la caution", example: "Dupont" },
      { placeholder: "${C_caution_prenom}", type: "C", description: "Prénom de la caution", example: "Jean" },
      { placeholder: "${C_caution_adresse}", type: "C", description: "Adresse de la caution", example: "15 rue des Lilas" },
      { placeholder: "${C_caution_code_postal}", type: "C", description: "Code postal de la caution", example: "75015" },
      { placeholder: "${C_caution_ville}", type: "C", description: "Ville de la caution", example: "Paris" },
      { placeholder: "${C_caution_montant}", type: "C", description: "Montant de la caution", example: "50 000,00 €" },
      { placeholder: "${C_caution_date_naissance}", type: "C", description: "Date de naissance", example: "15/06/1970" },
      { placeholder: "${C_caution_lieu_naissance}", type: "C", description: "Lieu de naissance", example: "Lyon" },
      { placeholder: "${C_caution_nationalite}", type: "C", description: "Nationalité", example: "Française" },
      { placeholder: "${B_caution_solidaire}", type: "B", description: "Indique si la caution est solidaire", example: "Oui/Non" },
      { placeholder: "${T_cautions_liste}", type: "T", description: "Tableau de toutes les cautions", example: "[Tableau]" }
    ]
  },
  {
    id: "mandat_caution",
    name: "Mandat Caution",
    description: "Variables liées aux mandats de caution",
    variables: [
      { placeholder: "${C_mandat_caution_id}", type: "C", description: "Identifiant du mandat de caution", example: "MAN-CAU-001" },
      { placeholder: "${C_mandat_caution_reference}", type: "C", description: "Référence du mandat", example: "MC-2024-001" },
      { placeholder: "${C_mandat_caution_date}", type: "C", description: "Date du mandat", example: "18/01/2024" },
      { placeholder: "${C_mandat_caution_montant}", type: "C", description: "Montant garanti", example: "75 000,00 €" },
      { placeholder: "${C_mandat_caution_duree}", type: "C", description: "Durée du mandat", example: "48 mois" },
      { placeholder: "${C_mandat_caution_beneficiaire}", type: "C", description: "Bénéficiaire du mandat", example: "TD Finance" },
      { placeholder: "${B_mandat_caution_actif}", type: "B", description: "Indique si le mandat est actif", example: "Oui/Non" }
    ]
  },
  {
    id: "credit_fournisseur",
    name: "Crédit Fournisseur",
    description: "Variables liées au crédit fournisseur",
    variables: [
      { placeholder: "${C_credit_fournisseur_id}", type: "C", description: "Identifiant du crédit fournisseur", example: "CF-001" },
      { placeholder: "${C_credit_fournisseur_montant}", type: "C", description: "Montant du crédit fournisseur", example: "25 000,00 €" },
      { placeholder: "${C_credit_fournisseur_date_debut}", type: "C", description: "Date de début du crédit", example: "01/02/2024" },
      { placeholder: "${C_credit_fournisseur_date_fin}", type: "C", description: "Date de fin du crédit", example: "01/08/2024" },
      { placeholder: "${C_credit_fournisseur_taux}", type: "C", description: "Taux appliqué", example: "2,5 %" },
      { placeholder: "${C_credit_fournisseur_echeances}", type: "C", description: "Nombre d'échéances", example: "6" },
      { placeholder: "${B_credit_fournisseur_actif}", type: "B", description: "Indique si le crédit est actif", example: "Oui/Non" }
    ]
  },
  {
    id: "attestation_prix_net",
    name: "Attestation Prix Net Fournisseur",
    description: "Variables liées à l'attestation de prix net fournisseur",
    variables: [
      { placeholder: "${C_attestation_prix_net_id}", type: "C", description: "Identifiant de l'attestation", example: "APN-001" },
      { placeholder: "${C_attestation_prix_net_date}", type: "C", description: "Date de l'attestation", example: "15/01/2024" },
      { placeholder: "${C_attestation_prix_net_montant_ht}", type: "C", description: "Montant HT attesté", example: "100 000,00 €" },
      { placeholder: "${C_attestation_prix_net_montant_remise}", type: "C", description: "Montant de la remise", example: "5 000,00 €" },
      { placeholder: "${C_attestation_prix_net_montant_final}", type: "C", description: "Montant final après remise", example: "95 000,00 €" },
      { placeholder: "${C_attestation_prix_net_validite}", type: "C", description: "Durée de validité", example: "30 jours" },
      { placeholder: "${B_attestation_prix_net_valide}", type: "B", description: "Indique si l'attestation est valide", example: "Oui/Non" }
    ]
  },
  {
    id: "societe_assurance",
    name: "Société (Assurance)",
    description: "Variables liées à la société d'assurance",
    variables: [
      { placeholder: "${C_societe_assurance_id}", type: "C", description: "Identifiant de la société d'assurance", example: "SOC-ASS-001" },
      { placeholder: "${C_societe_assurance_nom}", type: "C", description: "Nom de la société d'assurance", example: "Allianz" },
      { placeholder: "${C_societe_assurance_adresse}", type: "C", description: "Adresse de la société", example: "1 cours Michelet" },
      { placeholder: "${C_societe_assurance_code_postal}", type: "C", description: "Code postal", example: "92800" },
      { placeholder: "${C_societe_assurance_ville}", type: "C", description: "Ville", example: "Puteaux" },
      { placeholder: "${C_societe_assurance_telephone}", type: "C", description: "Téléphone", example: "01 40 00 00 00" },
      { placeholder: "${C_societe_assurance_email}", type: "C", description: "Email de contact", example: "contact@allianz.fr" },
      { placeholder: "${C_societe_assurance_numero_agrement}", type: "C", description: "Numéro d'agrément", example: "AGR-123456" },
      { placeholder: "${I_societe_assurance_logo}", type: "I", description: "Logo de la société d'assurance", example: "[Image]" }
    ]
  },
  {
    id: "simulation",
    name: "Simulation",
    description: "Variables liées aux simulations de financement",
    variables: [
      { placeholder: "${C_simulation_id}", type: "C", description: "Identifiant de la simulation", example: "SIM-001" },
      { placeholder: "${C_simulation_date}", type: "C", description: "Date de la simulation", example: "10/01/2024" },
      { placeholder: "${C_simulation_montant}", type: "C", description: "Montant simulé", example: "100 000,00 €" },
      { placeholder: "${C_simulation_duree}", type: "C", description: "Durée de financement simulée", example: "48 mois" },
      { placeholder: "${C_simulation_taux}", type: "C", description: "Taux simulé", example: "3,25 %" },
      { placeholder: "${C_simulation_loyer}", type: "C", description: "Loyer mensuel simulé", example: "2 350,00 €" },
      { placeholder: "${C_simulation_premier_loyer}", type: "C", description: "Premier loyer majoré", example: "5 000,00 €" },
      { placeholder: "${C_simulation_valeur_residuelle}", type: "C", description: "Valeur résiduelle simulée", example: "10 000,00 €" },
      { placeholder: "${C_simulation_cout_total}", type: "C", description: "Coût total du financement", example: "117 800,00 €" },
      { placeholder: "${C_simulation_periodicite}", type: "C", description: "Périodicité des loyers", example: "Mensuel" },
      { placeholder: "${B_simulation_avec_assurance}", type: "B", description: "Simulation avec assurance", example: "Oui/Non" },
      { placeholder: "${T_simulation_echeancier}", type: "T", description: "Tableau de l'échéancier simulé", example: "[Tableau]" }
    ]
  }
];

/**
 * Obtenir toutes les variables aplaties
 */
export function getAllVariables() {
  const variables = [];
  categories.forEach(category => {
    category.variables.forEach(variable => {
      variables.push({
        ...variable,
        categoryId: category.id,
        categoryName: category.name
      });
    });
  });
  return variables;
}

/**
 * Rechercher des variables
 */
export function searchVariables(query) {
  const lowerQuery = query.toLowerCase();
  return getAllVariables().filter(variable =>
    variable.placeholder.toLowerCase().includes(lowerQuery) ||
    variable.description.toLowerCase().includes(lowerQuery) ||
    variable.categoryName.toLowerCase().includes(lowerQuery)
  );
}

/**
 * Obtenir les variables par catégorie
 */
export function getVariablesByCategory(categoryId) {
  const category = categories.find(c => c.id === categoryId);
  return category ? category.variables : [];
}

/**
 * Obtenir le nom du type
 */
export function getTypeName(type) {
  const types = {
    'C': 'Champ',
    'B': 'Booléen',
    'I': 'Image',
    'T': 'Tableau'
  };
  return types[type] || type;
}

/**
 * Obtenir la couleur du type
 */
export function getTypeColor(type) {
  const colors = {
    'C': '#0078d4', // Bleu - Champ
    'B': '#107c10', // Vert - Booléen
    'I': '#8764b8', // Violet - Image
    'T': '#d83b01'  // Orange - Tableau
  };
  return colors[type] || '#666666';
}
