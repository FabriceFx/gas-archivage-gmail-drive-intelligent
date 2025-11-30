/**
 * @fileoverview Script d'archivage intelligent Gmail -> Drive.
 * Fonctionnalités : Conversion PDF, Classement Chronologique, Renommage, Reporting.
 * @author Fabrice Faucheux
 * @version 5.0.0 (Finale)
 */

// --- Configuration Globale ---
const CONFIGURATION = {
  LIBELLE_GMAIL: "Facture",
  NOM_DOSSIER_RACINE: "Factures",
  EMAIL_ADMIN: "votre.email@domaine.com", // <--- Remplacer par votre email
  TYPES_ACCEPTES: ["application/pdf", "image/jpeg", "image/png"]
};

/**
 * Fonction principale : Orchestrateur.
 */
const sauvegarderFacturesVersDrive = () => {
  console.time("ExecutionTime");
  
  let rapportDIncidents = [];
  let messagesATraiter = [];

  try {
    // 1. Initialisation
    const dossierRacine = recupererOuCreerDossier(CONFIGURATION.NOM_DOSSIER_RACINE);
    if (!dossierRacine) throw new Error(`Racine introuvable : ${CONFIGURATION.NOM_DOSSIER_RACINE}`);

    // 2. Recherche
    const requete = `label:${CONFIGURATION.LIBELLE_GMAIL} is:unread`;
    const fils = GmailApp.search(requete);

    if (fils.length === 0) {
      console.info("Rien à signaler.");
      return;
    }

    console.info(`[INFO] Traitement de ${fils.length} fil(s).`);

    // 3. Boucle de traitement
    fils.forEach(fil => {
      const messages = fil.getMessages();

      messages.forEach(message => {
        if (message.isUnread()) {
          const piecesJointes = message.getAttachments();
          const dateMessage = message.getDate();
          const dossierCible = obtenirDossierChronologique(dossierRacine, dateMessage);
          
          // Récupération et nettoyage du nom de l'expéditeur
          const nomExpediteur = extraireNomExpediteur(message.getFrom());

          piecesJointes.forEach(piece => {
            if (CONFIGURATION.TYPES_ACCEPTES.includes(piece.getContentType())) {
              try {
                // A. Conversion/Standardisation en PDF
                const blobPdf = preparerBlobPdf(piece);
                
                // B. Calcul du nouveau nom standardisé
                // Format : YYYY-MM-DD_Expediteur_NomOriginal.pdf
                const prefixeDate = Utilities.formatDate(dateMessage, Session.getScriptTimeZone(), "yyyy-MM-dd");
                const nomOriginalNettoye = nettoyerNomFichier(blobPdf.getName());
                const nomFinal = `${prefixeDate}_${nomExpediteur}_${nomOriginalNettoye}`;
                
                // Application du nouveau nom
                blobPdf.setName(nomFinal);

                // C. Sauvegarde
                dossierCible.createFile(blobPdf);
                console.log(`[SUCCÈS] Sauvegardé : ${nomFinal}`);

              } catch (erreurFichier) {
                const contexte = `Fichier: ${piece.getName()} (Exp: ${nomExpediteur})`;
                console.error(`${contexte} - ${erreurFichier.message}`);
                rapportDIncidents.push({ contexte: contexte, message: erreurFichier.message });
              }
            }
          });
          
          messagesATraiter.push(message);
        }
      });
    });

    // 4. Marquage "Lu"
    if (messagesATraiter.length > 0) {
      GmailApp.markMessagesRead(messagesATraiter);
    }

  } catch (erreurCritique) {
    console.error(`[CRITIQUE] ${erreurCritique.message}`);
    rapportDIncidents.push({ contexte: "Arrêt critique", message: erreurCritique.message });
  } finally {
    if (rapportDIncidents.length > 0) envoyerNotificationAdmin(rapportDIncidents);
    console.timeEnd("ExecutionTime");
  }
};

/**
 * Extrait le nom "propre" de l'en-tête From.
 * Ex: "Amazon <auto@amazon.com>" devient "Amazon"
 */
const extraireNomExpediteur = (chaineFrom) => {
  // Regex pour capturer le texte avant les chevrons <email> ou entre guillemets
  const regexNom = /^"?([^"<]+)"?\s*<.*>$/;
  const match = chaineFrom.match(regexNom);
  
  let nom = match ? match[1].trim() : chaineFrom; // Si pas de match, on garde tout (ex: juste l'email)
  
  // Si le nom est un email, on prend ce qu'il y a avant le @
  if (nom.includes("@")) {
    nom = nom.split("@")[0];
  }
  
  return nettoyerNomFichier(nom); // On nettoie les caractères spéciaux
};

/**
 * Supprime les caractères interdits dans les noms de fichiers Windows/Unix/Drive
 * et remplace les espaces par des underscores pour la propreté.
 */
const nettoyerNomFichier = (texte) => {
  return texte
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // Enlève les accents
    .replace(/[^a-zA-Z0-9.\-_]/g, "_") // Remplace caractères spéciaux par _
    .replace(/_{2,}/g, "_") // Évite les doubles underscores __
    .trim(); 
};

/**
 * Convertit en PDF si nécessaire.
 */
const preparerBlobPdf = (pieceJointe) => {
  const typeMime = pieceJointe.getContentType();
  if (typeMime === "application/pdf") return pieceJointe.copyBlob();
  
  const blobConverti = pieceJointe.getAs("application/pdf");
  // On remplace l'extension originale par .pdf
  const nomSansExtension = pieceJointe.getName().replace(/\.[^/.]+$/, "");
  blobConverti.setName(nomSansExtension + ".pdf");
  return blobConverti;
};

// --- Fonctions Utilitaires Standards ---

const obtenirDossierChronologique = (dossierRacine, date) => {
  const timeZone = Session.getScriptTimeZone();
  const nomAnnee = Utilities.formatDate(date, timeZone, "yyyy");
  const nomMois = Utilities.formatDate(date, timeZone, "MM_MMMM"); 
  
  const dossierAnnee = recupererOuCreerSousDossier(dossierRacine, nomAnnee);
  return recupererOuCreerSousDossier(dossierAnnee, nomMois);
};

const recupererOuCreerSousDossier = (parent, nom) => {
  const iterateur = parent.getFoldersByName(nom);
  return iterateur.hasNext() ? iterateur.next() : parent.createFolder(nom);
};

const recupererOuCreerDossier = (nom) => {
  const iterateur = DriveApp.getFoldersByName(nom);
  return iterateur.hasNext() ? iterateur.next() : DriveApp.createFolder(nom);
};

const envoyerNotificationAdmin = (listeErreurs) => {
  const sujet = `⚠️ Alerte Script : Erreurs Archivage`;
  let corps = `<h3>Rapport d'incidents</h3><ul>` + 
              listeErreurs.map(e => `<li><b>${e.contexte}</b>: ${e.message}</li>`).join('') + 
              `</ul>`;
  GmailApp.sendEmail(CONFIGURATION.EMAIL_ADMIN, sujet, "HTML requis", { htmlBody: corps });
};
