const APP = {
  SHEETS: {
    AGENTS: 'AGENTS',
    USERS: 'USERS',
    PARAMS: 'PARAMETRES',
    THEME_MODELS: 'MODELES_THEMES',
    TRACKING: 'SUIVI_EMARGEMENT',
    QUIZ_MODELES: 'QUIZ_MODELES',
    QUIZ_SESSIONS: 'QUIZ_SESSIONS',
    QUIZ_REPONSES: 'QUIZ_REPONSES'
  },
  PARAM_KEYS: {
    TEMPLATE_ID: 'DOC_TEMPLATE_ID',
    OUTPUT_FOLDER_ID: 'OUTPUT_FOLDER_ID',
    DEFAULT_TITLE: 'DEFAULT_INTITULE_SESSION',
    DEFAULT_DURATION: 'DEFAULT_DUREE_SESSION',
    DEFAULT_LOCATION: 'DEFAULT_LIEU',
    APP_NAME: 'APP_NAME',
    QUIZ_TEMPLATE_ID: 'QUIZ_TEMPLATE_ID'
  },
  STATUS: {
    PDF_GENERATED: 'Généré',
    PDF_NOT_GENERATED: 'Non généré',
    SIGNED: 'Signé',
    PENDING: 'À signer'
  },
  QUIZ_STATUS: {
    DRAFT: 'Brouillon',
    ACTIVE: 'Active',
    CLOSED: 'Fermée'
  },
  QUIZ_REPONSE_STATUS: {
    SUBMITTED: 'Soumis',
    GRADED: 'Noté',
    PDF_GENERATED: 'PDF généré'
  },
  MARKERS: {
    INTERVENANT: '[[INTERVENANT_SIGNATURE_BLOCK]]',
    AGENT: '[[AGENT_SIGNATURE_BLOCK]]'
  },
  SIGNATURE_TITLE_DEFAULT: 'Team leader',
  SCRIPT_PROPS: {
    SPREADSHEET_ID: 'APP_SPREADSHEET_ID'
  }
};



function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Émargement')
    .addItem('Initialiser les feuilles', 'setupWorkbook')
    .addItem('Créer un modèle Google Doc par défaut', 'createDefaultTemplateDoc')
    .addItem('Générer un hash de mot de passe', 'promptPasswordHash')
    .addToUi();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Émargement mensuel')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupWorkbook() {
  rememberActiveSpreadsheet_();
  initializeWorkbook_();
  SpreadsheetApp.getUi().alert(
    'Initialisation terminée.\n\n' +
      '1) Renseigne AGENTS\n' +
      '2) Crée ton premier utilisateur dans l’onglet USERS ou depuis l’interface\n' +
      '3) Convertis le modèle en Google Docs\n' +
      '4) Renseigne DOC_TEMPLATE_ID et OUTPUT_FOLDER_ID\n' +
      '5) Déploie le script en application web'
  );
}

function appInit() {
  rememberActiveSpreadsheet_();
  initializeWorkbook_();
  const params = getParamsMap_();
  return {
    appName: params[APP.PARAM_KEYS.APP_NAME] || 'Émargement mensuel',
    hasTemplate: !!params[APP.PARAM_KEYS.TEMPLATE_ID],
    hasOutputFolder: !!params[APP.PARAM_KEYS.OUTPUT_FOLDER_ID]
  };
}

function rememberActiveSpreadsheet_() {
  try {
    const active = SpreadsheetApp.getActive();
    if (active && active.getId) {
      PropertiesService.getScriptProperties().setProperty(APP.SCRIPT_PROPS.SPREADSHEET_ID, active.getId());
      return active;
    }
  } catch (err) { console.warn('rememberActiveSpreadsheet_:', err); }
  return null;
}

function getAppSpreadsheet_() {
  try {
    const active = SpreadsheetApp.getActive();
    if (active && active.getId && active.getSheetByName(APP.SHEETS.AGENTS)) {
      PropertiesService.getScriptProperties().setProperty(APP.SCRIPT_PROPS.SPREADSHEET_ID, active.getId());
      return active;
    }
  } catch (err) { console.warn('getAppSpreadsheet_ active lookup:', err); }

  const storedId = PropertiesService.getScriptProperties().getProperty(APP.SCRIPT_PROPS.SPREADSHEET_ID);
  if (storedId) {
    return SpreadsheetApp.openById(storedId);
  }

  const fallback = rememberActiveSpreadsheet_();
  if (fallback) return fallback;
  throw new Error('Classeur source introuvable. Lance setupWorkbook() depuis le Google Sheet lié, puis recharge l’application.');
}


function promptPasswordHash() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Générer un hash', 'Entre le mot de passe à hasher :', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const hash = sha256Hex_(response.getResponseText());
  ui.alert('Hash SHA-256', hash, ui.ButtonSet.OK);
}

function createDefaultTemplateDoc() {
  const doc = DocumentApp.create('MODELE - PROD-ENR-004C - Feuille de présence');
  const body = doc.getBody();
  body.clear();

  // Marges réduites pour plus d'espace
  body
    .setMarginTop(36)
    .setMarginBottom(36)
    .setMarginLeft(54)
    .setMarginRight(54);

  // En-tête avec logo/titre
  const header = body.appendTable([
    ['odity', 'FEUILLE DE PRÉSENCE / ÉMARGEMENT', 'Émetteur : Production\nDate : 09/01/2020\nRéférence : PROD-ENR-004C']
  ]);
  header.setBorderWidth(0);

  const hRow = header.getRow(0);
  const hCell0 = hRow.getCell(0);
  const hCell1 = hRow.getCell(1);
  const hCell2 = hRow.getCell(2);

  hCell0.editAsText()
    .setBold(true)
    .setFontSize(18)
    .setForegroundColor('#1a1a2e');

  hCell1.editAsText()
    .setBold(true)
    .setFontSize(13)
    .setForegroundColor('#1a1a2e');
  hCell1.setWidth(300);

  hCell2.editAsText()
    .setFontSize(8)
    .setForegroundColor('#555555');

  hRow.setMinimumHeight(40);

  body.appendParagraph('').setSpacingAfter(4);

  // Infos session
  appendLabeledLine_(body, 'Intitulé de session/stage : ', '{{INTITULE_SESSION}}');
  appendLabeledLine_(body, 'Date de la session/stage : ', '{{DATE_SESSION}}');
  appendLabeledLine_(body, 'Durée de la session/stage : ', '{{DUREE_SESSION}}');
  appendLabeledLine_(body, 'Lieu : ', '{{LIEU}}');

  body.appendParagraph('').setSpacingAfter(4);

  // Tableau intervenant (ADMIN/USER) — signature manager
  const tableInter = body.appendTable([
    ['Intervenants', 'Signature intervenant (Team leader)'],
    ['{{INTERVENANT}}', APP.MARKERS.INTERVENANT]
  ]);
  tableInter.setBorderWidth(1);
  tableInter.setBorderColor('#cccccc');

  const interHeaderRow = tableInter.getRow(0);
  const interCell0 = interHeaderRow.getCell(0);
  const interCell1 = interHeaderRow.getCell(1);

  interCell0.setBackgroundColor('#f0f0f8');
  interCell1.setBackgroundColor('#f0f0f8');

  interCell0.editAsText().setBold(true).setFontSize(10);
  interCell1.editAsText().setBold(true).setFontSize(10);

  interHeaderRow.setMinimumHeight(28);
  tableInter.getRow(1).setMinimumHeight(90);

  body.appendParagraph('').setSpacingAfter(4);

  // Thèmes abordés
  const themesTitle = body.appendParagraph('Intitulés des thèmes/sujets abordés :');
  themesTitle.editAsText().setBold(true).setFontSize(10);

  const themesContent = body.appendParagraph('{{THEMES_ABORDES}}');
  themesContent.editAsText().setFontSize(10);

  body.appendParagraph('').setSpacingAfter(4);

  // Tableau agents — signature agent / ambassadeur
  const tableAgents = body.appendTable([
    ['#', 'Matricule', 'Nom & Prénom', 'Sous-équipe', 'Signature ambassadeur'],
    ['1', '{{AGENT_MATRICULE}}', '{{AGENT_NOM_PRENOM}}', '{{SOUS_EQUIPE}}', '{{SIGNATURE}}'],
    ['2', '', '', '', ''],
    ['3', '', '', '', '']
  ]);
  tableAgents.setBorderWidth(1);
  tableAgents.setBorderColor('#cccccc');

  const agHeader = tableAgents.getRow(0);
  for (let c = 0; c < 5; c++) {
    const cell = agHeader.getCell(c);
    cell.setBackgroundColor('#f0f0f8');
    cell.editAsText().setBold(true).setFontSize(10);
  }

  agHeader.setMinimumHeight(28);

  // Largeurs de colonnes
  tableAgents.getRow(0).getCell(0).setWidth(28);   // #
  tableAgents.getRow(0).getCell(1).setWidth(72);   // Matricule
  tableAgents.getRow(0).getCell(2).setWidth(130);  // Nom & Prénom
  tableAgents.getRow(0).getCell(3).setWidth(90);   // Sous-équipe
  tableAgents.getRow(0).getCell(4).setWidth(160);  // Signature

  // Hauteur minimale pour les lignes de signature
  for (let r = 1; r <= 3; r++) {
    tableAgents.getRow(r).setMinimumHeight(75);
  }

  body.appendParagraph('').setSpacingAfter(4);

  // Remarques
  const remarksTitle = body.appendParagraph('Remarques des intervenants :');
  remarksTitle.editAsText().setBold(true).setFontSize(10);

  const remarksContent = body.appendParagraph('{{REMARQUES}}');
  remarksContent.editAsText().setFontSize(10);

  doc.saveAndClose();

  upsertParam_(
    getAppSpreadsheet_().getSheetByName(APP.SHEETS.PARAMS),
    APP.PARAM_KEYS.TEMPLATE_ID,
    doc.getId()
  );

  SpreadsheetApp.getUi().alert(
    'Modèle Google Docs créé avec succès.\n\n' +
    'ID : ' + doc.getId() + '\n\n' +
    'Note : colonnes Signature = "Team leader" pour admin/intervenant, "Ambassadeur" pour agent.'
  );
}

function login(userId, password) {
  setupWorkbookSilently_();
  const users = getUsers_();
  const cleanUserId = String(userId || '').trim();
  const user = users.find(u => u.userId === cleanUserId && u.active);
  if (!user) throw new Error('Identifiant inconnu ou inactif.');

  const hash = sha256Hex_(String(password || ''));
  if (hash !== user.passwordHash) throw new Error('Mot de passe incorrect.');

  const token = Utilities.getUuid();
  const payload = {
    userId: user.userId,
    name: user.name,
    role: user.role,
    title: user.title || APP.SIGNATURE_TITLE_DEFAULT,
    agentMatricule: user.agentMatricule || '',
    expiresAt: Date.now() + 8 * 60 * 60 * 1000
  };
  PropertiesService.getScriptProperties().setProperty('SESSION_' + token, JSON.stringify(payload));

  return {
    token,
    user: {
      userId: user.userId,
      name: user.name,
      role: user.role,
      title: user.title || APP.SIGNATURE_TITLE_DEFAULT,
      agentMatricule: user.agentMatricule || ''
    }
  };
}

function logout(token) {
  if (token) PropertiesService.getScriptProperties().deleteProperty('SESSION_' + token);
  return { success: true };
}

function getBootstrapData(token) {
  const session = requireSession_(token);
  const params = getParamsMap_();
  return {
    currentUser: session,
    settings: {
      templateId: params[APP.PARAM_KEYS.TEMPLATE_ID] || '',
      outputFolderId: params[APP.PARAM_KEYS.OUTPUT_FOLDER_ID] || '',
      defaultTitle: params[APP.PARAM_KEYS.DEFAULT_TITLE] || '',
      defaultDuration: params[APP.PARAM_KEYS.DEFAULT_DURATION] || '',
      defaultLocation: params[APP.PARAM_KEYS.DEFAULT_LOCATION] || '',
      appName: params[APP.PARAM_KEYS.APP_NAME] || 'Émargement mensuel'
    },
    agents: getAgents_(),
    users: isAdmin_(session) ? getUsers_() : [],
    themeModels: getThemeModels_(),
    currentPeriod: periodFromDate_(new Date())
  };
}

function getDashboard(token, period) {
  const session = requireSession_(token);
  const cleanPeriod = normalizePeriod_(period || periodFromDate_(new Date()));
  const records = getTrackingRowsByPeriodFiltered_(session, cleanPeriod)
    .sort((a, b) => {
      const nameA = String(a.agentName || '').toLowerCase();
      const nameB = String(b.agentName || '').toLowerCase();
      return nameA.localeCompare(nameB, 'fr');
    });

  const rows = records.map(rec => ({
    matricule: rec.agentMatricule,
    name: rec.agentName,
    prenom: rec.agentPrenom,
    sousEquipe: rec.sousEquipe,
    statutPdf: rec.statutPdf || APP.STATUS.PDF_NOT_GENERATED,
    statutSignatureIntervenant: rec.statutSignatureIntervenant || APP.STATUS.PENDING,
    statutSignatureAgent: rec.statutSignatureAgent || APP.STATUS.PENDING,
    dateSession: formatMaybeDateFr_(rec.dateSession),
    intervenant: String(rec.intervenantNom || ''),
    pdfUrl: String(rec.pdfUrl || ''),
    sourceDocUrl: String(rec.sourceDocUrl || ''),
    rowNumber: rec.rowNumber || null,
    canSignIntervenant: canSessionSignIntervenant_(session, rec),
    canSignAgent: canSessionSignAgent_(session, rec),
    signatureIntervenantAt: formatMaybeDateTimeFr_(rec.dateSignatureIntervenant),
    signatureAgentAt: formatMaybeDateTimeFr_(rec.dateSignatureAgent)
  }));

  return { period: cleanPeriod, rows };
}

function getSignatureQueue(token, period) {
  const session = requireSession_(token);
  const cleanPeriod = normalizePeriod_(period || periodFromDate_(new Date()));
  let records = getTrackingRowsByPeriodFiltered_(session, cleanPeriod);

  if (isAdmin_(session)) {
    records = records.filter(r => r.rowNumber);
  } else if (isAgent_(session)) {
    records = records.filter(r => r.rowNumber && isSameAgentForSession_(session, r.agentMatricule));
  } else {
    records = records.filter(r => r.rowNumber && r.intervenantUserId === session.userId);
  }

  const rows = records.map(rec => ({
    matricule: rec.agentMatricule,
    name: rec.agentName,
    prenom: rec.agentPrenom,
    sousEquipe: rec.sousEquipe,
    statutPdf: rec.statutPdf || APP.STATUS.PDF_NOT_GENERATED,
    statutSignatureIntervenant: rec.statutSignatureIntervenant || APP.STATUS.PENDING,
    statutSignatureAgent: rec.statutSignatureAgent || APP.STATUS.PENDING,
    dateSession: formatMaybeDateFr_(rec.dateSession),
    intervenant: String(rec.intervenantNom || ''),
    pdfUrl: String(rec.pdfUrl || ''),
    sourceDocUrl: String(rec.sourceDocUrl || ''),
    rowNumber: rec.rowNumber || null,
    canSignIntervenant: canSessionSignIntervenant_(session, rec),
    canSignAgent: canSessionSignAgent_(session, rec),
    signatureIntervenantAt: formatMaybeDateTimeFr_(rec.dateSignatureIntervenant),
    signatureAgentAt: formatMaybeDateTimeFr_(rec.dateSignatureAgent)
  }));

  return { period: cleanPeriod, rows };
}

function saveSettings(token, payload) {
  const session = requireSession_(token);
  if (!isAdmin_(session)) throw new Error('Accès réservé à un profil admin.');

  const ss = getAppSpreadsheet_();
  const paramsSheet = ss.getSheetByName(APP.SHEETS.PARAMS);

  upsertParam_(paramsSheet, APP.PARAM_KEYS.TEMPLATE_ID, String(payload.templateId || '').trim());
  upsertParam_(paramsSheet, APP.PARAM_KEYS.OUTPUT_FOLDER_ID, String(payload.outputFolderId || '').trim());
  upsertParam_(paramsSheet, APP.PARAM_KEYS.DEFAULT_TITLE, String(payload.defaultTitle || '').trim());
  upsertParam_(paramsSheet, APP.PARAM_KEYS.DEFAULT_DURATION, String(payload.defaultDuration || '').trim());
  upsertParam_(paramsSheet, APP.PARAM_KEYS.DEFAULT_LOCATION, String(payload.defaultLocation || '').trim());
  upsertParam_(paramsSheet, APP.PARAM_KEYS.APP_NAME, String(payload.appName || 'Émargement mensuel').trim());

  if (Array.isArray(payload.themeModels)) {
    const sheet = ss.getSheetByName(APP.SHEETS.THEME_MODELS);
    clearDataKeepHeader_(sheet);
    const rows = payload.themeModels
      .filter(item => item && String(item.name || '').trim())
      .map(item => [
        String(item.name || '').trim(),
        String(item.content || '').trim(),
        item.active === false ? 'NON' : 'OUI'
      ]);
    if (rows.length) sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  return { success: true };
}

function saveUserAdmin(token, payload) {
  const session = requireSession_(token);
  if (!isAdmin_(session)) throw new Error('Accès réservé à un profil admin.');

  const userId = String(payload.userId || '').trim();
  const name = String(payload.name || '').trim();
  const role = String(payload.role || 'USER').trim().toUpperCase();
  const title = String(payload.title || APP.SIGNATURE_TITLE_DEFAULT).trim();
  const active = payload.active === false ? 'NON' : 'OUI';
  const password = String(payload.password || '');
  const agentMatricule = String(payload.agentMatricule || '').trim();

  if (!userId) throw new Error('UserID obligatoire.');
  if (!name) throw new Error('Nom / prénom obligatoire.');
  if (role === 'AGENT' && !agentMatricule) throw new Error('Matricule agent obligatoire pour un profil AGENT.');

  const sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.USERS);
  const data = getDataRows_(sheet);
  let rowNumber = null;
  let existingHash = '';
  data.forEach((row, i) => {
    if (String(row[0] || '').trim() === userId) {
      rowNumber = i + 2;
      existingHash = String(row[2] || '').trim();
    }
  });

  const passwordHash = password ? sha256Hex_(password) : existingHash;
  if (!passwordHash) throw new Error('Mot de passe obligatoire pour créer un utilisateur.');

  const row = [userId, name, passwordHash, role, title, active, role === 'AGENT' ? agentMatricule : ''];
  if (rowNumber) {
    sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { success: true };
}

function deleteUserAdmin(token, userId) {
  const session = requireSession_(token);
  if (!isAdmin_(session)) throw new Error('Accès réservé à un profil admin.');

  const cleanUserId = String(userId || '').trim();
  if (!cleanUserId) throw new Error('UserID manquant.');
  if (cleanUserId === session.userId) throw new Error('Tu ne peux pas supprimer ton propre compte.');

  const sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.USERS);
  const data = getDataRows_(sheet);
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === cleanUserId) {
      sheet.deleteRow(i + 2);
      return { success: true };
    }
  }
  throw new Error('Utilisateur introuvable.');
}

function generateAttendancePdfs(token, payload) {
  const session = requireSession_(token);
  if (isAgent_(session)) throw new Error('Un agent ne peut pas générer des fiches.');
  const params = getParamsMap_();
  const templateId = String(params[APP.PARAM_KEYS.TEMPLATE_ID] || '').trim();
  const outputFolderId = String(params[APP.PARAM_KEYS.OUTPUT_FOLDER_ID] || '').trim();

  if (!templateId) throw new Error('Le paramètre DOC_TEMPLATE_ID est vide.');
  if (!outputFolderId) throw new Error('Le paramètre OUTPUT_FOLDER_ID est vide.');

  const selectedMatricules = Array.isArray(payload.selectedMatricules) ? payload.selectedMatricules : [];
  if (!selectedMatricules.length) throw new Error('Aucun agent sélectionné.');

  const sessionDate = parseLocalDate_(payload.sessionDate);
  const period = periodFromDate_(sessionDate);
  const monthLabel = monthLabelFr_(sessionDate);
  const trackingSheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.TRACKING);
  const trackingIndex = getTrackingIndexByPeriod_(period);
  const agentsByMatricule = getAgents_().reduce((acc, item) => (acc[item.matricule] = item, acc), {});
  const outputFolder = DriveApp.getFolderById(outputFolderId);
  const periodFolder = getOrCreateSubFolder_(outputFolder, period + ' - ' + monthLabel);
  const templateFile = DriveApp.getFileById(templateId);

  if (templateFile.getMimeType() !== MimeType.GOOGLE_DOCS) {
    throw new Error("Le modèle doit être un Google Docs natif. Ouvre le .docx puis fais 'Fichier > Enregistrer comme Google Docs', puis renseigne le nouvel ID dans PARAMETRES > DOC_TEMPLATE_ID.");
  }

  const agentThemesMap = (payload.agentThemesMap && typeof payload.agentThemesMap === 'object') ? payload.agentThemesMap : {};
  const results = [];

  selectedMatricules.forEach(matricule => {
    const agent = agentsByMatricule[matricule];
    if (!agent) {
      results.push({ matricule, success: false, message: 'Agent introuvable.' });
      return;
    }

    try {
      const sessionDateStr = Utilities.formatDate(sessionDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const existing = trackingIndex[matricule + '|' + sessionDateStr] || null;
      const pdfName = buildPdfName_(agent.prenom || agent.fullName, sessionDate);
      const sourceDocName = buildSourceDocName_(agent.fullName, sessionDate);

      // Nettoyage des doublons orphelins par nom (safe avant makeCopy)
      trashFilesByNameInFolder_(periodFolder, pdfName);
      trashFilesByNameInFolder_(periodFolder, sourceDocName);

      const docCopy = templateFile.makeCopy(sourceDocName, periodFolder);
      const doc = openDocumentWithRetry_(docCopy.getId(), 5);
      const body = doc.getBody();

      // Thèmes spécifiques par agent (import CSV) ou thèmes communs
      const agentThemes = String(agentThemesMap[matricule] || payload.themesText || '');

      replaceTextSafely_(body, '{{INTITULE_SESSION}}', String(payload.intituleSession || ''));
      replaceTextSafely_(body, '{{DATE_SESSION}}', formatDateFr_(sessionDate));
      replaceTextSafely_(body, '{{DUREE_SESSION}}', String(payload.duree || ''));
      replaceTextSafely_(body, '{{LIEU}}', String(payload.lieu || ''));
      replaceTextSafely_(body, '{{THEMES_ABORDES}}', agentThemes);
      replaceTextSafely_(body, '{{AGENT_MATRICULE}}', agent.matricule);
      replaceTextSafely_(body, '{{AGENT_NOM_PRENOM}}', agent.fullName);
      replaceTextSafely_(body, '{{SOUS_EQUIPE}}', agent.sousEquipe);
      replaceTextSafely_(body, '{{REMARQUES}}', String(payload.remarques || ''));

      prepareIntervenantLayout_(doc, session.userId);
      prepareAgentSignatureAnchor_(doc);

      const record = {
        recordId: existing ? existing.recordId : Utilities.getUuid(),
        period: period,
        dateSession: Utilities.formatDate(sessionDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        monthLabel: monthLabel,
        agentMatricule: agent.matricule,
        agentName: agent.fullName,
        agentPrenom: agent.prenom,
        sousEquipe: agent.sousEquipe,
        intituleSession: String(payload.intituleSession || ''),
        duree: String(payload.duree || ''),
        lieu: String(payload.lieu || ''),
        intervenantUserId: session.userId,
        intervenantNom: cleanIntervenantLabel_(session.name),
        intervenantTitle: session.title || APP.SIGNATURE_TITLE_DEFAULT,
        themeModel: String(payload.themeModelName || ''),
        themesText: agentThemes,
        remarques: String(payload.remarques || ''),
        sourceDocFileId: docCopy.getId(),
        sourceDocUrl: docCopy.getUrl(),
        pdfFileId: '',
        pdfUrl: '',
        statutPdf: APP.STATUS.PDF_NOT_GENERATED,
        statutSignatureIntervenant: existing ? (existing.statutSignatureIntervenant || APP.STATUS.PENDING) : APP.STATUS.PENDING,
        dateSignatureIntervenant: existing ? existing.dateSignatureIntervenant : '',
        signeIntervenantPar: existing ? existing.signeIntervenantPar : '',
        statutSignatureAgent: existing ? (existing.statutSignatureAgent || APP.STATUS.PENDING) : APP.STATUS.PENDING,
        dateSignatureAgent: existing ? existing.dateSignatureAgent : '',
        signeAgentPar: existing ? existing.signeAgentPar : '',
        dateGeneration: new Date(),
        generePar: session.name,
        majLe: new Date()
      };

      renderSignatureBlocks_(doc, record);
      doc.saveAndClose();

      const pdfFile = regeneratePdfFromSource_(record, periodFolder, pdfName);
      record.pdfFileId = pdfFile.getId();
      record.pdfUrl = pdfFile.getUrl();
      record.statutPdf = APP.STATUS.PDF_GENERATED;

      // Suppression des anciens fichiers APRÈS génération réussie (évite la perte de données)
      if (existing && existing.pdfFileId) trashFileIfExists_(existing.pdfFileId);
      if (existing && existing.sourceDocFileId) trashFileIfExists_(existing.sourceDocFileId);

      upsertTrackingRecord_(trackingSheet, existing ? existing.rowNumber : null, record);
      SpreadsheetApp.flush();

      results.push({
        matricule: agent.matricule,
        agent: agent.fullName,
        success: true,
        pdfUrl: pdfFile.getUrl(),
        period: period,
        message: existing ? 'PDF régénéré et ancien fichier remplacé.' : 'PDF généré.'
      });
    } catch (err) {
      results.push({
        matricule: agent.matricule,
        agent: agent.fullName,
        success: false,
        period: period,
        message: err.message || String(err)
      });
    }
  });

  return { results: results, period: period };
}


function applySignatureToRecord_(record, signatureType, session) {
  if (signatureType === 'intervenant') {
    if (!canSessionSignIntervenant_(session, record)) {
      throw new Error('Tu ne peux pas signer cette fiche comme intervenant.');
    }
    record.statutSignatureIntervenant = APP.STATUS.SIGNED;
    record.dateSignatureIntervenant = new Date();
    record.signeIntervenantPar = session.name + ' (' + session.userId + ')';
    return record;
  }

  if (signatureType === 'agent') {
    if (!canSessionSignAgent_(session, record)) {
      throw new Error('Tu ne peux pas signer cette fiche comme agent.');
    }
    record.statutSignatureAgent = APP.STATUS.SIGNED;
    record.dateSignatureAgent = new Date();
    record.signeAgentPar = session.name + ' (' + session.userId + ')';
    return record;
  }

  throw new Error('Type de signature invalide.');
}

function signRecord(token, payload) {
  const session = requireSession_(token);
  const signatureType = String(payload.signatureType || '').trim();
  const password = String(payload.password || '');
  const rowNumber = Number(payload.rowNumber || 0);

  if (!rowNumber || rowNumber < 2) throw new Error('Ligne invalide.');
  if (!password) throw new Error('Mot de passe requis pour signer.');

  validateSessionPassword_(session.userId, password);

  const sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.TRACKING);
  if (rowNumber > sheet.getLastRow()) throw new Error('Ligne hors limites.');
  const record = getTrackingRecordByRow_(rowNumber);
  if (!record || !record.recordId) throw new Error('Enregistrement introuvable.');
  if (!record.sourceDocFileId) throw new Error('Document source introuvable.');
  if (!record.pdfFileId && !record.sourceDocFileId) throw new Error('Aucun document associé.');

  applySignatureToRecord_(record, signatureType, session);

  const doc = openDocumentWithRetry_(record.sourceDocFileId, 5);
  renderSignatureBlocks_(doc, record);
  doc.saveAndClose();

  const folder = DriveApp.getFileById(record.sourceDocFileId).getParents().hasNext()
    ? DriveApp.getFileById(record.sourceDocFileId).getParents().next()
    : DriveApp.getFolderById(String(getParamsMap_()[APP.PARAM_KEYS.OUTPUT_FOLDER_ID] || ''));

  const pdfName = buildPdfName_(record.agentPrenom || extractFirstName_(record.agentName), parseLocalDate_(record.dateSession));
  if (record.pdfFileId) trashFileIfExists_(record.pdfFileId);
  trashFilesByNameInFolder_(folder, pdfName);
  const pdfFile = regeneratePdfFromSource_(record, folder, pdfName);

  record.pdfFileId = pdfFile.getId();
  record.pdfUrl = pdfFile.getUrl();
  record.statutPdf = APP.STATUS.PDF_GENERATED;
  record.majLe = new Date();

  upsertTrackingRecord_(sheet, rowNumber, record);
  SpreadsheetApp.flush();

  return { success: true, pdfUrl: pdfFile.getUrl() };
}



function bulkSignRecords(token, payload) {
  const session = requireSession_(token);
  const password = String((payload && payload.password) || '');
  const signatureType = String((payload && payload.signatureType) || '').trim();
  const rowNumbers = Array.isArray(payload && payload.rowNumbers) ? payload.rowNumbers.map(n => Number(n || 0)).filter(n => n >= 2) : [];

  if (!password) throw new Error('Mot de passe requis.');
  if (!rowNumbers.length) throw new Error('Aucune fiche sélectionnée.');
  if (rowNumbers.length > 50) throw new Error('Maximum 50 fiches par lot. Vous en avez sélectionné ' + rowNumbers.length + '.');
  if (!signatureType) throw new Error('Type de signature requis.');

  validateSessionPassword_(session.userId, password);

  const sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.TRACKING);
  const updated = [];
  const skipped = [];

  rowNumbers.forEach(rowNumber => {
    try {
      const record = getTrackingRecordByRow_(rowNumber);
      if (!record || !record.recordId || !record.sourceDocFileId) {
        skipped.push({ rowNumber: rowNumber, reason: 'Fiche introuvable ou document source manquant.' });
        return;
      }

      applySignatureToRecord_(record, signatureType, session);

      const doc = openDocumentWithRetry_(record.sourceDocFileId, 5);
      renderSignatureBlocks_(doc, record);
      doc.saveAndClose();

      const folder = DriveApp.getFileById(record.sourceDocFileId).getParents().hasNext()
        ? DriveApp.getFileById(record.sourceDocFileId).getParents().next()
        : DriveApp.getFolderById(String(getParamsMap_()[APP.PARAM_KEYS.OUTPUT_FOLDER_ID] || ''));

      const pdfName = buildPdfName_(record.agentPrenom || extractFirstName_(record.agentName), parseLocalDate_(record.dateSession));
      if (record.pdfFileId) trashFileIfExists_(record.pdfFileId);
      trashFilesByNameInFolder_(folder, pdfName);
      const pdfFile = regeneratePdfFromSource_(record, folder, pdfName);

      record.pdfFileId = pdfFile.getId();
      record.pdfUrl = pdfFile.getUrl();
      record.statutPdf = APP.STATUS.PDF_GENERATED;
      record.majLe = new Date();

      upsertTrackingRecord_(sheet, rowNumber, record);
      updated.push({ rowNumber: rowNumber, pdfUrl: pdfFile.getUrl() });
    } catch (err) {
      skipped.push({ rowNumber: rowNumber, reason: err.message || String(err) });
    }
  });

  SpreadsheetApp.flush();
  return { success: true, updatedCount: updated.length, skipped: skipped, updated: updated };
}

function validateSessionPassword_(userId, password) {
  const users = getUsers_();
  const user = users.find(u => u.userId === String(userId || '').trim() && u.active);
  if (!user) throw new Error('Utilisateur introuvable.');
  if (sha256Hex_(String(password || '')) !== user.passwordHash) {
    throw new Error('Mot de passe incorrect.');
  }
  return true;
}

function regeneratePdfFromSource_(record, folder, pdfName) {
  const sourceDocId = String(record.sourceDocFileId || '').trim();
  if (!sourceDocId) {
    throw new Error('Document source introuvable pour la régénération du PDF.');
  }

  let lastErr = null;

  for (let attempt = 0; attempt < 5; attempt++) {
    try {
      // Laisse le temps à Google Docs de persister l'image avant conversion PDF
      Utilities.sleep(attempt === 0 ? 1800 : 1200 + attempt * 700);

      const sourceFile = DriveApp.getFileById(sourceDocId);
      const pdfBlob = sourceFile.getAs(MimeType.PDF).setName(pdfName);

      if (!pdfBlob || pdfBlob.getBytes().length === 0) {
        throw new Error('PDF vide après conversion.');
      }

      const pdfFile = folder.createFile(pdfBlob);
      pdfFile.setDescription('Feuille de présence - ' + record.agentName + ' - ' + record.monthLabel);
      return pdfFile;
    } catch (err) {
      lastErr = err;
    }
  }

  throw new Error(
    'Impossible de régénérer le PDF avec la signature : ' +
    (lastErr && lastErr.message ? lastErr.message : String(lastErr))
  );
}

function renderSignatureBlocks_(doc, record) {
  renderIntervenantBlock_(doc, record);
  renderAgentBlock_(doc, record);
}

function renderIntervenantBlock_(doc, record) {
  const body = doc.getBody();
  const anchors = findIntervenantCells_(body);
  if (!anchors || !anchors.labelCell) return;

  const labelCell = anchors.labelCell;
  const signatureCell = anchors.signatureCell || anchors.labelCell;

  clearTableCell_(labelCell);
  const p = labelCell.appendParagraph(String(record.intervenantUserId || '').trim());
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  p.setSpacingBefore(6);
  p.setSpacingAfter(0);
  const t = p.editAsText();
  t.setFontSize(10);
  t.setBold(true);

  if (signatureCell && signatureCell !== labelCell) {
    clearTableCell_(signatureCell);
    if (record.statutSignatureIntervenant === APP.STATUS.SIGNED) {
      const inserted = insertSignatureImageInCell_(signatureCell, buildIntervenantSignatureImageModel_(record), true);
      if (inserted) appendSignatureCaption_(signatureCell, buildIntervenantSignatureCaption_(record), true);
    }
    appendHiddenMarker_(signatureCell, APP.MARKERS.INTERVENANT);
  } else {
    appendHiddenMarker_(labelCell, APP.MARKERS.INTERVENANT);
  }
}

function renderAgentBlock_(doc, record) {
  const body = doc.getBody();
  const markerResult = body.findText(escapeRegex_(APP.MARKERS.AGENT));
  if (!markerResult) return;
  const cell = getParentCell_(markerResult.getElement());
  if (!cell) return;

  clearTableCell_(cell);
  if (record.statutSignatureAgent === APP.STATUS.SIGNED) {
    const inserted = insertSignatureImageInCell_(cell, buildAgentSignatureImageModel_(record), true);
    if (inserted) appendSignatureCaption_(cell, buildAgentSignatureCaption_(record), true);
  }
  appendHiddenMarker_(cell, APP.MARKERS.AGENT);
}

function buildIntervenantSignatureCaption_(record) {
  const userId = String(record.intervenantUserId || '').trim();
  const dateSession = formatMaybeDateFr_(record.dateSession);
  return ['Débriefé par ' + userId, dateSession].filter(Boolean).join(' ');
}

function buildAgentSignatureCaption_(record) {
  const agentLabel = [String(record.agentMatricule || '').trim(), String(record.agentPrenom || extractFirstName_(record.agentName)).trim()].filter(Boolean).join(' ');
  const dateSession = formatMaybeDateFr_(record.dateSession);
  return ['Vu et visé par ' + agentLabel, dateSession].filter(Boolean).join(' ');
}

function appendSignatureCaption_(cell, textValue, centered) {
  const paragraph = cell.appendParagraph(String(textValue || '').trim());
  paragraph.setAlignment(centered ? DocumentApp.HorizontalAlignment.CENTER : DocumentApp.HorizontalAlignment.LEFT);
  paragraph.setSpacingBefore(3);
  paragraph.setSpacingAfter(2);
  const text = paragraph.editAsText();
  text.setBold(false);
  text.setItalic(true);
  text.setFontSize(7);
  try { text.setForegroundColor('#444444'); } catch (e) {}
  try { text.setFontFamily('Arial'); } catch (e) {}
  return paragraph;
}


function prepareIntervenantLayout_(doc, userId) {
  const body = doc.getBody();
  const anchors = findIntervenantCells_(body);
  if (!anchors || !anchors.labelCell) return false;

  const currentCell = anchors.labelCell;
  const adjacentCell = anchors.signatureCell;

  clearTableCell_(currentCell);
  const p = currentCell.appendParagraph(String(userId || '').trim());
  p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  const t = p.editAsText();
  t.setFontSize(8);
  t.setBold(false);

  if (adjacentCell && adjacentCell !== currentCell) {
    clearTableCell_(adjacentCell);
    appendHiddenMarker_(adjacentCell, APP.MARKERS.INTERVENANT);
  }
  return true;
}

function prepareAgentSignatureAnchor_(doc) {
  const body = doc.getBody();
  replaceTextSafely_(body, '{{SIGNATURE}}', APP.MARKERS.AGENT);
}

function findIntervenantCells_(body) {
  const placeholderResult = body.findText('\{\{INTERVENANT\}\}');
  if (placeholderResult) {
    const cell = getParentCell_(placeholderResult.getElement());
    if (cell) {
      return { labelCell: cell, signatureCell: getAdjacentCell_(cell, 1) || null };
    }
  }

  const tables = body.getTables();
  for (let t = 0; t < tables.length; t++) {
    const table = tables[t];
    if (table.getNumRows() < 2) continue;
    const firstRow = table.getRow(0);
    const headerTexts = [];
    for (let c = 0; c < firstRow.getNumCells(); c++) {
      headerTexts.push(String(firstRow.getCell(c).getText() || '').trim().toLowerCase());
    }
    const joined = headerTexts.join(' | ');
    if (joined.indexOf('intervenant') === -1) continue;

    const targetRow = table.getRow(1);
    if (targetRow.getNumCells() < 1) continue;
    return {
      labelCell: targetRow.getCell(0),
      signatureCell: targetRow.getNumCells() > 1 ? targetRow.getCell(1) : null
    };
  }
  return null;
}

function getAdjacentCell_(cell, offset) {
  const row = cell.getParent && cell.getParent();
  if (!row || !row.getNumCells) return null;
  const currentIndex = getCellIndexInRow_(cell);
  if (currentIndex < 0) return null;
  const targetIndex = currentIndex + Number(offset || 0);
  if (targetIndex < 0 || targetIndex >= row.getNumCells()) return null;
  return row.getCell(targetIndex);
}

function getCellIndexInRow_(cell) {
  const row = cell.getParent && cell.getParent();
  if (!row || !row.getNumCells) return -1;
  for (let i = 0; i < row.getNumCells(); i++) {
    if (row.getCell(i) === cell) return i;
  }
  return -1;
}

function buildIntervenantSignatureImageModel_(record) {
  // ADMIN/USER (intervenant) → template 'manager'
  return {
    template: 'manager',
    managerDisplay: cleanIntervenantLabel_(record.intervenantNom) || String(record.intervenantUserId || '').trim(),
    managerUserId:  String(record.intervenantUserId || '').trim(),
    managerTitle:   String(record.intervenantTitle || APP.SIGNATURE_TITLE_DEFAULT).trim(),
    sessionDate:    formatMaybeDateFr_(record.dateSession),
    width:  252,   // 3.5 inches — maintient ratio SVG 1200:320 (3.75:1) → height=67
    height: 67,
    alignCenter: true
  };
}

function buildAgentSignatureImageModel_(record) {
  // AGENT (ambassadeur) → template 'agent'
  return {
    template: 'agent',
    signatureTitle: getAgentSignatureTitleForRecord_(record),  // ex: 'Ambassadeur'
    prenom:    String(record.agentPrenom || extractFirstName_(record.agentName)).trim(),
    matricule: String(record.agentMatricule || '').trim(),
    sessionDate: formatMaybeDateFr_(record.dateSession),
    width:  216,   // 3 inches — ratio SVG 3.75:1 → height=58
    height: 58,
    alignCenter: true
  };
}

function getAgentSignatureTitleForRecord_(record) {
  const explicitTitle = String(record.agentSignatureTitle || '').trim();
  if (explicitTitle) return explicitTitle;
  const matricule = String(record.agentMatricule || '').trim();
  if (matricule) {
    const user = getUsers_().find(item => String(item.agentMatricule || '').trim() === matricule);
    if (user && String(user.title || '').trim()) return String(user.title || '').trim();
  }
  return 'Ambassadeur';
}

function escapeXml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function appendHiddenMarker_(cell, marker) {
  const p = cell.appendParagraph(marker);
  const text = p.editAsText();
  text.setFontSize(1);
  try { text.setForegroundColor('#ffffff'); } catch (err) { console.warn('setForegroundColor:', err); }
}

function clearTableCell_(cell) {
  while (cell.getNumChildren() > 0) {
    cell.removeChild(cell.getChild(0));
  }
}

function getParentCell_(element) {
  let current = element;
  while (current) {
    if (current.getType && current.getType() === DocumentApp.ElementType.TABLE_CELL) {
      return current.asTableCell();
    }
    current = current.getParent ? current.getParent() : null;
  }
  return null;
}

function insertSignatureImageInCell_(cell, model, centered) {
  const alignment = centered
    ? DocumentApp.HorizontalAlignment.CENTER
    : DocumentApp.HorizontalAlignment.LEFT;
 
  try {
    if (String((model && model.template) || '').toLowerCase() === 'manager') {
      // ── Bloc intervenant / Team Leader ─────────────────────────
      const title = String(model.managerTitle   || APP.SIGNATURE_TITLE_DEFAULT).trim();
 
      // 1. Étiquette du rôle
      const titlePara = cell.appendParagraph('Signature ' + title);
      titlePara.setAlignment(alignment);
      titlePara.setSpacingBefore(3);
      titlePara.setSpacingAfter(4);
      const titleText = titlePara.editAsText();
      titleText.setFontSize(7);
      titleText.setBold(false);
      titleText.setItalic(false);
      try { titleText.setForegroundColor('#505560'); } catch (e) {}
 
    } 
    return true; // ← succès garanti dès qu'on arrive ici
 
  } catch (err) {
    // Dernier recours : insère au moins le nom en texte brut
    try {
      const fallbackName = (model.template === 'manager')
        ? String(model.managerDisplay || model.managerUserId || 'Signé').trim()
        : String(model.prenom || 'Signé').trim();
      const p = cell.appendParagraph(fallbackName);
      p.setAlignment(centered ? DocumentApp.HorizontalAlignment.CENTER : DocumentApp.HorizontalAlignment.LEFT);
      p.editAsText().setItalic(true).setFontSize(12);
      return true;
    } catch (err2) {
      return false;
    }
  }
}

function getTrackingRowsByPeriod_(period) {
  const cleanPeriod = normalizePeriod_(period || periodFromDate_(new Date()));
  const sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.TRACKING);
  const data = getDataRows_(sheet);
  return data
    .map((row, i) => mapTrackingRow_(row, i + 2))
    .filter(item => deriveTrackingPeriod_(item) === cleanPeriod);
}

function getTrackingRowsByPeriodFiltered_(session, period) {
  const rows = getTrackingRowsByPeriod_(period);
  if (isAdmin_(session)) return rows;
  if (isAgent_(session)) {
    return rows.filter(item => isSameAgentForSession_(session, item.agentMatricule));
  }
  // USER (intervenant) : ne voit que ses propres fiches
  return rows.filter(item =>
    String(item.intervenantUserId || '').trim() === String(session.userId || '').trim()
  );
}

function canSessionSignIntervenant_(session, record) {
  if (!record || !record.rowNumber) return false;
  if (String(record.statutSignatureIntervenant || APP.STATUS.PENDING) === APP.STATUS.SIGNED) return false;
  if (isAgent_(session)) return false;
  return isAdmin_(session) || String(record.intervenantUserId || '').trim() === String(session.userId || '').trim();
}

function canSessionSignAgent_(session, record) {
  if (!record || !record.rowNumber) return false;
  if (String(record.statutSignatureAgent || APP.STATUS.PENDING) === APP.STATUS.SIGNED) return false;
  if (isAdmin_(session)) return true;
  if (!isAgent_(session)) return false;
  return isSameAgentForSession_(session, record.agentMatricule);
}

function getUsers_() {
  const sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.USERS);
  const data = getDataRows_(sheet);
  return data.map(row => ({
    userId: String(row[0] || '').trim(),
    name: String(row[1] || '').trim(),
    passwordHash: String(row[2] || '').trim(),
    role: String(row[3] || 'USER').trim().toUpperCase(),
    title: String(row[4] || APP.SIGNATURE_TITLE_DEFAULT).trim(),
    active: /^oui|true|1$/i.test(String(row[5] || 'OUI').trim()),
    agentMatricule: String(row[6] || '').trim()
  })).filter(item => item.userId);
}

function getAgents_() {
  const sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.AGENTS);
  const data = getDataRows_(sheet);
  return data.map(row => {
    const nom = String(row[1] || '').trim();
    const prenom = String(row[2] || '').trim();
    return {
      matricule: String(row[0] || '').trim(),
      nom: nom,
      prenom: prenom,
      fullName: [prenom, nom].filter(Boolean).join(' ').trim(),
      sousEquipe: String(row[3] || '').trim(),
      active: /^oui|true|1$/i.test(String(row[4] || 'OUI').trim())
    };
  }).filter(item => item.matricule && item.active);
}

function getThemeModels_() {
  const sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.THEME_MODELS);
  const data = getDataRows_(sheet);
  return data.map(row => ({
    name: String(row[0] || '').trim(),
    content: String(row[1] || '').trim(),
    active: !/^non|false|0$/i.test(String(row[2] || 'OUI').trim())
  })).filter(item => item.name);
}

function getParamsMap_() {
  const sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.PARAMS);
  const data = getDataRows_(sheet);
  return data.reduce((acc, row) => {
    const key = String(row[0] || '').trim();
    if (key) acc[key] = row[1];
    return acc;
  }, {});
}

function initializeWorkbook_() {
  rememberActiveSpreadsheet_();
  const ss = getAppSpreadsheet_();

  ensureSheet_(ss, APP.SHEETS.AGENTS, ['Matricule', 'Nom', 'Prenom', 'SousEquipe', 'Actif']);
  ensureSheet_(ss, APP.SHEETS.USERS, ['UserID', 'NomPrenom', 'PasswordHash', 'Role', 'Title', 'Actif', 'AgentMatricule']);
  ensureSheet_(ss, APP.SHEETS.PARAMS, ['Cle', 'Valeur']);
  ensureSheet_(ss, APP.SHEETS.THEME_MODELS, ['NomModele', 'Contenu', 'Actif']);
  ensureSheet_(ss, APP.SHEETS.TRACKING, [
    'RecordId',
    'Periode',
    'DateSession',
    'MoisLibelle',
    'AgentMatricule',
    'AgentNomPrenom',
    'AgentPrenom',
    'SousEquipe',
    'IntituleSession',
    'Duree',
    'Lieu',
    'IntervenantUserID',
    'IntervenantNom',
    'IntervenantTitle',
    'ThemeModele',
    'ThemesTexte',
    'Remarques',
    'SourceDocFileId',
    'SourceDocUrl',
    'PdfFileId',
    'PdfUrl',
    'StatutPdf',
    'StatutSignatureIntervenant',
    'DateSignatureIntervenant',
    'SigneIntervenantPar',
    'StatutSignatureAgent',
    'DateSignatureAgent',
    'SigneAgentPar',
    'DateGeneration',
    'GenerePar',
    'MajLe'
  ]);

  // ── Quiz sheets ──
  ensureSheet_(ss, APP.SHEETS.QUIZ_MODELES, [
    'ModeleId', 'Titre', 'Description', 'QuestionsJson', 'Actif', 'CreePar', 'CreeLe'
  ]);
  ensureSheet_(ss, APP.SHEETS.QUIZ_SESSIONS, [
    'SessionId', 'Token', 'ModeleId', 'Titre', 'DateSession', 'Lieu', 'Duree', 'Actif', 'CreePar', 'CreeLe'
  ]);
  ensureSheet_(ss, APP.SHEETS.QUIZ_REPONSES, [
    'ReponseId', 'SessionId', 'Matricule', 'Nom', 'Prenom', 'Groupe',
    'IsExterne', 'ReponsesJson', 'Score', 'NoteMax', 'CorrigePar', 'CorrigeLe',
    'PdfFileId', 'PdfUrl', 'TrackingRowNumber', 'SoumisLe'
  ]);

  seedDefaultParams_();
  seedThemeModels_();
  formatSheets_();
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

function seedDefaultParams_() {
  const sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.PARAMS);
  const existing = getParamsMap_();
  const defaults = {};
  defaults[APP.PARAM_KEYS.TEMPLATE_ID] = existing[APP.PARAM_KEYS.TEMPLATE_ID] || '';
  defaults[APP.PARAM_KEYS.OUTPUT_FOLDER_ID] = existing[APP.PARAM_KEYS.OUTPUT_FOLDER_ID] || '';
  defaults[APP.PARAM_KEYS.DEFAULT_TITLE] = existing[APP.PARAM_KEYS.DEFAULT_TITLE] || 'Émargement mensuel';
  defaults[APP.PARAM_KEYS.DEFAULT_DURATION] = existing[APP.PARAM_KEYS.DEFAULT_DURATION] || '1 journée';
  defaults[APP.PARAM_KEYS.DEFAULT_LOCATION] = existing[APP.PARAM_KEYS.DEFAULT_LOCATION] || '';
  defaults[APP.PARAM_KEYS.APP_NAME] = existing[APP.PARAM_KEYS.APP_NAME] || 'Émargement mensuel';

  Object.keys(defaults).forEach(key => upsertParam_(sheet, key, defaults[key]));
}

function seedThemeModels_() {
  const sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.THEME_MODELS);
  if (sheet.getLastRow() > 1) return;
  const rows = [
    ['Standard', 'Accueil\nPoint sécurité\nConsignes\nRappel qualité\nQuestions / réponses', 'OUI'],
    ['Réunion mensuelle', 'Bilan du mois\nObjectifs du mois suivant\nPoints bloquants\nPlan d’action', 'OUI']
  ];
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function formatSheets_() {
  const ss = getAppSpreadsheet_();
  Object.values(APP.SHEETS).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
    sheet.autoResizeColumns(1, Math.min(sheet.getLastColumn(), 12));
  });
}

function setupWorkbookSilently_() {
  rememberActiveSpreadsheet_();
  const ss = getAppSpreadsheet_();
  if (!ss.getSheetByName(APP.SHEETS.AGENTS)) initializeWorkbook_();
  ensureQuizSheets_();
}

function ensureQuizSheets_() {
  var ss = getAppSpreadsheet_();
  if (!ss.getSheetByName(APP.SHEETS.QUIZ_MODELES)) {
    ensureSheet_(ss, APP.SHEETS.QUIZ_MODELES, [
      'ModeleId', 'Titre', 'Description', 'QuestionsJson', 'Actif', 'CreePar', 'CreeLe'
    ]);
  }
  if (!ss.getSheetByName(APP.SHEETS.QUIZ_SESSIONS)) {
    ensureSheet_(ss, APP.SHEETS.QUIZ_SESSIONS, [
      'SessionId', 'Token', 'ModeleId', 'Titre', 'DateSession', 'Lieu', 'Duree', 'Actif', 'CreePar', 'CreeLe'
    ]);
  }
  if (!ss.getSheetByName(APP.SHEETS.QUIZ_REPONSES)) {
    ensureSheet_(ss, APP.SHEETS.QUIZ_REPONSES, [
      'ReponseId', 'SessionId', 'Matricule', 'Nom', 'Prenom', 'Groupe',
      'IsExterne', 'ReponsesJson', 'Score', 'NoteMax', 'CorrigePar', 'CorrigeLe',
      'PdfFileId', 'PdfUrl', 'TrackingRowNumber', 'SoumisLe'
    ]);
  }
}

function upsertParam_(sheet, key, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const idx = keys.findIndex(item => String(item).trim() === key);
    if (idx >= 0) {
      sheet.getRange(idx + 2, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function getTrackingIndexByPeriod_(period) {
  const sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.TRACKING);
  const data = getDataRows_(sheet);
  const index = {};
  data.forEach((row, i) => {
    const rowPeriod = String(row[1] || '').trim();
    const matricule = String(row[4] || '').trim();
    const dateSession = String(row[2] || '').trim();
    if (rowPeriod === period && matricule) {
      const key = matricule + '|' + dateSession;
      index[key] = mapTrackingRow_(row, i + 2);
    }
  });
  return index;
}

function getTrackingRecordByRow_(rowNumber) {
  const sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.TRACKING);
  if (!rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) return null;
  const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  return mapTrackingRow_(row, rowNumber);
}

function mapTrackingRow_(row, rowNumber) {
  while (row.length < 31) row.push('');
  return {
    rowNumber: rowNumber,
    recordId: row[0],
    period: row[1],
    dateSession: row[2],
    monthLabel: row[3],
    agentMatricule: row[4],
    agentName: row[5],
    agentPrenom: row[6],
    sousEquipe: row[7],
    intituleSession: row[8],
    duree: row[9],
    lieu: row[10],
    intervenantUserId: row[11],
    intervenantNom: row[12],
    intervenantTitle: row[13],
    themeModel: row[14],
    themesText: row[15],
    remarques: row[16],
    sourceDocFileId: row[17],
    sourceDocUrl: row[18],
    pdfFileId: row[19],
    pdfUrl: row[20],
    statutPdf: row[21],
    statutSignatureIntervenant: row[22],
    dateSignatureIntervenant: row[23],
    signeIntervenantPar: row[24],
    statutSignatureAgent: row[25],
    dateSignatureAgent: row[26],
    signeAgentPar: row[27],
    dateGeneration: row[28],
    generePar: row[29],
    majLe: row[30]
  };
}

function deriveTrackingPeriod_(record) {
  const rawPeriod = record && record.period ? record.period : '';
  try {
    const explicit = rawPeriod ? normalizePeriod_(rawPeriod) : '';
    if (explicit) return explicit;
  } catch (err) { console.warn('deriveTrackingPeriod_ normalizePeriod_:', err); }

  const rawDate = record ? record.dateSession : '';
  if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
    return Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  const text = String(rawDate || '').trim();
  const isoMatch = text.match(/^(\d{4}-\d{2})/);
  if (isoMatch) return isoMatch[1];
  const frDateMatch = text.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (frDateMatch) return frDateMatch[3] + '-' + frDateMatch[2];
  return '';
}

function upsertTrackingRecord_(sheet, rowNumber, record) {
  const row = [[
    record.recordId,
    record.period,
    record.dateSession,
    record.monthLabel,
    record.agentMatricule,
    record.agentName,
    record.agentPrenom,
    record.sousEquipe,
    record.intituleSession,
    record.duree,
    record.lieu,
    record.intervenantUserId,
    record.intervenantNom,
    record.intervenantTitle,
    record.themeModel,
    record.themesText,
    record.remarques,
    record.sourceDocFileId,
    record.sourceDocUrl,
    record.pdfFileId,
    record.pdfUrl,
    record.statutPdf,
    record.statutSignatureIntervenant,
    record.dateSignatureIntervenant,
    record.signeIntervenantPar,
    record.statutSignatureAgent,
    record.dateSignatureAgent,
    record.signeAgentPar,
    record.dateGeneration,
    record.generePar,
    record.majLe
  ]];

  if (rowNumber) {
    sheet.getRange(rowNumber, 1, 1, row[0].length).setValues(row);
  } else {
    sheet.appendRow(row[0]);
  }
}

function getDataRows_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
}

function clearDataKeepHeader_(sheet) {
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
}

function sha256Hex_(input) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  return bytes.map(b => {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function requireSession_(token) {
  const raw = PropertiesService.getScriptProperties().getProperty('SESSION_' + token);
  if (!raw) throw new Error('Session expirée. Reconnecte-toi.');
  const session = JSON.parse(raw);
  if (!session.expiresAt || session.expiresAt < Date.now()) {
    PropertiesService.getScriptProperties().deleteProperty('SESSION_' + token);
    throw new Error('Session expirée. Reconnecte-toi.');
  }
  return session;
}

function isAdmin_(session) {
  return String(session.role || '').toUpperCase() === 'ADMIN';
}

function isAgent_(session) {
  return String(session.role || '').toUpperCase() === 'AGENT';
}

function isSameAgentForSession_(session, agentMatricule) {
  const sessionMatricule = String(session.agentMatricule || session.userId || '').trim();
  return !!sessionMatricule && sessionMatricule === String(agentMatricule || '').trim();
}

function trashFileIfExists_(fileId) {
  if (!fileId) return;
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (err) { console.warn('trashFileIfExists_:', err); }
}

function trashFilesByNameInFolder_(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) files.next().setTrashed(true);
}

function cleanIntervenantLabel_(name) {
  return String(name || '').replace(/\s*\([^)]*\)\s*/g, ' ').replace(/\s+/g, ' ').trim();
}

function formatMaybeDateFr_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return formatDateFr_(value);
  const str = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return formatDateFr_(parseLocalDate_(str));
  return str;
}

function formatMaybeDateTimeFr_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  }
  return String(value).trim();
}

function replaceTextSafely_(body, placeholder, value) {
  var safeValue = (value == null ? '' : String(value)).replace(/\$/g, '$$$$');
  body.replaceText(escapeRegex_(placeholder), safeValue);
}

function escapeRegex_(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function parseLocalDate_(input) {
  if (Object.prototype.toString.call(input) === '[object Date]' && !isNaN(input.getTime())) {
    return new Date(input.getFullYear(), input.getMonth(), input.getDate(), 12, 0, 0);
  }

  const clean = String(input == null ? '' : input).trim();
  if (!clean) throw new Error('Date de session manquante.');

  const isoDateMatch = clean.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoDateMatch) {
    return new Date(Number(isoDateMatch[1]), Number(isoDateMatch[2]) - 1, Number(isoDateMatch[3]), 12, 0, 0);
  }

  const isoDateTimeMatch = clean.match(/^(\d{4})-(\d{2})-(\d{2})[T\s].*$/);
  if (isoDateTimeMatch) {
    return new Date(Number(isoDateTimeMatch[1]), Number(isoDateTimeMatch[2]) - 1, Number(isoDateTimeMatch[3]), 12, 0, 0);
  }

  const frDateMatch = clean.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (frDateMatch) {
    return new Date(Number(frDateMatch[3]), Number(frDateMatch[2]) - 1, Number(frDateMatch[1]), 12, 0, 0);
  }

  throw new Error('Date de session invalide.');
}

function formatDateFr_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function periodFromDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');
}

function normalizePeriod_(period) {
  if (period instanceof Date && !isNaN(period.getTime())) {
    return Utilities.formatDate(period, Session.getScriptTimeZone(), 'yyyy-MM');
  }

  const clean = String(period || '').trim();
  if (!clean) throw new Error('Période invalide. Format attendu : yyyy-MM');
  if (/^\d{4}-\d{2}$/.test(clean)) return clean;

  const slashMatch = clean.match(/^(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    const month = ('0' + Number(slashMatch[1])).slice(-2);
    return slashMatch[2] + '-' + month;
  }

  const lower = clean.toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  const monthMap = {
    janvier: '01', fevrier: '02', mars: '03', avril: '04', mai: '05', juin: '06',
    juillet: '07', aout: '08', septembre: '09', octobre: '10', novembre: '11', decembre: '12'
  };
  const frMatch = lower.match(/^(janvier|fevrier|mars|avril|mai|juin|juillet|aout|septembre|octobre|novembre|decembre)\s+(\d{4})$/);
  if (frMatch) {
    return frMatch[2] + '-' + monthMap[frMatch[1]];
  }

  throw new Error('Période invalide. Format attendu : yyyy-MM');
}

function monthLabelFr_(date) {
  const months = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
  return months[date.getMonth()] + ' ' + date.getFullYear();
}

function sanitizeFileName_(str) {
  return String(str || '').replace(/[\/\\*?"<>|]/g, '_').trim();
}

function buildPdfName_(prenom, sessionDate) {
  var dateSuffix = Utilities.formatDate(sessionDate, Session.getScriptTimeZone(), 'dd-MM-yyyy');
  return 'PROD-ENR-004C-Feuille de présence - ' + sanitizeFileName_(prenom) + ' - ' + dateSuffix + '.pdf';
}

function buildSourceDocName_(fullName, sessionDate) {
  var dateSuffix = Utilities.formatDate(sessionDate, Session.getScriptTimeZone(), 'dd-MM-yyyy');
  return 'SRC - PROD-ENR-004C - ' + sanitizeFileName_(fullName) + ' - ' + dateSuffix;
}

function appendLabeledLine_(body, label, placeholder) {
  const p = body.appendParagraph('');
  p.appendText(label).setBold(true);
  p.appendText(placeholder);
  return p;
}

function getOrCreateSubFolder_(parentFolder, name) {
  const folders = parentFolder.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(name);
}

function openDocumentWithRetry_(docId, maxTries) {
  let lastErr = null;
  for (let i = 0; i < maxTries; i++) {
    try {
      if (i > 0) Utilities.sleep(500 + i * 300);
      return DocumentApp.openById(docId);
    } catch (err) {
      lastErr = err;
    }
  }
  throw new Error('Impossible d’accéder au document. Veuillez réessayer plus tard. Détail : ' + (lastErr && lastErr.message ? lastErr.message : lastErr));
}
function extractFirstName_(fullName) {
  return String(fullName || '').trim().split(/\s+/)[0] || '';
}


// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  MODULE QUIZ — Backend                                                      ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

// ── Accès données Quiz ──

function getQuizModeles_() {
  var sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.QUIZ_MODELES);
  if (!sheet) return [];
  var data = getDataRows_(sheet);
  return data.map(function(row, i) {
    return {
      rowNumber: i + 2,
      modeleId: String(row[0] || '').trim(),
      titre: String(row[1] || '').trim(),
      description: String(row[2] || '').trim(),
      questionsJson: String(row[3] || '[]'),
      actif: !/^non|false|0$/i.test(String(row[4] || 'OUI').trim()),
      creePar: String(row[5] || '').trim(),
      creeLe: row[6] ? String(row[6]) : ''
    };
  }).filter(function(m) { return m.modeleId; });
}

function getQuizSessions_() {
  var sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.QUIZ_SESSIONS);
  if (!sheet) return [];
  var data = getDataRows_(sheet);
  return data.map(function(row, i) {
    return {
      rowNumber: i + 2,
      sessionId: String(row[0] || '').trim(),
      token: String(row[1] || '').trim(),
      modeleId: String(row[2] || '').trim(),
      titre: String(row[3] || '').trim(),
      dateSession: String(row[4] || '').trim(),
      lieu: String(row[5] || '').trim(),
      duree: String(row[6] || '').trim(),
      actif: !/^non|false|0$/i.test(String(row[7] || 'OUI').trim()),
      creePar: String(row[8] || '').trim(),
      creeLe: row[9] ? String(row[9]) : ''
    };
  }).filter(function(s) { return s.sessionId; });
}

function getQuizReponses_() {
  var sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.QUIZ_REPONSES);
  if (!sheet) return [];
  var data = getDataRows_(sheet);
  return data.map(function(row, i) {
    return {
      rowNumber: i + 2,
      reponseId: String(row[0] || '').trim(),
      sessionId: String(row[1] || '').trim(),
      matricule: String(row[2] || '').trim(),
      nom: String(row[3] || '').trim(),
      prenom: String(row[4] || '').trim(),
      groupe: String(row[5] || '').trim(),
      isExterne: /^(oui|true|1)$/i.test(String(row[6] || 'NON').trim()),
      reponsesJson: String(row[7] || '[]'),
      score: row[8] !== undefined && row[8] !== '' ? Number(row[8]) : null,
      noteMax: row[9] !== undefined && row[9] !== '' ? Number(row[9]) : null,
      corrigePar: String(row[10] || '').trim(),
      corrigeLe: row[11] ? String(row[11]) : '',
      pdfFileId: String(row[12] || '').trim(),
      pdfUrl: String(row[13] || '').trim(),
      trackingRowNumber: row[14] ? Number(row[14]) : null,
      soumisLe: row[15] ? String(row[15]) : ''
    };
  }).filter(function(r) { return r.reponseId; });
}

function getQuizReponsesBySession_(sessionId) {
  return getQuizReponses_().filter(function(r) {
    return r.sessionId === sessionId;
  });
}

// ── Accès quiz par token (participant, pas besoin de session auth) ──

function quizAccessByToken(tokenInput, matriculeInput) {
  setupWorkbookSilently_();
  var token = String(tokenInput || '').trim();
  var matricule = String(matriculeInput || '').trim();
  if (!token) throw new Error('Token requis.');

  var sessions = getQuizSessions_();
  var session = sessions.find(function(s) { return s.token === token && s.actif; });
  if (!session) throw new Error('Session de quiz introuvable ou inactive.');

  var modeles = getQuizModeles_();
  var modele = modeles.find(function(m) { return m.modeleId === session.modeleId; });
  if (!modele) throw new Error('Modèle de quiz introuvable.');

  // Vérifier si déjà soumis
  var reponses = getQuizReponsesBySession_(session.sessionId);
  var existing = reponses.find(function(r) { return r.matricule === matricule; });
  if (existing) throw new Error('Tu as déjà répondu à ce quiz.');

  var questions = [];
  try { questions = JSON.parse(modele.questionsJson); } catch (e) { questions = []; }

  // Chercher l'agent par matricule
  var agent = null;
  if (matricule) {
    var agents = getAgents_();
    agent = agents.find(function(a) { return a.matricule === matricule; });
  }

  return {
    sessionId: session.sessionId,
    titre: session.titre || modele.titre,
    description: modele.description,
    dateSession: session.dateSession,
    lieu: session.lieu,
    duree: session.duree,
    questions: questions.map(function(q) {
      // Ne pas envoyer correctAnswer au participant
      return {
        id: q.id, type: q.type, question: q.question,
        options: q.options || [], points: q.points || 1
      };
    }),
    agent: agent ? { matricule: agent.matricule, fullName: agent.fullName, sousEquipe: agent.sousEquipe } : null,
    isExterne: !agent
  };
}

// ── Soumission des réponses par le participant ──

function quizSubmitAnswers(payload) {
  setupWorkbookSilently_();
  var sessionId = String(payload.sessionId || '').trim();
  var matricule = String(payload.matricule || '').trim();
  var nom = String(payload.nom || '').trim();
  var prenom = String(payload.prenom || '').trim();
  var groupe = String(payload.groupe || '').trim();
  var isExterne = !!payload.isExterne;
  var answers = payload.answers || {};

  if (!sessionId) throw new Error('Session ID manquant.');
  if (!matricule && !nom) throw new Error('Identifiant ou nom requis.');

  var sessions = getQuizSessions_();
  var session = sessions.find(function(s) { return s.sessionId === sessionId && s.actif; });
  if (!session) throw new Error('Session fermée ou introuvable.');

  // Vérifier doublon
  var reponses = getQuizReponsesBySession_(sessionId);
  var key = matricule || (nom + ' ' + prenom).trim();
  var existing = reponses.find(function(r) { return r.matricule === key; });
  if (existing) throw new Error('Réponses déjà soumises.');

  // Charger les questions pour correction auto
  var modeles = getQuizModeles_();
  var modele = modeles.find(function(m) { return m.modeleId === session.modeleId; });
  var questions = [];
  try { questions = JSON.parse(modele.questionsJson); } catch (e) { questions = []; }

  // Correction automatique (QCM et vrai/faux)
  var score = 0;
  var noteMax = 0;
  var detailedAnswers = questions.map(function(q) {
    var userAnswer = answers[q.id] || '';
    var points = Number(q.points || 1);
    noteMax += points;
    var correct = null;
    if (q.type === 'qcm' || q.type === 'vrai_faux') {
      correct = String(q.correctAnswer || '').trim().toLowerCase() === String(userAnswer).trim().toLowerCase();
      if (correct) score += points;
    }
    return {
      id: q.id, type: q.type, question: q.question,
      userAnswer: userAnswer, correctAnswer: q.correctAnswer || '',
      points: points, correct: correct
    };
  });

  var sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.QUIZ_REPONSES);
  var reponseId = Utilities.getUuid();
  sheet.appendRow([
    reponseId,
    sessionId,
    matricule || key,
    nom,
    prenom,
    groupe,
    isExterne ? 'OUI' : 'NON',
    JSON.stringify(detailedAnswers),
    score,
    noteMax,
    '', // corrigePar
    '', // corrigeLe
    '', // pdfFileId
    '', // pdfUrl
    '', // trackingRowNumber
    new Date() // soumisLe
  ]);

  return {
    success: true,
    score: score,
    noteMax: noteMax,
    message: 'Réponses enregistrées. Score provisoire : ' + score + '/' + noteMax
  };
}

// ── Admin : gestion des modèles de quiz ──

function getQuizBootstrapData(token) {
  var session = requireSession_(token);
  if (isAgent_(session)) throw new Error('Accès refusé.');
  ensureQuizSheets_();
  return {
    modeles: getQuizModeles_(),
    sessions: getQuizSessions_(),
    reponses: getQuizReponses_()
  };
}

function saveQuizModele(token, payload) {
  var session = requireSession_(token);
  if (!isAdmin_(session)) throw new Error('Accès réservé admin.');

  var modeleId = String(payload.modeleId || '').trim();
  var titre = String(payload.titre || '').trim();
  if (!titre) throw new Error('Titre requis.');

  var questionsJson = JSON.stringify(payload.questions || []);
  var sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.QUIZ_MODELES);

  if (modeleId) {
    // Mise à jour
    var modeles = getQuizModeles_();
    var existing = modeles.find(function(m) { return m.modeleId === modeleId; });
    if (!existing) throw new Error('Modèle introuvable.');
    var row = existing.rowNumber;
    sheet.getRange(row, 1, 1, 7).setValues([[
      modeleId,
      titre,
      String(payload.description || ''),
      questionsJson,
      payload.actif === false ? 'NON' : 'OUI',
      existing.creePar,
      existing.creeLe
    ]]);
  } else {
    // Création
    modeleId = Utilities.getUuid();
    sheet.appendRow([
      modeleId,
      titre,
      String(payload.description || ''),
      questionsJson,
      'OUI',
      session.name,
      new Date()
    ]);
  }
  return { success: true, modeleId: modeleId };
}

function deleteQuizModele(token, modeleId) {
  var session = requireSession_(token);
  if (!isAdmin_(session)) throw new Error('Accès réservé admin.');

  var sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.QUIZ_MODELES);
  var modeles = getQuizModeles_();
  var existing = modeles.find(function(m) { return m.modeleId === modeleId; });
  if (!existing) throw new Error('Modèle introuvable.');
  sheet.deleteRow(existing.rowNumber);
  return { success: true };
}

// ── Admin : gestion des sessions de quiz ──

function saveQuizSession(token, payload) {
  var session = requireSession_(token);
  if (!isAdmin_(session)) throw new Error('Accès réservé admin.');

  var sessionId = String(payload.sessionId || '').trim();
  var titre = String(payload.titre || '').trim();
  var modeleId = String(payload.modeleId || '').trim();
  if (!titre) throw new Error('Titre requis.');
  if (!modeleId) throw new Error('Modèle requis.');

  var sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.QUIZ_SESSIONS);

  if (sessionId) {
    // Mise à jour
    var sessions = getQuizSessions_();
    var existing = sessions.find(function(s) { return s.sessionId === sessionId; });
    if (!existing) throw new Error('Session introuvable.');
    sheet.getRange(existing.rowNumber, 1, 1, 10).setValues([[
      sessionId,
      existing.token,
      modeleId,
      titre,
      String(payload.dateSession || ''),
      String(payload.lieu || ''),
      String(payload.duree || ''),
      payload.actif === false ? 'NON' : 'OUI',
      existing.creePar,
      existing.creeLe
    ]]);
  } else {
    // Création avec token unique
    sessionId = Utilities.getUuid();
    var quizToken = Utilities.getUuid().split('-')[0].toUpperCase();
    sheet.appendRow([
      sessionId,
      quizToken,
      modeleId,
      titre,
      String(payload.dateSession || ''),
      String(payload.lieu || ''),
      String(payload.duree || ''),
      'OUI',
      session.name,
      new Date()
    ]);
  }
  return { success: true, sessionId: sessionId };
}

function toggleQuizSession(token, sessionId, actif) {
  var session = requireSession_(token);
  if (!isAdmin_(session)) throw new Error('Accès réservé admin.');

  var sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.QUIZ_SESSIONS);
  var sessions = getQuizSessions_();
  var existing = sessions.find(function(s) { return s.sessionId === sessionId; });
  if (!existing) throw new Error('Session introuvable.');
  sheet.getRange(existing.rowNumber, 8).setValue(actif ? 'OUI' : 'NON');
  return { success: true };
}

// ── Admin : correction manuelle ──

function gradeQuizReponse(token, payload) {
  var session = requireSession_(token);
  if (!isAdmin_(session)) throw new Error('Accès réservé admin.');

  var reponseId = String(payload.reponseId || '').trim();
  if (!reponseId) throw new Error('ID réponse manquant.');

  var all = getQuizReponses_();
  var reponse = all.find(function(r) { return r.reponseId === reponseId; });
  if (!reponse) throw new Error('Réponse introuvable.');

  var answers = [];
  try { answers = JSON.parse(reponse.reponsesJson); } catch (e) { answers = []; }

  // Appliquer les notes manuelles
  var manualGrades = payload.grades || {};
  var totalScore = 0;
  var noteMax = 0;
  answers.forEach(function(a) {
    noteMax += Number(a.points || 1);
    if (manualGrades[a.id] !== undefined) {
      a.correct = manualGrades[a.id] > 0;
      totalScore += Number(manualGrades[a.id]);
    } else if (a.correct === true) {
      totalScore += Number(a.points || 1);
    }
  });

  var sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.QUIZ_REPONSES);
  var row = reponse.rowNumber;
  sheet.getRange(row, 8).setValue(JSON.stringify(answers)); // ReponsesJson
  sheet.getRange(row, 9).setValue(totalScore);              // Score
  sheet.getRange(row, 10).setValue(noteMax);                // NoteMax
  sheet.getRange(row, 11).setValue(session.name);           // CorrigePar
  sheet.getRange(row, 12).setValue(new Date());             // CorrigeLe

  return { success: true, score: totalScore, noteMax: noteMax };
}

// ── Admin : génération PDF du quiz rempli ──

function generateQuizPdf(token, reponseId) {
  var session = requireSession_(token);
  if (!isAdmin_(session)) throw new Error('Accès réservé admin.');

  var all = getQuizReponses_();
  var reponse = all.find(function(r) { return r.reponseId === reponseId; });
  if (!reponse) throw new Error('Réponse introuvable.');

  var sessions = getQuizSessions_();
  var quizSession = sessions.find(function(s) { return s.sessionId === reponse.sessionId; });
  if (!quizSession) throw new Error('Session de quiz introuvable.');

  var params = getParamsMap_();
  var outputFolderId = String(params[APP.PARAM_KEYS.OUTPUT_FOLDER_ID] || '').trim();
  if (!outputFolderId) throw new Error('Dossier de sortie non configuré.');

  var answers = [];
  try { answers = JSON.parse(reponse.reponsesJson); } catch (e) { answers = []; }

  // Créer le document Google Docs
  var participantLabel = reponse.prenom
    ? (reponse.prenom + ' ' + reponse.nom).trim()
    : reponse.nom || reponse.matricule;
  var docTitle = 'Quiz - ' + sanitizeFileName_(quizSession.titre) + ' - ' + sanitizeFileName_(participantLabel);

  var doc = DocumentApp.create(docTitle);
  var body = doc.getBody();
  body.clear();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);

  // ── En-tête ──
  var headerTable = body.appendTable([
    ['QUIZ - ÉVALUATION', quizSession.titre]
  ]);
  headerTable.setBorderWidth(0);
  var hRow = headerTable.getRow(0);
  hRow.getCell(0).editAsText().setBold(true).setFontSize(14).setForegroundColor('#1a1a2e');
  hRow.getCell(1).editAsText().setBold(true).setFontSize(12).setForegroundColor('#333333');
  hRow.setMinimumHeight(36);

  body.appendParagraph('').setSpacingAfter(4);

  // ── Infos session ──
  var infoTable = body.appendTable([
    ['Date :', quizSession.dateSession || '-', 'Lieu :', quizSession.lieu || '-'],
    ['Durée :', quizSession.duree || '-', 'Participant :', participantLabel]
  ]);
  infoTable.setBorderWidth(1).setBorderColor('#cccccc');
  for (var ir = 0; ir < 2; ir++) {
    for (var ic = 0; ic < 4; ic++) {
      var cell = infoTable.getRow(ir).getCell(ic);
      cell.editAsText().setFontSize(9);
      if (ic % 2 === 0) { cell.setBackgroundColor('#f0f0f8'); cell.editAsText().setBold(true); }
    }
  }

  if (reponse.matricule) {
    var matRow = body.appendParagraph('Matricule : ' + reponse.matricule + (reponse.groupe ? '  |  Groupe : ' + reponse.groupe : ''));
    matRow.editAsText().setFontSize(9).setForegroundColor('#555555');
  }

  body.appendParagraph('').setSpacingAfter(8);

  // ── Score ──
  var scoreText = 'Score : ' + (reponse.score !== '' && reponse.score !== null ? reponse.score : '?') + ' / ' + (reponse.noteMax || '?');
  var scorePara = body.appendParagraph(scoreText);
  scorePara.editAsText().setBold(true).setFontSize(12).setForegroundColor('#1a1a2e');
  scorePara.setSpacingAfter(10);

  // ── Questions & réponses ──
  answers.forEach(function(a, idx) {
    // Titre de la question
    var qTitle = body.appendParagraph('Question ' + (idx + 1) + ' (' + (a.points || 1) + ' pt' + ((a.points || 1) > 1 ? 's' : '') + ') — ' + (a.type === 'qcm' ? 'QCM' : a.type === 'vrai_faux' ? 'Vrai/Faux' : a.type === 'reponse_courte' ? 'Réponse courte' : 'Texte libre'));
    qTitle.editAsText().setBold(true).setFontSize(10).setForegroundColor('#1a1a2e');
    qTitle.setSpacingBefore(10).setSpacingAfter(3);

    // Énoncé
    var qText = body.appendParagraph(String(a.question || ''));
    qText.editAsText().setFontSize(10);
    qText.setSpacingAfter(4);

    // Réponse donnée
    var answerLabel = 'Réponse : ' + String(a.userAnswer || '(aucune)');
    var aPara = body.appendParagraph(answerLabel);
    aPara.editAsText().setFontSize(10);

    // Correction pour QCM/V-F
    if (a.type === 'qcm' || a.type === 'vrai_faux') {
      var correctLabel = 'Bonne réponse : ' + String(a.correctAnswer || '-');
      var cPara = body.appendParagraph(correctLabel);
      cPara.editAsText().setFontSize(9).setForegroundColor('#555555');

      var resultLabel = a.correct ? '✓ Correct' : '✗ Incorrect';
      var rPara = body.appendParagraph(resultLabel);
      rPara.editAsText().setBold(true).setFontSize(10)
        .setForegroundColor(a.correct ? '#2d7a3a' : '#b5544e');
    } else if (a.correct !== null && a.correct !== undefined) {
      var mPara = body.appendParagraph(a.correct ? '✓ Validé' : '✗ Non validé');
      mPara.editAsText().setBold(true).setFontSize(10)
        .setForegroundColor(a.correct ? '#2d7a3a' : '#b5544e');
    }

    // Séparateur entre questions
    if (idx < answers.length - 1) {
      var sep = body.appendParagraph('─────────────────────────────');
      sep.editAsText().setFontSize(6).setForegroundColor('#cccccc');
      sep.setSpacingBefore(6).setSpacingAfter(6);
    }
  });

  // ── Pied de page ──
  body.appendParagraph('').setSpacingAfter(12);
  var footer = body.appendParagraph('Document généré le ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'));
  footer.editAsText().setFontSize(8).setForegroundColor('#999999').setItalic(true);

  doc.saveAndClose();

  // Déplacer dans le dossier de sortie
  var outputFolder = DriveApp.getFolderById(outputFolderId);
  var quizFolder = getOrCreateSubFolder_(outputFolder, 'Quiz');
  var file = DriveApp.getFileById(doc.getId());
  quizFolder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  // Convertir en PDF
  Utilities.sleep(1500);
  var pdfBlob = file.getAs(MimeType.PDF).setName(docTitle + '.pdf');
  var pdfFile = quizFolder.createFile(pdfBlob);

  // Mettre à jour la réponse avec le PDF
  var sheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.QUIZ_REPONSES);
  sheet.getRange(reponse.rowNumber, 13).setValue(pdfFile.getId());
  sheet.getRange(reponse.rowNumber, 14).setValue(pdfFile.getUrl());

  return { success: true, pdfUrl: pdfFile.getUrl(), pdfFileId: pdfFile.getId() };
}

// ── Rattachement quiz au module émargement/tracking ──

function linkQuizToTracking(token, reponseId) {
  var session = requireSession_(token);
  if (!isAdmin_(session)) throw new Error('Accès réservé admin.');

  var all = getQuizReponses_();
  var reponse = all.find(function(r) { return r.reponseId === reponseId; });
  if (!reponse) throw new Error('Réponse introuvable.');
  if (reponse.trackingRowNumber) throw new Error('Déjà rattaché au suivi.');
  if (!reponse.pdfFileId) throw new Error('Générez d\'abord le PDF avant de rattacher.');

  var sessions = getQuizSessions_();
  var quizSession = sessions.find(function(s) { return s.sessionId === reponse.sessionId; });
  if (!quizSession) throw new Error('Session de quiz introuvable.');

  // Chercher l'agent
  var agent = null;
  if (reponse.matricule && !reponse.isExterne) {
    var agents = getAgents_();
    agent = agents.find(function(a) { return a.matricule === reponse.matricule; });
  }

  var participantLabel = reponse.prenom
    ? (reponse.prenom + ' ' + reponse.nom).trim()
    : reponse.nom || reponse.matricule;

  var dateSession = quizSession.dateSession ? parseLocalDate_(quizSession.dateSession) : new Date();
  var period = periodFromDate_(dateSession);
  var monthLabel = monthLabelFr_(dateSession);

  var trackingSheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.TRACKING);
  var record = {
    recordId: Utilities.getUuid(),
    period: period,
    dateSession: Utilities.formatDate(dateSession, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    monthLabel: monthLabel,
    agentMatricule: reponse.matricule || '',
    agentName: participantLabel,
    agentPrenom: reponse.prenom || extractFirstName_(participantLabel),
    sousEquipe: agent ? agent.sousEquipe : (reponse.groupe || ''),
    intituleSession: 'Quiz : ' + (quizSession.titre || ''),
    duree: quizSession.duree || '',
    lieu: quizSession.lieu || '',
    intervenantUserId: session.userId,
    intervenantNom: session.name,
    intervenantTitle: session.title || APP.SIGNATURE_TITLE_DEFAULT,
    themeModel: '',
    themesText: 'Score : ' + (reponse.score !== null ? reponse.score : '?') + '/' + (reponse.noteMax || '?'),
    remarques: 'Quiz rattaché automatiquement',
    sourceDocFileId: '',
    sourceDocUrl: '',
    pdfFileId: reponse.pdfFileId,
    pdfUrl: reponse.pdfUrl,
    statutPdf: APP.STATUS.PDF_GENERATED,
    statutSignatureIntervenant: APP.STATUS.PENDING,
    dateSignatureIntervenant: '',
    signeIntervenantPar: '',
    statutSignatureAgent: APP.STATUS.PENDING,
    dateSignatureAgent: '',
    signeAgentPar: '',
    dateGeneration: new Date(),
    generePar: session.name,
    majLe: new Date()
  };

  upsertTrackingRecord_(trackingSheet, null, record);
  SpreadsheetApp.flush();

  // Récupérer le rowNumber de la nouvelle ligne
  var lastRow = trackingSheet.getLastRow();
  var repSheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.QUIZ_REPONSES);
  repSheet.getRange(reponse.rowNumber, 15).setValue(lastRow); // TrackingRowNumber

  return { success: true, message: 'Quiz rattaché au suivi — visible dans le dashboard et signeable.' };
}

// ── Import réponses quiz depuis Google Sheet (Google Forms) ──

function importQuizResponsesFromSheet(token, payload) {
  var session = requireSession_(token);
  if (!isAdmin_(session)) throw new Error('Accès réservé admin.');

  var sessionId = String(payload.sessionId || '').trim();
  if (!sessionId) throw new Error('Session de quiz manquante.');

  var sheetUrl = String(payload.sheetUrl || '').trim();
  if (!sheetUrl) throw new Error('URL du Google Sheet requis.');

  // Vérifier que la session existe
  var quizSessions = getQuizSessions_();
  var quizSession = quizSessions.find(function(s) { return s.sessionId === sessionId; });
  if (!quizSession) throw new Error('Session de quiz introuvable.');

  // Charger le modèle pour connaître les questions
  var modeles = getQuizModeles_();
  var modele = modeles.find(function(m) { return m.modeleId === quizSession.modeleId; });
  if (!modele) throw new Error('Modèle de quiz introuvable.');

  var questions = [];
  try { questions = JSON.parse(modele.questionsJson || '[]'); } catch (e) { questions = []; }

  // Ouvrir le Google Sheet source
  var sourceSpreadsheet;
  try {
    var fileId = extractFileIdFromUrl_(sheetUrl);
    sourceSpreadsheet = SpreadsheetApp.openById(fileId);
  } catch (e) {
    throw new Error('Impossible d\'ouvrir le Google Sheet. Vérifiez l\'URL et les droits d\'accès. ' + e.message);
  }

  var sourceSheet = sourceSpreadsheet.getSheets()[0]; // premier onglet
  var lastRow = sourceSheet.getLastRow();
  var startRow = parseInt(payload.startRow) || 2;
  if (lastRow < startRow) throw new Error('Le Sheet source ne contient aucune donnée à partir de la ligne ' + startRow + '.');

  var lastCol = sourceSheet.getLastColumn();
  var dataRange = sourceSheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol);
  var data = dataRange.getValues();

  // Mapper les lettres de colonne vers indices (A=0, B=1, etc.)
  function colLetterToIndex(letter) {
    letter = String(letter || '').trim().toUpperCase();
    if (!letter) return -1;
    var idx = 0;
    for (var i = 0; i < letter.length; i++) {
      idx = idx * 26 + (letter.charCodeAt(i) - 64);
    }
    return idx - 1; // 0-based
  }

  var colMatIdx = colLetterToIndex(payload.colMatricule || 'B');
  var colNomIdx = colLetterToIndex(payload.colNom || 'C');
  var colPrenomIdx = colLetterToIndex(payload.colPrenom || 'D');
  var colGroupeIdx = colLetterToIndex(payload.colGroupe || 'E');

  // Mapper les questions vers leurs colonnes
  var questionMappings = payload.questionMappings || [];
  var qColMap = {};
  questionMappings.forEach(function(m) {
    qColMap[m.questionId] = colLetterToIndex(m.column);
  });

  // Charger les réponses existantes pour détecter les doublons
  var existingReponses = getQuizReponsesBySession_(sessionId);
  var existingKeys = {};
  existingReponses.forEach(function(r) {
    existingKeys[r.matricule] = true;
  });

  var repSheet = getAppSpreadsheet_().getSheetByName(APP.SHEETS.QUIZ_REPONSES);
  var agents = getAgents_();

  var imported = 0;
  var skipped = 0;

  data.forEach(function(row) {
    var matricule = String(row[colMatIdx] || '').trim();
    var nom = String(row[colNomIdx] || '').trim();
    var prenom = colPrenomIdx >= 0 ? String(row[colPrenomIdx] || '').trim() : '';
    var groupe = colGroupeIdx >= 0 ? String(row[colGroupeIdx] || '').trim() : '';

    // Ignorer lignes vides
    var key = matricule || (nom + ' ' + prenom).trim();
    if (!key) return;

    // Vérifier doublon
    if (existingKeys[key]) { skipped++; return; }

    // Déterminer si c'est un agent connu
    var agent = null;
    if (matricule) {
      agent = agents.find(function(a) { return a.matricule === matricule; });
    }
    var isExterne = !agent;

    // Si l'agent est trouvé et nom/prenom vides, les remplir
    if (agent && !nom) {
      var parts = (agent.fullName || '').split(' ');
      prenom = prenom || parts[0] || '';
      nom = parts.slice(1).join(' ') || agent.fullName || '';
      groupe = groupe || agent.sousEquipe || '';
    }

    // Collecter les réponses
    var answers = {};
    questions.forEach(function(q) {
      var colIdx = qColMap[q.id];
      if (colIdx !== undefined && colIdx >= 0 && colIdx < row.length) {
        answers[q.id] = String(row[colIdx] || '').trim();
      }
    });

    // Correction automatique (QCM et vrai/faux)
    var score = 0;
    var noteMax = 0;
    var detailedAnswers = questions.map(function(q) {
      var userAnswer = answers[q.id] || '';
      var points = Number(q.points || 1);
      noteMax += points;
      var correct = null;
      if (q.type === 'qcm' || q.type === 'vrai_faux') {
        correct = String(q.correctAnswer || '').trim().toLowerCase() === String(userAnswer).trim().toLowerCase();
        if (correct) score += points;
      }
      return {
        id: q.id, type: q.type, question: q.question,
        userAnswer: userAnswer, correctAnswer: q.correctAnswer || '',
        points: points, correct: correct
      };
    });

    var reponseId = Utilities.getUuid();
    repSheet.appendRow([
      reponseId,
      sessionId,
      key,
      nom,
      prenom,
      groupe,
      isExterne ? 'OUI' : 'NON',
      JSON.stringify(detailedAnswers),
      score,
      noteMax,
      '', // corrigePar
      '', // corrigeLe
      '', // pdfFileId
      '', // pdfUrl
      '', // trackingRowNumber
      new Date() // soumisLe
    ]);

    existingKeys[key] = true;
    imported++;
  });

  SpreadsheetApp.flush();
  return { success: true, imported: imported, skipped: skipped };
}

function extractFileIdFromUrl_(url) {
  // Supporte les URLs Google Sheets/Drive standards
  var match = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  // Peut-être un ID direct
  if (/^[a-zA-Z0-9_-]{20,}$/.test(String(url).trim())) return url.trim();
  throw new Error('URL invalide : impossible d\'extraire l\'ID du fichier.');
}