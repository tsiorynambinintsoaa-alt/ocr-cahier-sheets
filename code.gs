/**
 * ===================================================================
 *  OCR CAHIER  →  GOOGLE SHEETS
 *  100 % gratuit : Google Apps Script + API Gemini (offre gratuite)
 * ===================================================================
 *
 *  Principe de fiabilité :
 *   1. Les colonnes de la photo sont DÉFINIES À L'AVANCE par l'utilisateur
 *      (1 à 30 colonnes). Date, Heure et Matricule peuvent être choisis
 *      plusieurs fois, ainsi que Code pièce.
 *   2. La FINITION se choisit dans une liste déroulante : ses valeurs sont
 *      les valeurs UNIQUES de la colonne B de la feuille des matricules.
 *   3. Une zone de CONSIGNES LIBRES permet de préciser à l'IA les
 *      particularités de la photo.
 *   4. Chaque valeur est contrôlée par des règles strictes
 *      (date ≤ aujourd'hui, heure dans la plage, code sans espace,
 *      matricule P + 4 chiffres qui existe dans la finition choisie).
 *   5. Toute cellule illisible, hors règle ou incohérente est coloriée
 *      en ROUGE avec une note qui explique pourquoi.
 */

// ---------------------------------------------------------------
//  RÉGLAGES (modifiables ici)
// ---------------------------------------------------------------
const CONFIG = {
  MODELES_A: ['gemini-3.6-flash', 'gemini-3.8-flash', 'gemini-flash-latest', 'gemini-flash-lite-latest'],
  MODELES_B: ['gemini-3.6-flash', 'gemini-3.8-flash', 'gemini-flash-latest', 'gemini-flash-lite-latest'],

  FEUILLE_MATRICULES: 'Finition',
  FEUILLE_RESULTAT: 'Résultat',

  CHAMPS_CONNUS: [
    { id: 'date',      label: 'Date',                      validable: true },
    { id: 'heure',     label: 'Heure',                     validable: true },
    { id: 'matricule', label: 'Matricule',                 validable: true },
    { id: 'code',      label: 'Code pièce',                validable: true },
    { id: 'texte',     label: 'Texte libre',               validable: false },
    { id: 'ignorer',   label: '— Ignorer cette colonne —', validable: false }
  ],

  DOUBLONS_MAX: { date: 30, heure: 30, matricule: 30, code: 30 },
  COLONNES_DEFAUT: ['date', 'heure', 'matricule', 'code'],
  MAX_COLONNES: 30,

  PREFIXES_CODE: ['TS', 'HS', 'HNB'],

  HEURE_MIN: '7:30',
  HEURE_MAX: '16:35',

  MAX_CONSIGNES: 800,

  ROUGE: '#ff4d4d',
  MAX_LIGNES: 150,
  DOUBLE_LECTURE: false,
  TEMPS_MAX_MS: 240000,
  TENTATIVES: 2,
  ATTENTE_REPRISE_MS: 5000,
};

// ---------------------------------------------------------------
//  MENU + INTERFACE
// ---------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('OCR Cahier')
    .addItem('Ouvrir l’interface', 'ouvrirInterface')
    .addToUi();
}

function ouvrirInterface() {
  const html = HtmlService.createHtmlOutputFromFile('Interface')
    .setWidth(1300)
    .setHeight(880);
  SpreadsheetApp.getUi().showModelessDialog(html, 'OCR Cahier → Google Sheets');
}

function getInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reg = lireReglages_();
  let finitions = [];
  try { finitions = chargerMatricules_(ss, reg.feuilleMatricules).finitions; } catch (e) { finitions = []; }
  return {
    feuilles: ss.getSheets().map(function (s) { return s.getName(); }),
    reglages: reg,
    finitions: finitions,
    maxLignes: CONFIG.MAX_LIGNES,
    champsDisponibles: CONFIG.CHAMPS_CONNUS,
    doublonsMax: CONFIG.DOUBLONS_MAX,
    maxColonnes: CONFIG.MAX_COLONNES
  };
}

/** Valeurs uniques de la colonne B de la feuille indiquée (pour la liste déroulante). */
function getFinitions(nomFeuille) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return chargerMatricules_(ss, nomFeuille).finitions;
}

function lireReglages_() {
  const p = PropertiesService.getUserProperties();
  let colonnes;
  try { colonnes = JSON.parse(p.getProperty('COLONNES') || ''); } catch (e) { colonnes = null; }
  if (!Array.isArray(colonnes) || !colonnes.length) colonnes = CONFIG.COLONNES_DEFAUT.slice();
  let prefixes;
  try { prefixes = JSON.parse(p.getProperty('PREFIXES_CODE') || ''); } catch (e) { prefixes = null; }
  if (!Array.isArray(prefixes) || !prefixes.length) prefixes = CONFIG.PREFIXES_CODE.slice();
  return {
    cleDefinie: !!p.getProperty('GEMINI_API_KEY'),
    feuilleMatricules: p.getProperty('FEUILLE_MATRICULES') || CONFIG.FEUILLE_MATRICULES,
    feuilleResultat: p.getProperty('FEUILLE_RESULTAT') || CONFIG.FEUILLE_RESULTAT,
    heureMin: p.getProperty('HEURE_MIN') || CONFIG.HEURE_MIN,
    heureMax: p.getProperty('HEURE_MAX') || CONFIG.HEURE_MAX,
    reporterDate: (p.getProperty('REPORTER_DATE') || 'oui') === 'oui',
    colonnes: colonnes,
    prefixesCode: prefixes,
    consignes: p.getProperty('CONSIGNES') || '',
    finition: p.getProperty('FINITION') || '',
    dispositions: lireDispositions_()
  };
}

/** Dispositions de colonnes mémorisées : { « nom » : ['date','heure',…] } */
function lireDispositions_() {
  try {
    const o = JSON.parse(PropertiesService.getUserProperties().getProperty('DISPOSITIONS') || '{}');
    return (o && typeof o === 'object' && !Array.isArray(o)) ? o : {};
  } catch (e) { return {}; }
}

function saveSettings(s) {
  const p = PropertiesService.getUserProperties();
  const cle = String((s && s.cle) || '').trim();
  if (cle) p.setProperty('GEMINI_API_KEY', cle);
  const fm = String((s && s.feuilleMatricules) || '').trim();
  const fr = String((s && s.feuilleResultat) || '').trim();
  if (fm && fr && fm === fr) {
    throw new Error('La feuille des matricules et la feuille résultat doivent être différentes.');
  }
  if (fm) p.setProperty('FEUILLE_MATRICULES', fm);
  if (fr) p.setProperty('FEUILLE_RESULTAT', fr);
  if (s && Array.isArray(s.colonnes) && s.colonnes.length) {
    p.setProperty('COLONNES', JSON.stringify(validerColonnes_(s.colonnes)));
  }
  if (s && s.prefixesCode != null) {
    p.setProperty('PREFIXES_CODE', JSON.stringify(validerPrefixes_(s.prefixesCode)));
  }
  // --- Dispositions de colonnes nommées (facultatif) ---
  const nomDispo = String((s && s.nomDisposition) || '').trim().slice(0, 40);
  if (nomDispo) {
    if (!s || !Array.isArray(s.colonnes) || !s.colonnes.length) throw new Error('Aucune colonne à mémoriser.');
    const dispos = lireDispositions_();
    dispos[nomDispo] = validerColonnes_(s.colonnes);
    p.setProperty('DISPOSITIONS', JSON.stringify(dispos));
  }
  const suppr = String((s && s.supprimerDisposition) || '').trim();
  if (suppr) {
    const dispos = lireDispositions_();
    delete dispos[suppr];
    p.setProperty('DISPOSITIONS', JSON.stringify(dispos));
  }
  return lireReglages_();
}

function testerConnexion() {
  const cle = PropertiesService.getUserProperties().getProperty('GEMINI_API_KEY');
  if (!cle) throw new Error('Aucune clé d’accès enregistrée. Saisissez-la dans Réglages puis enregistrez.');
  const modeles = [CONFIG.MODELES_A[0]];
  return modeles.map(function (m) {
    try {
      const r = UrlFetchApp.fetch(urlModele_(m), {
        method: 'post',
        contentType: 'application/json',
        headers: { 'x-goog-api-key': cle },
        payload: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: 'Réponds uniquement : OK' }] }] }),
        muteHttpExceptions: true
      });
      const code = r.getResponseCode();
      if (code === 200) return { modele: m, ok: true, statut: 'disponible' };
      return { modele: m, ok: false, statut: 'HTTP ' + code + ' — ' + messageErreur_(r) };
    } catch (e) {
      return { modele: m, ok: false, statut: String(e.message || e) };
    }
  });
}

// ---------------------------------------------------------------
//  CODES PIÈCE POSSIBLES (préfixes)
// ---------------------------------------------------------------
function validerPrefixes_(brut) {
  const liste = Array.isArray(brut) ? brut : String(brut == null ? '' : brut).split(/[\s,;]+/);
  const out = [];
  liste.forEach(function (x) {
    const pre = String(x).trim().toUpperCase();
    if (!pre) return;
    if (!/^[A-Z]{1,4}$/.test(pre)) {
      throw new Error('Préfixe de code invalide : « ' + pre + ' » (1 à 4 lettres, séparés par des virgules).');
    }
    if (out.indexOf(pre) < 0) out.push(pre);
  });
  return out.length ? out : CONFIG.PREFIXES_CODE.slice();
}

// ---------------------------------------------------------------
//  COLONNES DÉFINIES À L'AVANCE
// ---------------------------------------------------------------
function validerColonnes_(colonnes) {
  if (!Array.isArray(colonnes) || !colonnes.length) {
    throw new Error('Définissez au moins une colonne.');
  }
  if (colonnes.length > CONFIG.MAX_COLONNES) {
    throw new Error('Maximum ' + CONFIG.MAX_COLONNES + ' colonnes.');
  }
  const idsConnus = CONFIG.CHAMPS_CONNUS.map(function (c) { return c.id; });
  const compte = {};
  let utile = false;
  colonnes.forEach(function (c) {
    if (idsConnus.indexOf(c) < 0) throw new Error('Type de colonne inconnu : « ' + c + ' ».');
    if (c !== 'ignorer') utile = true;
    compte[c] = (compte[c] || 0) + 1;
    const max = CONFIG.MAX_COLONNES;
    if (compte[c] > max) {
      const def = CONFIG.CHAMPS_CONNUS.filter(function (x) { return x.id === c; })[0];
      throw new Error('La colonne « ' + (def ? def.label : c) + ' » est choisie trop de fois (maximum ' + max + ').');
    }
  });
  if (!utile) throw new Error('Au moins une colonne utile est requise (pas uniquement « Ignorer »).');
  return colonnes.slice();
}

function colonnesAvecCle_(colonnes) {
  const totaux = {}, compteurs = {};
  colonnes.forEach(function (c) { if (c !== 'ignorer') totaux[c] = (totaux[c] || 0) + 1; });
  return colonnes.map(function (champ, i) {
    const def = CONFIG.CHAMPS_CONNUS.filter(function (c) { return c.id === champ; })[0];
    if (!def) throw new Error('Type de colonne inconnu : ' + champ);
    if (champ === 'ignorer') return { position: i + 1, champ: champ, cle: null, libelle: def.label };
    compteurs[champ] = (compteurs[champ] || 0) + 1;
    let cle = champ, libelle = def.label;
    if (totaux[champ] > 1) {
      cle = champ + '_' + compteurs[champ];
      libelle = def.label + ' ' + compteurs[champ];
    }
    return { position: i + 1, champ: champ, cle: cle, libelle: libelle };
  });
}

// ---------------------------------------------------------------
//  TRAITEMENT PRINCIPAL
// ---------------------------------------------------------------
/**
 * p = { image (base64), mime, colonnes, prefixesCode, consignes,
 *       finition (valeur choisie dans la liste, si colonne Matricule),
 *       nbLignes, heureMin, heureMax (si colonne Heure), reporterDate }
 */
function traiterImage(p) {
  p = p || {};
  const debutTraitement = Date.now();
  // Nombre de lignes FACULTATIF : vide → on prend toutes les lignes visibles sur la photo.
  const nbBrut = String(p.nbLignes == null ? '' : p.nbLignes).trim();
  let nb = null;
  if (nbBrut !== '') {
    nb = parseInt(nbBrut, 10);
    if (!(nb >= 1 && nb <= CONFIG.MAX_LIGNES)) {
      throw new Error('Nombre de lignes invalide (entre 1 et ' + CONFIG.MAX_LIGNES + ', ou laissez vide).');
    }
  }
  if (!p.image) throw new Error('Aucune image reçue.');

  const props = PropertiesService.getUserProperties();
  const cle = props.getProperty('GEMINI_API_KEY');
  if (!cle) throw new Error('Clé d’accès manquante : saisissez-la dans Réglages puis enregistrez.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reg = lireReglages_();
  if (reg.feuilleMatricules === reg.feuilleResultat) {
    throw new Error('La feuille des matricules et la feuille résultat doivent être différentes (Réglages).');
  }

  const colonnesBrutes = (Array.isArray(p.colonnes) && p.colonnes.length) ? p.colonnes : reg.colonnes;
  const colonnes = validerColonnes_(colonnesBrutes);
  const defsColonnes = colonnesAvecCle_(colonnes);
  const champsUtiles = defsColonnes.filter(function (d) { return d.cle; });
  const CHAMPS = champsUtiles.map(function (d) { return d.cle; });
  const EN_TETES = champsUtiles.map(function (d) { return d.libelle; });
  const typeParCle = {};
  defsColonnes.forEach(function (d) { if (d.cle) typeParCle[d.cle] = d.champ; });
  const aMatricule = defsColonnes.some(function (d) { return d.champ === 'matricule'; });
  const aCode = defsColonnes.some(function (d) { return d.champ === 'code'; });
  const aHeure = defsColonnes.some(function (d) { return d.champ === 'heure'; });

  const prefixesCode = (p.prefixesCode != null) ? validerPrefixes_(p.prefixesCode) : reg.prefixesCode;
  const consignes = String(p.consignes || '').replace(/\s+$/, '').slice(0, CONFIG.MAX_CONSIGNES);

  let heureMin = minutesDepuisTexte_(p.heureMin);
  let heureMax = minutesDepuisTexte_(p.heureMax);
  if (aHeure) {
    if (heureMin === null || heureMax === null) throw new Error('Heures min/max invalides (format 7:30).');
    if (heureMin >= heureMax) throw new Error('L’heure minimale doit être avant l’heure maximale.');
  } else if (heureMin === null || heureMax === null || heureMin >= heureMax) {
    heureMin = minutesDepuisTexte_(CONFIG.HEURE_MIN);
    heureMax = minutesDepuisTexte_(CONFIG.HEURE_MAX);
  }

  // --- Finition (obligatoire si au moins une colonne Matricule existe) ---
  let finition = null, parFinition = null, setFinition = null;
  if (aMatricule) {
    finition = String(p.finition == null ? '' : p.finition).trim();
    if (!finition) throw new Error('Choisissez la finition atelier dans la liste.');
    const matr = chargerMatricules_(ss, reg.feuilleMatricules);
    if (matr.finitions.indexOf(finition) < 0) {
      throw new Error('La finition « ' + finition + ' » n’existe pas dans la colonne B de la feuille « ' +
        reg.feuilleMatricules + ' ».');
    }
    setFinition = {};
    Object.keys(matr.parFinition).forEach(function (mat) {
      if (matr.parFinition[mat].indexOf(finition) >= 0) setFinition[mat] = true;
    });
    parFinition = matr.parFinition;
  }

  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  const d = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd').split('-').map(Number);
  const ctx = {
    aujourdhui: { a: d[0], m: d[1], j: d[2], texte: pad2_(d[2]) + '/' + pad2_(d[1]) + '/' + d[0] },
    heureMin: heureMin,
    heureMax: heureMax,
    prefixes: aCode ? prefixesCode : [],
    finition: finition,
    parFinition: parFinition,
    setFinition: setFinition
  };

  const lect = lectureDouble_(cle, { b64: p.image, mime: p.mime || 'image/jpeg' }, nb,
    debutTraitement + CONFIG.TEMPS_MAX_MS, defsColonnes, CHAMPS, consignes);

  const valParCle = validateursParCle_(typeParCle);
  const res = construireResultat_(lect.A.lignes, lect.B.lignes, ctx, p.reporterDate !== false, CHAMPS, valParCle, typeParCle);

  const info = ecrireResultat_(ss, reg.feuilleResultat, res, CHAMPS, EN_TETES, typeParCle);

  props.setProperty('HEURE_MIN', fmtMin_(heureMin));
  props.setProperty('HEURE_MAX', fmtMin_(heureMax));
  props.setProperty('REPORTER_DATE', p.reporterDate === false ? 'non' : 'oui');
  props.setProperty('COLONNES', JSON.stringify(colonnes));
  props.setProperty('CONSIGNES', consignes);
  if (finition) props.setProperty('FINITION', finition);
  if (aCode) props.setProperty('PREFIXES_CODE', JSON.stringify(prefixesCode));

  info.modeleA = lect.A.modele;
  info.modeleB = lect.B.modele;
  info.finition = finition;
  info.nbFourni = nb !== null;
  return info;
}

// ---------------------------------------------------------------
//  LISTE DES MATRICULES  (colonne A = matricule, colonne B = finition)
// ---------------------------------------------------------------
function comparerFinitions_(a, b) {
  const na = /^-?\d+(\.\d+)?$/.test(a), nb = /^-?\d+(\.\d+)?$/.test(b);
  if (na && nb) return parseFloat(a) - parseFloat(b);
  if (na) return -1;
  if (nb) return 1;
  return a.localeCompare(b, 'fr', { numeric: true, sensitivity: 'base' });
}

function chargerMatricules_(ss, nom) {
  const sh = nom ? ss.getSheetByName(nom) : null;
  if (!sh) {
    throw new Error('Feuille des matricules introuvable (« ' + (nom || '') + ' »). Choisissez-la dans Réglages.');
  }
  const n = sh.getLastRow();
  if (n < 1) throw new Error('La feuille « ' + nom + ' » est vide.');
  const vals = sh.getRange(1, 1, n, 2).getValues();
  const parFinition = {};
  const uniques = {};
  let nb = 0;
  vals.forEach(function (r) {
    const mm = String(r[0]).match(/\d+/);
    if (!mm || mm[0].length > 4) return;          // ignore l'en-tête et les lignes sans matricule
    const fin = String(r[1] == null ? '' : r[1]).trim();
    if (!fin) return;
    const mat = 'P' + ('0000' + mm[0]).slice(-4);
    if (!parFinition[mat]) parFinition[mat] = [];
    if (parFinition[mat].indexOf(fin) < 0) parFinition[mat].push(fin);
    uniques[fin] = true;
    nb++;
  });
  const finitions = Object.keys(uniques).sort(comparerFinitions_);
  if (!finitions.length) {
    throw new Error('Aucune finition trouvée dans la feuille « ' + nom +
      ' » (colonne A = matricule, colonne B = finition atelier).');
  }
  return { parFinition: parFinition, finitions: finitions, nb: nb };
}

// ---------------------------------------------------------------
//  APPEL À GEMINI
// ---------------------------------------------------------------
function urlModele_(modele) {
  return 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(modele) + ':generateContent';
}

function schemaReponse_(champs) {
  const properties = {};
  champs.forEach(function (c) { properties[c] = { type: 'STRING' }; });
  return {
    type: 'OBJECT',
    properties: {
      lignes: {
        type: 'ARRAY',
        items: {
          type: 'OBJECT',
          properties: properties,
          required: champs,
          propertyOrdering: champs
        }
      }
    },
    required: ['lignes']
  };
}

function consigneSysteme_(defsColonnes, consignes) {
  const descriptions = {
    date: 'telle qu’écrite (ex. « 03/10 »). Elle n’est souvent écrite que sur la première ligne : ne la recopie PAS sur les lignes suivantes, laisse "" si la case est vide.',
    heure: 'telle qu’écrite (ex. « 7h30 », « 10h06 »). Si seulement un tiret ou « -1- » est écrit, mets "-".',
    matricule: 'uniquement les chiffres, sans lettre devant (la lettre P, D ou H écrite devant n’est pas à recopier).',
    code: 'lettres puis chiffres tels qu’écrits (ex. « TS 189 », « HNB 139 », « HS 17 »).',
    texte: 'telle qu’écrite, sans interprétation.'
  };
  const listeColonnes = defsColonnes.map(function (d) {
    if (d.champ === 'ignorer') {
      return '  - Colonne ' + d.position + ' : à IGNORER totalement (ne doit PAS figurer dans le JSON).';
    }
    return '  - Colonne ' + d.position + ' : « ' + d.libelle + ' » (clé JSON « ' + d.cle + ' ») — ' +
      (descriptions[d.champ] || '');
  }).join('\n');
  const clesUtiles = defsColonnes.filter(function (d) { return d.cle; }).map(function (d) { return d.cle; });
  const exempleJson = clesUtiles.map(function (c) { return '"' + c + '":""'; }).join(',');

  const lignes = [
    'Tu es un expert en transcription de tableaux manuscrits (cahier d’un atelier).',
    'La photo montre un tableau dont les colonnes, de gauche à droite, sont EXACTEMENT :',
    listeColonnes,
    '',
    'Tu transcris EXACTEMENT ce qui est écrit, sans corriger, sans deviner, sans compléter.',
    '',
    'RÈGLES',
    '1. Réponds uniquement par un objet JSON : {"lignes":[{' + exempleJson + '}, ...]}.',
    '2. Le tableau « lignes » contient une entrée par ligne du tableau, de haut en bas, dans l’ordre de la page ' +
      '(si un nombre de lignes précis est demandé dans le message, respecte-le EXACTEMENT). ' +
      'Une ligne = une entrée du tableau (toutes les colonnes alignées horizontalement). ' +
      'Une ligne où rien n’est écrit = toutes les clés avec une chaîne vide.',
    '3. N’inclus AUCUNE clé pour les colonnes marquées « à IGNORER » ci-dessus.',
    '4. Même en cas de doute, écris toujours la valeur ou le caractère le plus probable afin que la cellule contienne une valeur. ' +
      'Ne mets « ? » que si le caractère est totalement impossible à distinguer.',
    '5. Si une valeur est barrée et remplacée par une autre, donne la valeur qui n’est pas barrée ; si c’est impossible à lire, mets « ? ».',
    '6. Ignore ce qui est hors tableau (annotations dans la marge, calculs en haut de page).',
    '7. L’écriture est irrégulière : compare chaque chiffre avec les autres occurrences sur la page avant de décider.'
  ];
  if (consignes) {
    lignes.push('',
      'CONSIGNES PARTICULIÈRES DE L’UTILISATEUR POUR CETTE PHOTO',
      '(elles précisent les règles ci-dessus : applique-les en priorité) :',
      consignes);
  }
  return lignes.join('\n');
}

function consigneUtilisateur_(nb, variante, defsColonnes) {
  if (nb === null) {
    return 'Transcris toutes les lignes visibles du tableau de cette page, de haut en bas, dans l’ordre. ' +
      'N’invente aucune ligne qui ne figure pas sur la photo ; garde les lignes vides situées entre deux lignes remplies. ' +
      'Retourne le JSON avec une entrée dans « lignes » par ligne du tableau.';
  }
  if (variante === 'B') {
    const ordre = defsColonnes.filter(function (d) { return d.cle; })
      .slice().reverse().map(function (d) { return d.libelle; }).join(', puis ');
    return 'Procède en deux temps : lis d’abord chaque colonne séparément de haut en bas ' +
      '(' + ordre + '), puis assemble les lignes. ' +
      'Retourne le JSON avec EXACTEMENT ' + nb + ' éléments dans « lignes », alignés sur les lignes de la page.';
  }
  return 'Transcris les ' + nb + ' lignes du tableau de cette page, de haut en bas. ' +
    'Retourne le JSON avec EXACTEMENT ' + nb + ' éléments dans « lignes » (lignes vides comprises).';
}

function requeteGemini_(modele, cle, image, nb, variante, avecSchema, defsColonnes, champs, consignes) {
  const body = {
    systemInstruction: { parts: [{ text: consigneSysteme_(defsColonnes, consignes) }] },
    contents: [{
      role: 'user',
      parts: [
        { inlineData: { mimeType: image.mime, data: image.b64 } },
        { text: consigneUtilisateur_(nb, variante, defsColonnes) }
      ]
    }],
    generationConfig: {
      responseMimeType: 'application/json',
      temperature: 0,
      candidateCount: 1
    }
  };
  if (avecSchema) body.generationConfig.responseSchema = schemaReponse_(champs);
  return {
    url: urlModele_(modele),
    method: 'post',
    contentType: 'application/json',
    followRedirects: true,
    headers: { 'x-goog-api-key': cle },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  };
}

function messageErreur_(resp) {
  try {
    const j = JSON.parse(resp.getContentText());
    if (j && j.error && j.error.message) return String(j.error.message).slice(0, 300);
  } catch (e) { /* ignore */ }
  return String(resp.getContentText() || '').slice(0, 200);
}

function traiterReponse_(resp, nb, champs) {
  const code = resp.getResponseCode();
  if (code !== 200) {
    const msg = messageErreur_(resp);
    let type = 'contenu';
    if (code === 401 || code === 403 || /API key not valid|API_KEY_INVALID|PERMISSION_DENIED/i.test(msg)) type = 'fatal';
    else if (code === 404) type = 'modele';
    else if (code === 429) type = 'quota';
    else if (code === 500 || code === 502 || code === 503 || code === 504) type = 'attente';
    else if (code === 400) type = /schema/i.test(msg) ? 'schema' : 'modele';
    return { ok: false, type: type, erreur: 'HTTP ' + code + ' — ' + msg };
  }
  let data;
  try { data = JSON.parse(resp.getContentText()); }
  catch (e) { return { ok: false, type: 'contenu', erreur: 'réponse illisible' }; }

  const cand = data && data.candidates && data.candidates[0];
  const parts = cand && cand.content && cand.content.parts;
  if (!parts || !parts.length) {
    const raison = (cand && cand.finishReason) || (data && data.promptFeedback && data.promptFeedback.blockReason) || 'inconnue';
    return { ok: false, type: 'contenu', erreur: 'réponse vide (' + raison + ')' };
  }
  const texte = parts.filter(function (x) { return !x.thought && typeof x.text === 'string'; })
    .map(function (x) { return x.text; }).join('');
  let obj;
  try {
    obj = JSON.parse(texte.replace(/^\s*```(?:json)?/i, '').replace(/```\s*$/, '').trim());
  } catch (e) {
    return { ok: false, type: 'contenu', erreur: 'JSON invalide dans la réponse' };
  }
  const arr = Array.isArray(obj) ? obj : (obj && obj.lignes);
  if (!Array.isArray(arr)) return { ok: false, type: 'contenu', erreur: 'champ « lignes » absent' };
  if (nb !== null && arr.length !== nb) {
    return { ok: false, type: 'contenu', erreur: arr.length + ' lignes lues au lieu de ' + nb };
  }
  if (nb === null && !arr.length) {
    return { ok: false, type: 'contenu', erreur: 'aucune ligne lue' };
  }
  const lignes = (nb === null ? arr.slice(0, CONFIG.MAX_LIGNES) : arr).map(function (l) {
    l = l || {};
    const out = {};
    champs.forEach(function (c) { out[c] = String(l[c] == null ? '' : l[c]); });
    return out;
  });
  return { ok: true, lignes: lignes };
}

function fetchSecurise_(req) {
  const params = {};
  Object.keys(req).forEach(function (k) { if (k !== 'url') params[k] = req[k]; });
  try { return UrlFetchApp.fetch(req.url, params); }
  catch (e) { return { getResponseCode: function () { return 503; }, getContentText: function () { return String(e.message || e); } }; }
}

function verifierBudget_(deadline) {
  if (deadline && Date.now() >= deadline) {
    throw new Error('Durée maximale presque atteinte. Relancez le traitement avec la même image.');
  }
}

function lireAvecReprises_(lecture, cle, image, nb, premiere, deadline, defsColonnes, champs, consignes) {
  const journal = [];
  for (let m = 0; m < lecture.modeles.length; m++) {
    const modele = lecture.modeles[m];
    let avecSchema = true;
    for (let t = 0; t < CONFIG.TENTATIVES; t++) {
      verifierBudget_(deadline);
      let r;
      if (m === 0 && t === 0 && premiere) {
        r = premiere;
      } else {
        r = traiterReponse_(
          fetchSecurise_(requeteGemini_(modele, cle, image, nb, lecture.variante, avecSchema, defsColonnes, champs, consignes)),
          nb, champs);
      }
      if (r.ok) return { lignes: r.lignes, modele: modele };
      journal.push(r.erreur);
      if (r.type === 'fatal') {
        throw new Error('Clé d’accès refusée. Vérifiez la clé dans Réglages.');
      }
      if (r.type === 'quota') {
        throw new Error('Quota de lecture dépassé (HTTP 429). Aucune nouvelle tentative n’est effectuée. Réessayez plus tard.');
      }
      if (r.type === 'schema') { avecSchema = false; continue; }
      if (r.type === 'modele') break;
      if (r.type === 'attente' && t + 1 < CONFIG.TENTATIVES) {
        const attente = Math.min(CONFIG.ATTENTE_REPRISE_MS * (t + 1),
          Math.max(0, (deadline || Date.now() + CONFIG.ATTENTE_REPRISE_MS) - Date.now() - 250));
        if (attente > 0) Utilities.sleep(attente);
      }
    }
  }
  throw new Error('Lecture impossible — ' + journal.slice(-3).join(' | '));
}

function lectureDouble_(cle, image, nb, deadline, defsColonnes, champs, consignes) {
  const lectures = [
    { nom: 'A', modeles: CONFIG.MODELES_A, variante: 'A' },
    { nom: 'B', modeles: CONFIG.MODELES_B, variante: 'B' }
  ];
  if (!CONFIG.DOUBLE_LECTURE) {
    const rapide = lireAvecReprises_(lectures[0], cle, image, nb, null, deadline, defsColonnes, champs, consignes);
    return {
      A: rapide,
      B: { lignes: rapide.lignes, modele: rapide.modele + ' (validation locale)' }
    };
  }
  verifierBudget_(deadline);
  let premieres = [null, null];
  try {
    const resps = UrlFetchApp.fetchAll(lectures.map(function (l) {
      return requeteGemini_(l.modeles[0], cle, image, nb, l.variante, true, defsColonnes, champs, consignes);
    }));
    premieres = resps.map(function (r) { return traiterReponse_(r, nb, champs); });
  } catch (e) {
    premieres = [null, null];
  }
  const sortie = {};
  lectures.forEach(function (l, i) {
    sortie[l.nom] = lireAvecReprises_(l, cle, image, nb, premieres[i], deadline, defsColonnes, champs, consignes);
  });
  return sortie;
}

// ---------------------------------------------------------------
//  RÈGLES DE VALIDATION  (fonctions pures)
// ---------------------------------------------------------------
function pad2_(n) { return (n < 10 ? '0' : '') + n; }
function fmtMin_(min) { return Math.floor(min / 60) + ':' + pad2_(min % 60); }
function texteNet_(raw) { return String(raw == null ? '' : raw).replace(/\s+/g, ' ').trim(); }
function estVide_(s) { return /^[\s\-–—_.·]*$/.test(s); }
function resVide_() { return { vide: true, texte: '', ok: true, raison: '' }; }
function resInvalide_(texte, raison) { return { vide: false, texte: texte, ok: false, raison: raison }; }

function minutesDepuisTexte_(txt) {
  const m = String(txt == null ? '' : txt).trim().toLowerCase().match(/^(\d{1,2})\s*[h:.]\s*(\d{2})$/);
  if (!m) return null;
  const h = parseInt(m[1], 10), mi = parseInt(m[2], 10);
  if (h > 23 || mi > 59) return null;
  return h * 60 + mi;
}

function validerHeure_(raw, ctx) {
  const s = texteNet_(raw).toLowerCase();
  if (estVide_(s)) return resVide_();
  const compact = s.replace(/\s+/g, '');
  if (s.indexOf('?') >= 0) return resInvalide_(compact, 'caractère illisible');
  const m = s.match(/^(\d{1,2})\s*[h:.,]\s*(\d{2})?$/);
  if (!m) return resInvalide_(compact, 'format d’heure non reconnu');
  const h = parseInt(m[1], 10), mi = m[2] ? parseInt(m[2], 10) : 0;
  if (h > 23 || mi > 59) return resInvalide_(compact, 'heure impossible');
  const minutes = h * 60 + mi;
  const r = { vide: false, texte: h + ':' + pad2_(mi), serial: minutes / 1440, ok: true, raison: '' };
  if (minutes < ctx.heureMin || minutes > ctx.heureMax) {
    r.ok = false;
    r.raison = 'heure hors plage (' + fmtMin_(ctx.heureMin) + ' – ' + fmtMin_(ctx.heureMax) + ')';
  }
  return r;
}

function validerDate_(raw, ctx) {
  const s = texteNet_(raw);
  if (estVide_(s)) return resVide_();
  const compact = s.replace(/\s+/g, '');
  if (s.indexOf('?') >= 0) return resInvalide_(compact, 'caractère illisible');
  const m = s.match(/^(\d{1,2})\s*[\/.\-\s]\s*(\d{1,2})(?:\s*[\/.\-\s]\s*(\d{4}|\d{2}))?$/);
  if (!m) return resInvalide_(compact, 'format de date non reconnu');
  const jj = parseInt(m[1], 10), mm = parseInt(m[2], 10);
  const aa = m[3] ? (m[3].length === 2 ? 2000 + parseInt(m[3], 10) : parseInt(m[3], 10)) : ctx.aujourdhui.a;
  const dt = new Date(Date.UTC(aa, mm - 1, jj));
  if (mm < 1 || mm > 12 || jj < 1 || dt.getUTCMonth() !== mm - 1 || dt.getUTCDate() !== jj) {
    return resInvalide_(compact, 'date impossible');
  }
  const t = ctx.aujourdhui;
  const r = {
    vide: false,
    texte: pad2_(jj) + '/' + pad2_(mm) + '/' + aa,
    serial: Math.round((Date.UTC(aa, mm - 1, jj) - Date.UTC(1899, 11, 30)) / 86400000),
    ok: true,
    raison: ''
  };
  if (aa * 10000 + mm * 100 + jj > t.a * 10000 + t.m * 100 + t.j) {
    r.ok = false;
    r.raison = 'date postérieure à aujourd’hui (' + t.texte + ')';
  }
  return r;
}

function validerCode_(raw, ctx) {
  const s = texteNet_(raw);
  if (estVide_(s)) return resVide_();
  const code = s.toUpperCase().replace(/\s+/g, '');
  if (code.indexOf('?') >= 0) return resInvalide_(code, 'caractère illisible');
  const m = code.match(/^([A-Z]{1,4})(\d{1,4})$/);
  if (!m) return resInvalide_(code, 'format inattendu (lettres puis chiffres, ex. TS189)');
  const r = { vide: false, texte: code, ok: true, raison: '' };
  if (ctx.prefixes && ctx.prefixes.length && ctx.prefixes.indexOf(m[1]) < 0) {
    r.ok = false;
    r.raison = 'préfixe « ' + m[1] + ' » inconnu (connus : ' + ctx.prefixes.join(', ') + ')';
  }
  return r;
}

function voisins_(texte, setFinition) {
  const out = [];
  Object.keys(setFinition || {}).forEach(function (cand) {
    if (cand.length !== texte.length) return;
    let diff = 0;
    for (let i = 0; i < cand.length; i++) if (cand[i] !== texte[i]) diff++;
    if (diff === 1) out.push(cand);
  });
  return out.slice(0, 3);
}

function validerMatricule_(raw, ctx) {
  const s = texteNet_(raw).toUpperCase();
  if (estVide_(s)) return resVide_();
  const compact = s.replace(/\s+/g, '');
  if (compact.indexOf('?') >= 0) return resInvalide_(compact, 'caractère illisible');
  const m = compact.match(/^([A-Z]{0,2})(\d+)$/);
  if (!m) return resInvalide_(compact, 'caractères inattendus dans le matricule');
  if (m[2].length > 4) return resInvalide_('P' + m[2], 'plus de 4 chiffres');
  const texte = 'P' + ('0000' + m[2]).slice(-4);
  const r = { vide: false, texte: texte, ok: true, raison: '' };
  if (ctx.parFinition) {
    const f = ctx.parFinition[texte];
    if (!f || f.indexOf(ctx.finition) < 0) {
      r.ok = false;
      r.raison = 'absent de la finition ' + ctx.finition;
      if (f && f.length) {
        r.raison += ' (existe en finition ' + f.join(', ') + ')';
      } else {
        const v = voisins_(texte, ctx.setFinition);
        if (v.length) r.raison += ' ; proche : ' + v.join(', ');
      }
    }
  }
  return r;
}

function validerTexte_(raw) {
  const s = texteNet_(raw);
  if (estVide_(s)) return resVide_();
  if (s.indexOf('?') >= 0) return resInvalide_(s, 'caractère illisible');
  return { vide: false, texte: s, ok: true, raison: '' };
}

const VALIDATEURS_ = {
  date: validerDate_,
  code: validerCode_,
  heure: validerHeure_,
  matricule: validerMatricule_
};

function validateursParCle_(typeParCle) {
  const map = {};
  Object.keys(typeParCle).forEach(function (cle) {
    map[cle] = VALIDATEURS_[typeParCle[cle]] || validerTexte_;
  });
  return map;
}

function evaluerChamp_(champ, rawA, rawB, ctx, valParCle) {
  const f = valParCle[champ] || validerTexte_;
  const a = f(rawA, ctx), b = f(rawB, ctx);
  if (a.vide && b.vide) return { vide: true, texte: '', rouge: false, note: '' };
  let c, rouge = false, note = '';
  if (!a.vide && !b.vide && a.texte === b.texte) {
    c = a;
    if (!a.ok) { rouge = true; note = a.raison; }
  } else {
    if (a.vide) c = b;
    else if (b.vide) c = a;
    else if (a.ok && !b.ok) c = a;
    else if (b.ok && !a.ok) c = b;
    else c = a;
    rouge = true;
    note = 'Les deux lectures diffèrent — A : « ' + (a.vide ? 'vide' : a.texte) +
      ' » / B : « ' + (b.vide ? 'vide' : b.texte) + ' »';
    if (!c.ok) note += ' — ' + c.raison;
  }
  const out = { vide: false, texte: c.texte, rouge: rouge, note: note };
  if (c.serial !== undefined) out.serial = c.serial;
  return out;
}

function construireResultat_(A, B, ctx, reporterDate, champs, valParCle, typeParCle) {
  const lignes = [], problemes = [];
  let nbVides = 0;
  const clesDate = champs.filter(function (ch) { return typeParCle[ch] === 'date'; });
  const derniereParDate = {};
  for (let i = 0; i < A.length; i++) {
    const cellules = {};
    let toutVide = true;
    champs.forEach(function (ch) {
      cellules[ch] = evaluerChamp_(ch, A[i][ch], B[i][ch], ctx, valParCle);
      if (!cellules[ch].vide) toutVide = false;
    });
    if (toutVide) {
      lignes.push({ vide: true, cellules: cellules });
      nbVides++;
      continue;
    }
    clesDate.forEach(function (ch) {
      if (cellules[ch].vide && reporterDate && derniereParDate[ch]) {
        const src = derniereParDate[ch].cellule;
        cellules[ch] = {
          vide: false, texte: src.texte, serial: src.serial, rouge: src.rouge,
          note: src.rouge ? 'Date reportée de la ligne ' + derniereParDate[ch].ligne + ' (incertaine) — ' + src.note : ''
        };
      } else if (!cellules[ch].vide) {
        derniereParDate[ch] = { cellule: cellules[ch], ligne: i + 1 };
      }
    });
    champs.forEach(function (ch) {
      const c = cellules[ch];
      if (c.vide) { c.rouge = true; c.note = 'Valeur absente ou non écrite'; }
      if (c.rouge) problemes.push({ index: i, champ: ch, texte: c.texte, note: c.note });
    });
    lignes.push({ vide: false, cellules: cellules });
  }
  return { lignes: lignes, problemes: problemes, nbVides: nbVides };
}

// ---------------------------------------------------------------
//  ÉCRITURE DANS GOOGLE SHEETS
// ---------------------------------------------------------------
function ecrireResultat_(ss, nomFeuille, res, champs, enTetes, typeParCle) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const sh = ss.getSheetByName(nomFeuille) || ss.insertSheet(nomFeuille);
    const nCol = champs.length;
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, nCol).setValues([enTetes]).setFontWeight('bold');
      sh.setFrozenRows(1);
    } else if (sh.getLastColumn() !== nCol) {
      throw new Error('La feuille « ' + nomFeuille + ' » contient déjà ' + sh.getLastColumn() +
        ' colonne(s) alors que la disposition actuelle en définit ' + nCol +
        '. Utilisez une feuille vide ou reprenez la même disposition de colonnes.');
    }
    const debut = sh.getLastRow() + 1;
    const n = res.lignes.length;
    const FORMATS = { date: 'dd/mm/yyyy', heure: 'H:mm' };

    const valeurs = [], formats = [], couleurs = [], notes = [];
    res.lignes.forEach(function (l) {
      const v = [], f = [], c = [], nt = [];
      champs.forEach(function (ch) {
        const cell = l.cellules[ch];
        if (cell.vide) {
          v.push(''); f.push('@');
          c.push(cell.rouge ? CONFIG.ROUGE : null);
          nt.push(cell.rouge ? cell.note : '');
          return;
        }
        const fmt = FORMATS[typeParCle[ch]];
        if (cell.serial !== undefined && fmt) { v.push(cell.serial); f.push(fmt); }
        else { v.push(cell.texte); f.push('@'); }
        c.push(cell.rouge ? CONFIG.ROUGE : null);
        nt.push(cell.rouge ? cell.note : '');
      });
      valeurs.push(v); formats.push(f); couleurs.push(c); notes.push(nt);
    });

    const rng = sh.getRange(debut, 1, n, nCol);
    rng.setNumberFormats(formats);
    rng.setValues(valeurs);
    rng.setBackgrounds(couleurs);
    rng.setNotes(notes);
    SpreadsheetApp.flush();
    sh.activate();
    rng.activate();

    const details = res.problemes.slice(0, 200).map(function (p) {
      return {
        ligne: debut + p.index,
        champ: enTetes[champs.indexOf(p.champ)],
        lu: p.texte,
        raison: p.note
      };
    });
    let videsFin = 0;
    for (let i = res.lignes.length - 1; i >= 0 && res.lignes[i].vide; i--) videsFin++;

    return {
      feuille: nomFeuille,
      debut: debut,
      fin: debut + n - 1,
      nbLignes: n,
      nbVides: res.nbVides,
      videsFin: videsFin,
      nbRouges: res.problemes.length,
      details: details
    };
  } finally {
    lock.releaseLock();
  }
}