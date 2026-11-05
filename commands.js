/*
 * Bouton « Charte », Le Froid Pecomark, Service Informatique, v1.2
 * Sur le message en cours de rédaction :
 *   1. applique Century Gothic 10 pt #00447A à la partie rédigée
 *   2. laisse intacts la signature et le message cité (réponse, transfert)
 *   3. signature : si le message porte déjà la signature Outlook, la mémorise sur ce poste
 *      (texte et images) ; sinon, y remet la dernière signature mémorisée
 * Décision du 12/10/2026, Flora Llopis, responsable communication.
 * La signature n'est jamais recréée : c'est une copie de celle insérée par Outlook.
 */

const POLICE = "'Century Gothic', Arial, sans-serif";
const TAILLE = "10pt";
const COULEUR = "#00447A";

// Signature propre à l'expéditeur, puis tous les blocs à ne jamais retoucher
const SEL_SIGNATURE_PROPRE = "#Signature, [id^='Signature']";
const SEL_SIGNATURE = SEL_SIGNATURE_PROPRE + ", #x_Signature, [id^='x_Signature']";
const SEL_CITATION = "#appendonsend, #divRplyFwdMsg, [id^='divRplyFwdMsg'], #mail-editor-reference-message-container";
const SEL_PROTEGE = SEL_SIGNATURE + ", " + SEL_CITATION;

// Mémoire de la signature, propre au poste (stockage du navigateur intégré d'Outlook)
const CLE_SIGNATURE = "lfCharte.signature.v1";
const PREFIXE_IMAGE = "lf-img:";

Office.onReady(() => {});

function habillerCharte(event) {
  const item = Office.context.mailbox.item;
  traiter(item)
    .then((message) => notifier(item, message, false, event))
    .catch((err) => notifier(item, "Bouton Charte : " + (err && err.message ? err.message : err), true, event));
}

async function traiter(item) {
  const corps = await lireCorps(item);
  let message = "Mise en forme charte appliquée";
  let signature = "";

  if (gereSignature()) {
    const presente = signatureDuMessage(corps);
    if (presente) {
      // Signature déjà là : on la garde telle quelle et on rafraîchit la mémoire
      try {
        await memoriserSignature(item, presente);
      } catch (err) {
        // Une copie plus ancienne reste utilisable : on ne signale l'échec que sans mémoire
        if (!lireMemoire()) message += ", signature non mémorisée (" + (err && err.message ? err.message : err) + ")";
      }
    } else {
      const memoire = lireMemoire();
      if (memoire) {
        signature = await joindreImages(item, memoire);
        message = "Charte et signature appliquées";
      } else {
        message += ". Signature inconnue : cliquez une fois sur Charte dans un nouveau message signé";
      }
    }
  }

  // Après ajout d'images en ligne, relire le corps avant de l'écrire (exigence Outlook)
  const corpsFinal = signature ? await lireCorps(item) : corps;
  await ecrireCorps(item, construire(corpsFinal, signature));
  return message;
}

// Signature gérée dans le nouvel Outlook et Outlook sur le web seulement : l'Outlook
// classique ne marque pas sa signature, le bouton risquerait de la doubler
function gereSignature() {
  const diag = Office.context.diagnostics || {};
  const web = diag.hostName === "OutlookWebApp" || diag.hostName === "newOutlookWindows"
    || Office.context.platform === Office.PlatformType.OfficeOnline;
  return web && Office.context.requirements.isSetSupported("Mailbox", "1.8");
}

function lireCorps(item) {
  return new Promise((ok, ko) =>
    item.body.getAsync(Office.CoercionType.Html, (res) =>
      res.status === Office.AsyncResultStatus.Succeeded ? ok(res.value) : ko(res.error)));
}

function ecrireCorps(item, html) {
  return new Promise((ok, ko) =>
    item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, (res) =>
      res.status === Office.AsyncResultStatus.Succeeded ? ok() : ko(res.error)));
}

function estProtege(n) {
  return n.nodeType === 1 && (n.matches(SEL_PROTEGE) || n.querySelector(SEL_PROTEGE));
}

// Sépare la partie rédigée du reste (signature, message cité)
function decouper(doc) {
  // Descend dans les enveloppes uniques (div global) qui contiennent signature ou citation
  let racine = doc.body;
  for (;;) {
    const elts = Array.from(racine.childNodes).filter((n) => n.nodeType === 1 || n.textContent.trim());
    if (elts.length !== 1 || elts[0].nodeType !== 1 || elts[0].matches(SEL_PROTEGE)
        || !elts[0].querySelector(SEL_PROTEGE)) break;
    racine = elts[0];
  }
  const enfants = Array.from(racine.childNodes);

  // Coupe au premier bloc protégé
  let coupe = enfants.findIndex(estProtege);
  if (coupe === -1) coupe = enfants.length;
  // Inclut le filet <hr> placé juste avant l'en-tête de citation
  let i = coupe - 1;
  while (i >= 0 && enfants[i].nodeType === 3 && !enfants[i].textContent.trim()) i--;
  if (i >= 0 && enfants[i].nodeName === "HR") coupe = i;

  return { racine, redige: enfants.slice(0, coupe), reste: enfants.slice(coupe) };
}

// Signature de l'expéditeur placée avant le message cité (celle du correspondant ne compte pas)
function signatureDans(reste) {
  for (const n of reste) {
    if (n.nodeType !== 1) continue;
    if (n.matches(SEL_SIGNATURE_PROPRE)) return n;
    if (n.matches(SEL_CITATION) || n.querySelector(SEL_CITATION)) return null;
    const sig = n.querySelector(SEL_SIGNATURE_PROPRE);
    if (sig) return sig;
  }
  return null;
}

function signatureDuMessage(corpsHtml) {
  const doc = new DOMParser().parseFromString(corpsHtml, "text/html");
  const sig = signatureDans(decouper(doc).reste);
  return sig && (sig.textContent.trim() || sig.querySelector("img")) ? sig.outerHTML : null;
}

function construire(corpsHtml, signatureHtml) {
  const doc = new DOMParser().parseFromString(corpsHtml, "text/html");
  const { racine, redige, reste } = decouper(doc);

  // Bloc déjà posé par un clic précédent : on le réutilise au lieu d'en imbriquer un autre
  let bloc = redige.length === 1 && redige[0].nodeType === 1 && redige[0].getAttribute("data-lf-charte") !== null
    ? redige[0] : null;
  if (!bloc) {
    bloc = doc.createElement("div");
    redige.forEach((n) => bloc.appendChild(n));
  }
  bloc.setAttribute("data-lf-charte", "");
  bloc.setAttribute("style", `font-family: ${POLICE}; font-size: ${TAILLE}; color: ${COULEUR};`);

  // Neutralise les polices imposées par Outlook (Aptos, Calibri...) dans la partie rédigée
  bloc.querySelectorAll("[style]").forEach((el) => {
    if (el === bloc) return;
    el.style.removeProperty("font-family");
    el.style.removeProperty("font-size");
    if (el.nodeName !== "A") el.style.removeProperty("color");
    if (!el.getAttribute("style").trim()) el.removeAttribute("style");
  });
  bloc.querySelectorAll("font").forEach((el) => {
    el.removeAttribute("face");
    el.removeAttribute("size");
    el.removeAttribute("color");
  });

  racine.innerHTML = "";
  racine.appendChild(bloc);
  if (signatureHtml && !signatureDans(reste)) {
    const tmp = doc.createElement("div");
    tmp.innerHTML = signatureHtml;
    racine.appendChild(doc.createElement("br"));
    Array.from(tmp.childNodes).forEach((n) => racine.appendChild(n));
  }
  reste.forEach((n) => racine.appendChild(n));

  return "<!DOCTYPE html>" + doc.documentElement.outerHTML;
}

// ---------- Mémoire de la signature ----------

function lireMemoire() {
  try {
    const brut = localStorage.getItem(CLE_SIGNATURE);
    return brut ? JSON.parse(brut) : null;
  } catch (e) {
    return null;
  }
}

function typeImage(nom) {
  const ext = (nom.split(".").pop() || "").toLowerCase();
  return { jpg: "jpeg", jpeg: "jpeg", gif: "gif", bmp: "bmp", webp: "webp" }[ext] || "png";
}

// Copie la signature du message (HTML et images en ligne) dans la mémoire du poste
async function memoriserSignature(item, signatureHtml) {
  const doc = new DOMParser().parseFromString(signatureHtml, "text/html");
  const images = [];
  let pieces = null;

  for (const img of Array.from(doc.querySelectorAll("img"))) {
    const src = img.getAttribute("src") || "";
    if (src.startsWith("data:")) {
      const m = /^data:image\/([a-z0-9.+-]+);base64,(.*)$/i.exec(src);
      if (!m) throw new Error("image de signature illisible");
      images.push({ type: m[1].toLowerCase(), b64: m[2] });
    } else if (src.toLowerCase().startsWith("cid:")) {
      if (!pieces) pieces = await piecesJointes(item);
      const piece = trouverPiece(pieces, src.slice(4));
      if (!piece) throw new Error("image " + src.slice(4) + " introuvable");
      const contenu = await contenuPiece(item, piece.id);
      images.push({ type: typeImage(piece.name), b64: contenu });
    } else {
      continue; // image publiée sur le web : l'adresse est gardée telle quelle
    }
    img.setAttribute("src", PREFIXE_IMAGE + (images.length - 1));
  }

  const sig = doc.body.firstElementChild;
  const memoire = { html: sig ? sig.outerHTML : doc.body.innerHTML, images, date: new Date().toISOString() };
  try {
    localStorage.setItem(CLE_SIGNATURE, JSON.stringify(memoire));
  } catch (e) {
    throw new Error("stockage du poste indisponible ou plein");
  }
}

function piecesJointes(item) {
  return new Promise((ok, ko) =>
    item.getAttachmentsAsync((res) =>
      res.status === Office.AsyncResultStatus.Succeeded ? ok(res.value) : ko(res.error)));
}

function trouverPiece(pieces, cid) {
  const nu = (s) => (s || "").replace(/^<|>$/g, "").toLowerCase();
  const c = nu(cid);
  const enLigne = pieces.filter((p) => p.isInline);
  return enLigne.find((p) => nu(p.contentId) === c)
    || enLigne.find((p) => nu(p.name) === c || nu(p.name) === c.split("@")[0]);
}

function contenuPiece(item, id) {
  return new Promise((ok, ko) =>
    item.getAttachmentContentAsync(id, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return ko(res.error);
      if (res.value.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
        return ko(new Error("format d'image inattendu"));
      }
      ok(res.value.content);
    }));
}

// Joint les images mémorisées en pièces jointes en ligne, renvoie le HTML qui les cite
async function joindreImages(item, memoire) {
  let html = memoire.html;
  const marque = Date.now().toString(36);
  for (let i = 0; i < memoire.images.length; i++) {
    const image = memoire.images[i];
    const nom = `lf-signature-${marque}-${i + 1}.${image.type === "jpeg" ? "jpg" : image.type}`;
    await new Promise((ok, ko) =>
      item.addFileAttachmentFromBase64Async(image.b64, nom, { isInline: true }, (res) =>
        res.status === Office.AsyncResultStatus.Succeeded ? ok() : ko(res.error)));
    html = html.split(`"${PREFIXE_IMAGE}${i}"`).join(`"cid:${nom}"`);
  }
  return html;
}

function notifier(item, message, erreur, event) {
  const texte = message.length > 150 ? message.slice(0, 147) + "..." : message;
  const contenu = erreur
    ? { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message: texte }
    : { type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: texte, icon: "Icone.16", persistent: false };
  item.notificationMessages.replaceAsync("lfCharte", contenu, () => event.completed());
}

// Enregistrement de la commande (ancien et nouveau runtime Outlook)
if (Office.actions) Office.actions.associate("habillerCharte", habillerCharte);
window.habillerCharte = habillerCharte;
