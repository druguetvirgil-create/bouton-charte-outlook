/*
 * Bouton « Charte », Le Froid Pecomark, Service Informatique
 * Sur le message en cours de rédaction :
 *   1. applique Century Gothic 10 pt #00447A à la partie rédigée
 *   2. laisse intacts la signature et le message cité (réponse, transfert)
 * Décision du 12/10/2026, Flora Llopis, responsable communication.
 */

const POLICE = "'Century Gothic', Arial, sans-serif";
const TAILLE = "10pt";
const COULEUR = "#00447A";

// Insertion d'une signature HTML (signature.html sur l'hébergement).
// Désactivée : la charte interdit de recréer la signature d'entreprise.
const INSERER_SIGNATURE = false;
const MARQUEUR_SIGNATURE = "Responsable Informatique";

// Blocs à ne jamais retoucher : signature Outlook, puis début du message cité
const SEL_SIGNATURE = "#Signature, [id^='Signature'], #x_Signature, [id^='x_Signature']";
const SEL_CITATION = "#appendonsend, #divRplyFwdMsg, [id^='divRplyFwdMsg'], #mail-editor-reference-message-container";
const SEL_PROTEGE = SEL_SIGNATURE + ", " + SEL_CITATION;

Office.onReady(() => {});

function habillerCharte(event) {
  const item = Office.context.mailbox.item;

  Promise.all([INSERER_SIGNATURE ? chargerSignature() : Promise.resolve(""), lireCorps(item)])
    .then(([signature, corps]) => ecrireCorps(item, construire(corps, signature)))
    .then(() => notifier(item, "Mise en forme charte appliquée", false, event))
    .catch((err) => notifier(item, "Bouton Charte : " + (err && err.message ? err.message : err), true, event));
}

function chargerSignature() {
  return fetch("signature.html", { cache: "no-store" }).then((r) => {
    if (!r.ok) throw new Error("signature.html introuvable sur l'hébergement");
    return r.text();
  });
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

function construire(corpsHtml, signatureHtml) {
  const doc = new DOMParser().parseFromString(corpsHtml, "text/html");
  // Descend dans les enveloppes uniques (div global) qui contiennent signature ou citation
  let body = doc.body;
  for (;;) {
    const elts = Array.from(body.childNodes).filter((n) => n.nodeType === 1 || n.textContent.trim());
    if (elts.length !== 1 || elts[0].nodeType !== 1 || elts[0].matches(SEL_PROTEGE)
        || !elts[0].querySelector(SEL_PROTEGE)) break;
    body = elts[0];
  }
  const enfants = Array.from(body.childNodes);

  // Coupe au premier bloc protégé (signature ou message cité)
  let coupe = enfants.findIndex(estProtege);
  if (coupe === -1) coupe = enfants.length;
  // Inclut le filet <hr> placé juste avant l'en-tête de citation
  let i = coupe - 1;
  while (i >= 0 && enfants[i].nodeType === 3 && !enfants[i].textContent.trim()) i--;
  if (i >= 0 && enfants[i].nodeName === "HR") coupe = i;

  const redige = enfants.slice(0, coupe);
  const reste = enfants.slice(coupe);

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

  body.innerHTML = "";
  body.appendChild(bloc);

  const dejaSignee = doc.querySelector(SEL_SIGNATURE) !== null
    || reste.some((n) => n.textContent.includes(MARQUEUR_SIGNATURE))
    || bloc.textContent.includes(MARQUEUR_SIGNATURE);
  if (INSERER_SIGNATURE && signatureHtml && !dejaSignee) {
    const sig = doc.createElement("div");
    sig.innerHTML = "<br>" + signatureHtml;
    body.appendChild(sig);
  }
  reste.forEach((n) => body.appendChild(n));

  return "<!DOCTYPE html>" + doc.documentElement.outerHTML;
}

function notifier(item, message, erreur, event) {
  const contenu = erreur
    ? { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message }
    : { type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message, icon: "Icone.16", persistent: false };
  item.notificationMessages.replaceAsync("lfCharte", contenu, () => event.completed());
}

// Enregistrement de la commande (ancien et nouveau runtime Outlook)
if (Office.actions) Office.actions.associate("habillerCharte", habillerCharte);
window.habillerCharte = habillerCharte;
