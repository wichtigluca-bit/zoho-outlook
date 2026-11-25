// Ziel-URL der Zoho-CRM-Suche (wie in der Chrome-Extension)
const ZOHO_SUCHE = 'https://crm.zoho.eu/crm/org20093392829/search?searchword=';

// PW-Kundennummer im Format "SFB-12345" (auch "SFB 12345" oder "sfb-12345")
const MUSTER = /\bSFB[\s-]?(\d+)\b/i;

// Liefert die Nummer immer einheitlich als "SFB-12345"
function nummerIn(text) {
  if (!text) return null;
  const m = text.match(MUSTER);
  return m ? 'SFB-' + m[1] : null;
}

function bodyLesen(item) {
  return new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Text, (res) => {
      resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value || '' : '');
    });
  });
}

// ---- PDF-Anhänge (Fallback, nur wenn in der Mail nichts steht) ----

let pdfjsGeladen = null;
function pdfjsLaden() {
  if (!pdfjsGeladen) {
    pdfjsGeladen = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = 'assets/pdf.min.js';
      s.onload = () => {
        window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'assets/pdf.worker.min.js';
        resolve(window.pdfjsLib);
      };
      s.onerror = reject;
      document.head.appendChild(s);
    });
  }
  return pdfjsGeladen;
}

function anhangInhalt(item, id) {
  return new Promise((resolve) => {
    item.getAttachmentContentAsync(id, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return resolve(null);
      const v = res.value;
      resolve(v.format === Office.MailboxEnums.AttachmentContentFormat.Base64 ? v.content : null);
    });
  });
}

async function nummerAusPdf(base64) {
  const pdfjsLib = await pdfjsLaden();
  const bin = atob(base64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);

  const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
  const seiten = Math.min(pdf.numPages, 5); // Nummer steht praktisch immer vorne
  for (let p = 1; p <= seiten; p++) {
    const seite = await pdf.getPage(p);
    const inhalt = await seite.getTextContent();
    const text = inhalt.items.map((t) => t.str).join(' ');
    const nr = nummerIn(text);
    if (nr) return nr;
  }
  return null;
}

async function nummerAusAnhaengen(item) {
  if (!Office.context.requirements.isSetSupported('Mailbox', '1.8')) return null;
  const pdfs = (item.attachments || []).filter((a) =>
    !a.isInline && /\.pdf$/i.test(a.name || '')
  );
  for (const a of pdfs) {
    try {
      const b64 = await anhangInhalt(item, a.id);
      if (!b64) continue;
      const nr = await nummerAusPdf(b64);
      if (nr) return nr;
    } catch (e) {
      // beschädigte oder verschlüsselte PDF: nächste probieren
    }
  }
  return null;
}

// ---- Hinweis in der Mail-Infoleiste ----

function hinweis(item, text) {
  return new Promise((resolve) => {
    try {
      item.notificationMessages.replaceAsync('zohoSuche', {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: text.slice(0, 150),
        icon: 'Icon.16',
        persistent: false
      }, () => resolve());
    } catch (e) { resolve(); }
  });
}

function oeffnen(url) {
  try {
    Office.context.ui.openBrowserWindow(url);
  } catch (e) {
    window.open(url, '_blank');
  }
}

// ---- Button-Funktion ----

async function zohoSuchen(event) {
  const item = Office.context.mailbox.item;
  try {
    let nr = nummerIn(item.subject) || nummerIn(await bodyLesen(item));
    let quelle = 'Mail';

    if (!nr) {
      nr = await nummerAusAnhaengen(item);
      quelle = 'PDF-Anhang';
    }

    if (nr) {
      await hinweis(item, `Zoho-Suche: PW-Nummer ${nr} (aus ${quelle})`);
      oeffnen(ZOHO_SUCHE + encodeURIComponent(nr));
    } else {
      const absender = item.from && item.from.emailAddress;
      await hinweis(item, absender
        ? `Keine PW-Nummer gefunden – Suche nach Absender ${absender}`
        : 'Keine PW-Nummer gefunden');
      oeffnen(absender ? ZOHO_SUCHE + encodeURIComponent(absender) : ZOHO_SUCHE);
    }
  } catch (e) {
    oeffnen(ZOHO_SUCHE);
  } finally {
    event.completed();
  }
}

Office.onReady(() => {
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate('zohoSuchen', zohoSuchen);
  }
});
window.zohoSuchen = zohoSuchen;
