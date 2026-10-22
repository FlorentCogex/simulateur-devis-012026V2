// Envoi d'un document par Outlook (Microsoft 365) : génère le PDF du document affiché,
// puis crée un BROUILLON dans la boîte Outlook de la personne connectée, avec le PDF en
// pièce jointe. Rien n'est envoyé automatiquement : l'envoi se fait depuis Outlook.
//
// Prérequis : l'appli « Simulateur COGEX » déclarée dans Microsoft Entra (application
// monopage, URI de redirection .../auth.html, autorisation déléguée Mail.ReadWrite).
// Les bibliothèques (msal-browser pour la connexion, html2pdf pour le PDF) ne sont
// chargées qu'au premier clic sur « Envoyer par Outlook ».

const OUTLOOK_CONFIG = {
    clientId: 'A_RENSEIGNER',   // ID d'application (client)
    tenantId: 'A_RENSEIGNER'    // ID de l'annuaire (locataire)
};
const OUTLOOK_SCOPES = ['Mail.ReadWrite'];

const BIBLIOTHEQUES_OUTLOOK = {
    msal: 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.27.0/lib/msal-browser.min.js',
    html2pdf: 'https://cdn.jsdelivr.net/npm/html2pdf.js@0.10.2/dist/html2pdf.bundle.min.js'
};

const scriptsCharges = {};

function chargerScript(url) {
    if (!scriptsCharges[url]) {
        scriptsCharges[url] = new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = url;
            script.onload = resolve;
            script.onerror = () => { delete scriptsCharges[url]; reject(new Error('Impossible de charger ' + url + ' (connexion Internet ?)')); };
            document.head.appendChild(script);
        });
    }
    return scriptsCharges[url];
}

let clientMsal = null;

async function obtenirClientMsal() {
    if (clientMsal) return clientMsal;
    if (OUTLOOK_CONFIG.clientId === 'A_RENSEIGNER') {
        throw new Error("L'envoi par Outlook n'est pas encore configuré (identifiants Microsoft 365 manquants).");
    }
    await chargerScript(BIBLIOTHEQUES_OUTLOOK.msal);
    clientMsal = new msal.PublicClientApplication({
        auth: {
            clientId: OUTLOOK_CONFIG.clientId,
            authority: 'https://login.microsoftonline.com/' + OUTLOOK_CONFIG.tenantId,
            redirectUri: new URL('auth.html', document.baseURI).href
        },
        cache: { cacheLocation: 'localStorage' }
    });
    await clientMsal.initialize();
    return clientMsal;
}

async function obtenirJetonGraph() {
    const client = await obtenirClientMsal();
    const compte = client.getAllAccounts()[0];
    if (compte) {
        try {
            return (await client.acquireTokenSilent({ scopes: OUTLOOK_SCOPES, account: compte })).accessToken;
        } catch (e) { /* jeton expiré : on repasse par la fenêtre de connexion */ }
    }
    return (await client.acquireTokenPopup({ scopes: OUTLOOK_SCOPES, prompt: 'select_account' })).accessToken;
}

// Applique à l'écran les règles @media print de la page (en-tête, pied de page,
// fond blanc, boutons masqués) le temps de générer le PDF
function activerMiseEnPagePdf() {
    const regles = [];
    for (const feuille of Array.from(document.styleSheets)) {
        let liste;
        try { liste = feuille.cssRules; } catch (e) { continue; }
        for (const regle of Array.from(liste)) {
            if (regle.media && /print/.test(regle.media.mediaText)) {
                for (const r of Array.from(regle.cssRules)) {
                    if (!(r.cssText || '').startsWith('@page')) regles.push(r.cssText);
                }
            }
        }
    }
    const style = document.createElement('style');
    style.textContent = regles.join('\n')
        + '\n.no-print, .bandeau-outlook { display: none !important; }'
        + '\nbody { background: #fff !important; padding: 0 !important; }';
    document.head.appendChild(style);
    return () => style.remove();
}

async function genererPdfBlob() {
    await chargerScript(BIBLIOTHEQUES_OUTLOOK.html2pdf);
    const desactiver = activerMiseEnPagePdf();
    // Copie du contenu imprimable dans un bloc à la largeur utile d'une page A4
    const page = document.createElement('div');
    page.style.cssText = 'width:186mm;background:#fff;';
    Array.from(document.body.children).forEach(el => {
        if (el.matches('script, .no-print, .bandeau-outlook')) return;
        page.appendChild(el.cloneNode(true));
    });
    // Les règles d'impression « ne pas couper ce bloc » s'appliquent à des sections
    // entières : le PDF n'évite les coupures que ligne par ligne (paragraphes, lignes de tableau)
    const sansBlocsInsecables = document.createElement('style');
    sansBlocsInsecables.textContent = '* { page-break-inside: auto !important; break-inside: auto !important; }';
    document.head.appendChild(sansBlocsInsecables);
    try {
        return await html2pdf().set({
            margin: [12, 12, 14, 12],
            image: { type: 'jpeg', quality: 0.75 },
            html2canvas: { scale: 1.4, useCORS: true, backgroundColor: '#ffffff' },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait', compress: true },
            pagebreak: { mode: ['legacy'], avoid: ['p', 'li', 'tr', 'h1', 'h2', 'h3', '.tariff-row', '.description-item', '.signature-section', '.signature-block', '.condition-title', '.mission-title', '.page-footer', '.page-header'] }
        }).from(page).outputPdf('blob');
    } finally {
        sansBlocsInsecables.remove();
        desactiver();
    }
}

function blobEnBase64(blob) {
    return new Promise((resolve, reject) => {
        const lecteur = new FileReader();
        lecteur.onload = () => resolve(lecteur.result.split(',')[1]);
        lecteur.onerror = reject;
        lecteur.readAsDataURL(blob);
    });
}

function echapperHtml(texte) {
    return String(texte || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

// Bandeau d'état en haut de la page (progression, lien vers le brouillon, erreurs)
function afficherBandeauOutlook(html, couleur) {
    let bandeau = document.querySelector('.bandeau-outlook');
    if (!bandeau) {
        bandeau = document.createElement('div');
        bandeau.className = 'bandeau-outlook';
        bandeau.setAttribute('contenteditable', 'false');
        bandeau.style.cssText = 'position:fixed;top:70px;right:20px;z-index:10000;max-width:380px;padding:12px 16px;'
            + 'border-radius:6px;font:14px "Segoe UI",sans-serif;box-shadow:0 2px 8px rgba(0,0,0,.25);background:#fff;';
        document.body.appendChild(bandeau);
    }
    bandeau.style.borderLeft = '5px solid ' + (couleur || '#2b579a');
    bandeau.innerHTML = html + ' <a href="#" onclick="this.parentElement.remove();return false;" style="margin-left:8px;color:#888;text-decoration:none;">✕</a>';
}

// options : { nomFichier, objet, message (texte, lignes séparées par \n), destinataire }
async function envoyerParOutlook(options) {
    try {
        afficherBandeauOutlook('Connexion à Outlook…');
        const jeton = await obtenirJetonGraph();

        afficherBandeauOutlook('Préparation du PDF…');
        const pdf = await genererPdfBlob();

        afficherBandeauOutlook('Création du brouillon dans Outlook…');
        const destinataires = String(options.destinataire || '').split(/[;,]/).map(a => a.trim()).filter(Boolean)
            .map(adresse => ({ emailAddress: { address: adresse } }));
        const reponse = await fetch('https://graph.microsoft.com/v1.0/me/messages', {
            method: 'POST',
            headers: { 'Authorization': 'Bearer ' + jeton, 'Content-Type': 'application/json' },
            body: JSON.stringify({
                subject: options.objet,
                body: { contentType: 'HTML', content: echapperHtml(options.message).replace(/\n/g, '<br>') },
                toRecipients: destinataires,
                attachments: [{
                    '@odata.type': '#microsoft.graph.fileAttachment',
                    name: options.nomFichier + '.pdf',
                    contentType: 'application/pdf',
                    contentBytes: await blobEnBase64(pdf)
                }]
            })
        });
        if (!reponse.ok) throw new Error('Outlook a refusé la création du brouillon (' + reponse.status + ') : ' + (await reponse.text()).slice(0, 200));
        const brouillon = await reponse.json();

        const fenetre = window.open(brouillon.webLink, '_blank');
        afficherBandeauOutlook('✅ Brouillon créé dans Outlook (dossier <b>Brouillons</b>). '
            + (fenetre ? '' : '<br><a href="' + echapperHtml(brouillon.webLink) + '" target="_blank" style="color:#2b579a;font-weight:bold;">Ouvrir le brouillon</a>'),
            '#1a4d2e');
    } catch (e) {
        console.error(e);
        afficherBandeauOutlook('❌ ' + echapperHtml(e.message || e), '#c0392b');
    }
}
