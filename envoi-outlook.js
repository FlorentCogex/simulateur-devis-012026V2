// Envoi d'un document par Outlook (Microsoft 365) : génère le PDF du document affiché,
// puis crée un BROUILLON dans la boîte Outlook de la personne connectée, avec le PDF en
// pièce jointe. Rien n'est envoyé automatiquement : l'envoi se fait depuis Outlook.
//
// Prérequis : l'appli « Simulateur COGEX » déclarée dans Microsoft Entra (application
// monopage, URI de redirection .../auth.html, autorisation déléguée Mail.ReadWrite).
// Les bibliothèques (msal-browser pour la connexion, html2pdf pour le PDF) ne sont
// chargées qu'au premier clic sur « Envoyer par Outlook ».

const OUTLOOK_CONFIG = {
    clientId: 'efc3d5b8-a759-44e1-ad00-308cc2d6b5aa',   // ID d'application (client)
    tenantId: '0118976c-00bc-479c-a27c-b8afadb6e5f3'    // ID de l'annuaire (locataire)
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
// Fenêtre de saisie du destinataire, pré-remplie avec l'e-mail connu du dossier.
// Renvoie l'adresse saisie (éventuellement vide), ou null si l'utilisateur annule.
function demanderDestinataire(options) {
    return new Promise(resolve => {
        const fond = document.createElement('div');
        fond.className = 'no-print fenetre-outlook';
        fond.setAttribute('contenteditable', 'false');
        fond.style.cssText = 'position:fixed;inset:0;z-index:10001;background:rgba(0,0,0,.45);display:flex;align-items:center;justify-content:center;font-family:"Segoe UI",sans-serif;';
        fond.innerHTML = '<div style="background:#fff;border-radius:8px;padding:24px;width:min(440px,92vw);box-shadow:0 8px 30px rgba(0,0,0,.3);">'
            + '<div style="font-size:17px;font-weight:bold;color:#1a4d2e;margin-bottom:6px;">✉️ Envoyer par Outlook</div>'
            + '<div style="font-size:13px;color:#666;margin-bottom:16px;">Objet : ' + echapperHtml(options.objet) + '</div>'
            + '<label style="display:block;font-size:13px;font-weight:bold;color:#333;margin-bottom:6px;">E-mail du destinataire</label>'
            + '<input type="text" class="champ-destinataire" placeholder="exemple@domaine.fr" style="width:100%;box-sizing:border-box;padding:10px;border:1px solid #bbb;border-radius:4px;font-size:15px;">'
            + '<div style="font-size:12px;color:#888;margin-top:6px;">Plusieurs adresses : séparez-les par « ; ». Vous pouvez aussi laisser vide et le choisir dans Outlook.</div>'
            + '<div class="erreur-destinataire" style="font-size:13px;color:#c0392b;margin-top:8px;min-height:1em;"></div>'
            + '<div style="display:flex;gap:10px;justify-content:flex-end;margin-top:12px;">'
            + '<button type="button" class="annuler" style="padding:10px 18px;border:1px solid #bbb;background:#fff;border-radius:4px;cursor:pointer;font-size:14px;">Annuler</button>'
            + '<button type="button" class="valider" style="padding:10px 18px;border:none;background:#0f6cbd;color:#fff;border-radius:4px;cursor:pointer;font-size:14px;font-weight:bold;">Créer le brouillon</button>'
            + '</div></div>';
        document.body.appendChild(fond);

        const champ = fond.querySelector('.champ-destinataire');
        champ.value = options.destinataire || '';
        champ.focus();
        champ.select();

        const fermer = valeur => { fond.remove(); resolve(valeur); };
        const valider = () => {
            const adresses = champ.value.split(/[;,]/).map(a => a.trim()).filter(Boolean);
            const invalide = adresses.find(a => !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(a));
            if (invalide) {
                fond.querySelector('.erreur-destinataire').textContent = 'Adresse e-mail invalide : ' + invalide;
                champ.focus();
                return;
            }
            fermer(adresses.join('; '));
        };
        fond.querySelector('.valider').addEventListener('click', valider);
        fond.querySelector('.annuler').addEventListener('click', () => fermer(null));
        champ.addEventListener('keydown', e => {
            if (e.key === 'Enter') { e.preventDefault(); valider(); }
            if (e.key === 'Escape') fermer(null);
        });
    });
}

// options : { nomFichier, objet, message (texte, lignes séparées par \n), destinataire,
//             cleEmail (clé localStorage où mémoriser l'adresse saisie) }
async function envoyerParOutlook(options) {
    const destinataire = await demanderDestinataire(options);
    if (destinataire === null) return;
    options = { ...options, destinataire };
    if (options.cleEmail && destinataire) {
        try { localStorage.setItem(options.cleEmail, destinataire); } catch (e) { /* stockage indisponible */ }
    }

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
