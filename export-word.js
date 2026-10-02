// Export Word commun aux lettres et courriers.
// Le fichier .doc est généré au format MHTML : les images (logo) sont intégrées dans le
// fichier lui-même, elles s'affichent donc dans Word même hors connexion.

function imageEnBase64(src, hauteurMax) {
    return new Promise(resolve => {
        const source = new Image();
        source.onload = () => {
            const ratio = Math.min(1, hauteurMax / source.naturalHeight);
            const canvas = document.createElement('canvas');
            canvas.width = Math.round(source.naturalWidth * ratio);
            canvas.height = Math.round(source.naturalHeight * ratio);
            const ctx = canvas.getContext('2d');
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, canvas.width, canvas.height);
            ctx.drawImage(source, 0, 0, canvas.width, canvas.height);
            resolve({
                base64: canvas.toDataURL('image/jpeg', 0.92).split(',')[1],
                largeur: canvas.width,
                hauteur: canvas.height
            });
        };
        source.onerror = () => resolve(null);
        source.src = src;
    });
}

function utf8EnBase64(texte) {
    const octets = new TextEncoder().encode(texte);
    let binaire = '';
    for (let i = 0; i < octets.length; i += 0x8000) {
        binaire += String.fromCharCode.apply(null, octets.subarray(i, i + 0x8000));
    }
    return btoa(binaire);
}

function decouperBase64(base64) {
    return base64.match(/.{1,76}/g).join('\r\n');
}

// Retire .page-header / .page-footer du corps et renvoie leur version Word : une page HTML
// séparée contenant les blocs mso-element:header / footer (comme Word l'enregistre lui-même),
// ou null s'il n'y en a pas
function construireEntetePied(clone) {
    const entete = clone.querySelector('.page-header');
    const pied = clone.querySelector('.page-footer');
    if (!entete && !pied) return null;
    const vert = '#1A4D2E';
    const police = "font-family:'Bahnschrift Light','Segoe UI',sans-serif;";
    const filet = 'border:none;border-bottom:solid ' + vert + ' 1.0pt;mso-border-bottom-alt:solid ' + vert + ' .75pt;padding:0 0 5pt 0;';

    let htmlEntete = '';
    if (entete) {
        const nom = (entete.querySelector('.header-company-name') || {}).textContent || '';
        const profession = (entete.querySelector('.header-profession') || {}).textContent || '';
        const img = entete.querySelector('img');
        htmlEntete = '<div style="mso-element:header" id="h1">'
            + '<table width="100%" border="0" cellspacing="0" cellpadding="0" style="width:100%;border-collapse:collapse;border:none;mso-border-alt:none;mso-padding-alt:0 0 0 0;">'
            + '<tr><td width="65%" valign="bottom" style="width:65%;' + filet + '">'
            + '<p class="MsoHeader" style="margin:0;line-height:normal;"><b><span style="' + police + 'font-size:13pt;color:' + vert + ';">' + nom.trim() + '</span></b></p>'
            + '<p class="MsoHeader" style="margin:2pt 0 0 0;line-height:normal;"><span style="' + police + 'font-size:10.5pt;color:' + vert + ';">' + profession.trim() + '</span></p>'
            + '</td><td width="35%" valign="bottom" align="right" style="width:35%;' + filet + '">'
            + '<p class="MsoHeader" align="right" style="margin:0;text-align:right;line-height:normal;">' + (img ? img.outerHTML : '') + '</p>'
            + '</td></tr></table>'
            + '<p class="MsoHeader" style="margin:0;line-height:4pt;mso-line-height-rule:exactly;"><span style="font-size:4pt;"></span></p>'
            + '</div>';
        entete.remove();
    }

    let htmlPied = '';
    if (pied) {
        const lignes = Array.from(pied.querySelectorAll('.footer-line')).map(l => l.textContent.trim());
        htmlPied = '<div style="mso-element:footer" id="f1">'
            + '<p class="MsoFooter" align="center" style="margin:0;text-align:center;line-height:normal;border:none;border-top:solid ' + vert + ' 1.0pt;mso-border-top-alt:solid ' + vert + ' .75pt;padding:4pt 0 0 0;">'
            + '<span style="' + police + 'font-size:8.5pt;color:#333333;">' + lignes.join('<br>') + '</span></p>'
            + '</div>';
        pied.remove();
    }

    return '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">'
        + '<head><meta charset="utf-8"><link id="Main-File" rel="Main-File" href="file:///C:/document/document.htm"></head>'
        + '<body>' + htmlEntete + htmlPied + '</body></html>';
}

// options.enteteWord : en-tête et pied de page placés dans ceux de Word (utilisé par le courrier confrère)
async function exporterWord(clone, nomFichier, cssWord, options = {}) {
    const images = [];
    const balisesImg = Array.from(clone.querySelectorAll('img'));
    for (let i = 0; i < balisesImg.length; i++) {
        const img = balisesImg[i];
        const src = new URL(img.getAttribute('src'), document.baseURI).href;
        const hauteur = parseInt(img.getAttribute('height'), 10) || 48;
        // Résolution x4 pour un logo net à l'impression
        const donnees = await imageEnBase64(src, hauteur * 4);
        if (!donnees) { img.setAttribute('src', src); continue; }
        const emplacement = 'file:///C:/document/image' + i + '.jpg';
        img.setAttribute('src', emplacement);
        img.setAttribute('height', hauteur);
        img.setAttribute('width', Math.round(hauteur * donnees.largeur / donnees.hauteur));
        images.push({ emplacement, base64: donnees.base64 });
    }

    // En-tête et pied de page placés dans les vrais en-tête / pied de page Word
    // (répétés sur chaque page, pied collé en bas de page)
    const entetePied = options.enteteWord ? construireEntetePied(clone) : null;
    const emplacementEntetePied = 'file:///C:/document/header.htm';
    if (entetePied) {
        cssWord = cssWord.replace('@page WordSection1 {',
            '@page WordSection1 { mso-header-margin: 1cm; mso-footer-margin: 0.8cm;'
            + ' mso-header: url("' + emplacementEntetePied + '") h1; mso-footer: url("' + emplacementEntetePied + '") f1;');
        cssWord += ' p.MsoHeader, p.MsoFooter { margin: 0; }';
    }

    const html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">'
        + '<head><meta charset="utf-8"><title>' + nomFichier + '</title><style>' + cssWord + '</style></head>'
        + '<body><div class="WordSection1">' + clone.innerHTML + '</div></body></html>';

    const limite = '----=_Part_' + Date.now();
    let mhtml = 'MIME-Version: 1.0\r\n'
        + 'Content-Type: multipart/related; boundary="' + limite + '"; type="text/html"\r\n\r\n'
        + '--' + limite + '\r\n'
        + 'Content-Type: text/html; charset="utf-8"\r\n'
        + 'Content-Transfer-Encoding: base64\r\n'
        + 'Content-Location: file:///C:/document/document.htm\r\n\r\n'
        + decouperBase64(utf8EnBase64(html)) + '\r\n';
    if (entetePied) {
        mhtml += '--' + limite + '\r\n'
            + 'Content-Type: text/html; charset="utf-8"\r\n'
            + 'Content-Transfer-Encoding: base64\r\n'
            + 'Content-Location: ' + emplacementEntetePied + '\r\n\r\n'
            + decouperBase64(utf8EnBase64(entetePied)) + '\r\n';
    }
    images.forEach(image => {
        mhtml += '--' + limite + '\r\n'
            + 'Content-Type: image/jpeg\r\n'
            + 'Content-Transfer-Encoding: base64\r\n'
            + 'Content-Location: ' + image.emplacement + '\r\n\r\n'
            + decouperBase64(image.base64) + '\r\n';
    });
    mhtml += '--' + limite + '--\r\n';

    const blob = new Blob([mhtml], { type: 'application/msword' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = nomFichier + '.doc';
    link.click();
}
