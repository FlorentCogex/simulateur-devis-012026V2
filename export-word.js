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

async function exporterWord(clone, nomFichier, cssWord) {
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
