// ==UserScript==
// @name         IPECA - Export Remboursements Excel
// @description  Ajoute un bouton pour exporter les remboursements santé IPECA vers Excel
// @match        https://www.ipeca.fr/mes-remboursements*
// @match        https://www.ipeca.fr/ng/client/mes-remboursements*
// @namespace    https://github.com/laurentlemercier/tampermonkey-scripts
// @grant        GM_registerMenuCommand
// @grant        GM_unregisterMenuCommand
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        unsafeWindow
// @author       Laurent Lemercier et Claude AI
// @version      3.3
// @updateURL    https://raw.githubusercontent.com/laurentlemercier/tampermonkey-scripts/main/ipeca.user.js
// @downloadURL  https://raw.githubusercontent.com/laurentlemercier/tampermonkey-scripts/main/ipeca.user.js
// ==/UserScript==

(function () {
    'use strict';

    const BTN_STYLE = 'padding:6px 14px;background:#1F4E79;color:white;border:none;border-radius:4px;cursor:pointer;font-size:13px;font-family:Arial,sans-serif;vertical-align:middle;';
    // ── Mode debug : activable/désactivable via le menu Tampermonkey ─────────
    // (icône Tampermonkey → nom du script → "Debug"). Le choix est mémorisé.
    const BOX_STYLE = 'display:inline-flex;align-items:center;gap:10px;vertical-align:middle;';
    let DEBUG = false;
    try { if (typeof GM_getValue === 'function') DEBUG = !!GM_getValue('debug', false); } catch (e) { /* ignore */ }

    const W = (typeof unsafeWindow !== 'undefined') ? unsafeWindow : window;
    let debugMenuId = null;

    function showToast(msg) {
        if (!document.body) return;
        const t = document.createElement('div');
        t.textContent = msg;
        t.style.cssText = 'position:fixed;bottom:20px;right:20px;z-index:100000;background:#1F4E79;color:#fff;' +
            'padding:10px 16px;border-radius:6px;font:13px Arial,sans-serif;box-shadow:0 2px 8px rgba(0,0,0,.3);';
        document.body.appendChild(t);
        setTimeout(() => t.remove(), 3000);
    }

    function registerDebugMenu() {
        if (typeof GM_registerMenuCommand !== 'function') return;
        const canRefresh = typeof GM_unregisterMenuCommand === 'function';
        if (debugMenuId !== null && canRefresh) {
            try { GM_unregisterMenuCommand(debugMenuId); } catch (e) { /* ignore */ }
        }
        const label = canRefresh
            ? (DEBUG ? '🐞 Debug : ACTIVÉ (cliquer pour désactiver)' : '🐞 Debug : désactivé (cliquer pour activer)')
            : '🐞 Activer / désactiver le debug';
        debugMenuId = GM_registerMenuCommand(label, toggleDebug);
    }

    function toggleDebug() {
        DEBUG = !DEBUG;
        try { if (typeof GM_setValue === 'function') GM_setValue('debug', DEBUG); } catch (e) { /* ignore */ }
        console.log('[IPECA Export] Debug', DEBUG ? 'ACTIVÉ' : 'désactivé');
        showToast('IPECA Export — debug ' + (DEBUG ? 'activé (logs dans la console F12)' : 'désactivé'));
        if (typeof GM_unregisterMenuCommand === 'function') registerDebugMenu();
    }

    registerDebugMenu();
    if (DEBUG) console.log('[IPECA Export] Script chargé, debug ACTIVÉ');


    // ── Injection du bouton ──────────────────────────────────────────────────
    function findHost() {
        return document.querySelector('app-reimbursement-data-table .pagination')            // nouveau site
            || document.querySelector('.wrapper-refund-sante .pagination .content-pagination'); // ancien site
    }

    function injectButton(allowFallback) {
        const host = findHost();
        const existing = document.getElementById('ipeca-export-box');

        if (existing) {
            // Si le bloc était flottant et que la barre de pagination est apparue : on le déplace
            if (host && existing.dataset.floating === '1') {
                existing.dataset.floating = '0';
                existing.style.cssText = BOX_STYLE + 'margin-left:16px;';
                host.appendChild(existing);
            }
            return;
        }
        if (!host && !allowFallback) return;

        const box = document.createElement('span');
        box.id = 'ipeca-export-box';

        const btn = document.createElement('button');
        btn.id = 'ipeca-export-btn';
        btn.type = 'button';
        btn.textContent = '⬇ Exporter vers Excel';
        btn.style.cssText = BTN_STYLE;
        btn.addEventListener('mouseenter', () => btn.style.background = '#2E75B6');
        btn.addEventListener('mouseleave', () => btn.style.background = '#1F4E79');
        btn.addEventListener('click', exportToExcel);

        box.appendChild(btn);

        if (host) {
            box.style.cssText = BOX_STYLE + 'margin-left:16px;';
            box.dataset.floating = '0';
            host.appendChild(box);
        } else {
            box.style.cssText = BOX_STYLE + 'position:fixed;top:120px;right:20px;z-index:99999;background:#fff;padding:6px 8px;border-radius:6px;box-shadow:0 2px 8px rgba(0,0,0,.25);';
            box.dataset.floating = '1';
            document.body.appendChild(box);
            if (DEBUG) console.warn('[IPECA Export] Pagination introuvable : bloc flottant utilisé');
        }
        if (DEBUG) console.log('[IPECA Export] Bouton injecté');
    }

    // ── Extraction (ancien site) ───────────────────────────────────────────────
    function extractDataLegacy() {
        const rows = [];
        const mainTable = document.querySelector('.wrapper-refund-sante table.ipeca-table');
        if (!mainTable) return rows;

        const parseAmount = (str) =>
            parseFloat((str || '').replace(/\s/g, '').replace('€', '').replace(',', '.')) || 0;

        const topLevelRows = Array.from(
            mainTable.children[0]?.children ||
            mainTable.querySelectorAll(':scope > tbody > tr, :scope > tr')
        );

        let currentMonth = '';

        topLevelRows.forEach(tr => {
            const monthCell = tr.querySelector(':scope > td.ipeca-table-title-row');
            if (monthCell) { currentMonth = monthCell.textContent.trim(); return; }
            if (!tr.classList.contains('align-td')) return;

            const summaryRow = tr.querySelector('table.ipeca-content tr');
            if (!summaryRow) return;
            const sCells = summaryRow.querySelectorAll('td');
            if (sCells.length < 5) return;

            const getText = (cell) => {
                const clone = cell.cloneNode(true);
                clone.querySelectorAll('span.mobile').forEach(s => s.remove());
                return clone.textContent.trim();
            };

            const dateRemb = getText(sCells[0]);
            const destinat = getText(sCells[1]);
            // Convertit "27 MAR." + mois courant "MARS 2026" → "2026-03-27"
            // Le mois est lu depuis le titre de section (plus fiable que l'abréviation
            // du jour, car juin et juillet partagent tous les deux "JUI.")
            const MOIS = {
                'JANVIER': '01', 'FÉVRIER': '02', 'MARS': '03', 'AVRIL': '04',
                'MAI': '05', 'JUIN': '06', 'JUILLET': '07', 'AOÛT': '08',
                'SEPTEMBRE': '09', 'OCTOBRE': '10', 'NOVEMBRE': '11', 'DÉCEMBRE': '12',
            };
            const year  = (currentMonth.match(/\d{4}/) || [''])[0];
            const mWord = (currentMonth.match(/^([A-ZÉÛÎÀÈ]+)/) || [])[1] || '';
            const month = MOIS[mWord] || '00';
            const day   = (dateRemb.match(/^(\d{1,2})/) || [])[1]?.padStart(2, '0') || '00';
            const dateRembFull = year ? `${year}-${month}-${day}` : dateRemb;

            const subDetails = tr.querySelectorAll('td.sub-details');
            if (subDetails.length === 0) {
                rows.push({
                    mois: currentMonth, dateRemb: dateRembFull, destinataire: destinat,
                    beneficiaire: '', dateSoin: '', nature: '',
                    fraisReels: parseAmount(getText(sCells[2])),
                    rembSS: 0, partOblig: 0,
                    ipeca: parseAmount(getText(sCells[3])),
                    reste: parseAmount(getText(sCells[4])),
                    regularisation: false,
                });
                return;
            }

            subDetails.forEach(sub => {
                const header = sub.querySelector('.header-sub-details')?.textContent.trim() || '';
                const dateSoin = (header.match(/Soin du (\d{2}\/\d{2}\/\d{4})/) || [])[1] || '';
                const benef    = (header.match(/pour\s+([A-ZÉÀÈ]+)/i) || [])[1] || '';

                sub.querySelectorAll('table > tbody > tr, table > tr').forEach(detailTr => {
                    const dCells = detailTr.querySelectorAll(':scope > td');
                    if (dCells.length !== 6) return;
                    if (dCells[0].hasAttribute('height')) return;

                    const fraisReels = parseAmount(dCells[1].textContent);
                    rows.push({
                        mois: currentMonth, dateRemb: dateRembFull,
                        destinataire: destinat, beneficiaire: benef,
                        dateSoin, nature: dCells[0].textContent.replace(/€/g, '').trim(),
                        fraisReels, rembSS: parseAmount(dCells[2].textContent),
                        partOblig: parseAmount(dCells[3].textContent),
                        ipeca: parseAmount(dCells[4].textContent),
                        reste: parseAmount(dCells[5].textContent),
                        regularisation: fraisReels < 0,
                    });
                });
            });
        });
        return rows;
    }

    // ── Extraction (nouveau site Angular : /ng/client/mes-remboursements) ────
    function extractDataNg() {
        const rows = [];
        const table = document.querySelector('app-reimbursement-data-table table');
        if (!table) return rows;

        const clean = (s) => (s || '').replace(/\s+/g, ' ').trim();
        const strip = (s) => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase();
        const parseAmount = (str) =>
            parseFloat((str || '').replace(/\s/g, '').replace('€', '').replace(',', '.')) || 0;

        const MOIS = {
            'JANVIER': '01', 'FEVRIER': '02', 'MARS': '03', 'AVRIL': '04',
            'MAI': '05', 'JUIN': '06', 'JUILLET': '07', 'AOUT': '08',
            'SEPTEMBRE': '09', 'OCTOBRE': '10', 'NOVEMBRE': '11', 'DECEMBRE': '12',
        };
        const cellText = (tr, col) => {
            const td = tr.querySelector('td.cdk-column-' + col);
            return td ? clean(td.textContent) : '';
        };

        let currentMonth = '', year = '', monthNum = '00';
        let pending = null; // ligne résumé en attente de son bloc de détail

        const flushPending = () => {
            if (!pending) return;
            // Pas de détail disponible : ligne agrégée de secours
            rows.push({
                mois: currentMonth, dateRemb: pending.dateRemb, destinataire: pending.destinataire,
                beneficiaire: '', dateSoin: '', nature: '',
                fraisReels: pending.frais, rembSS: 0, partOblig: 0,
                ipeca: pending.ipeca, reste: pending.reste, regularisation: pending.frais < 0,
            });
            pending = null;
        };

        Array.from(table.querySelectorAll('tbody > tr')).forEach(tr => {
            // ── Titre de mois ("NOVEMBRE 2026") ────────────────────────────
            if (tr.classList.contains('group-row')) {
                flushPending();
                currentMonth = clean(tr.textContent);
                year = (currentMonth.match(/\d{4}/) || [''])[0];
                const word = (strip(currentMonth).match(/^[A-Z]+/) || [''])[0];
                monthNum = MOIS[word] || '00';
                return;
            }

            // ── Ligne de remboursement ─────────────────────────────────────
            if (tr.classList.contains('data-row')) {
                flushPending();
                const day = (cellText(tr, 'truncatDateDeRemboursement').match(/^(\d{1,2})/) || [])[1] || '00';
                pending = {
                    dateRemb: year ? `${year}-${monthNum}-${day.padStart(2, '0')}` : cellText(tr, 'truncatDateDeRemboursement'),
                    destinataire: cellText(tr, 'libelle'),
                    frais: parseAmount(cellText(tr, 'depense')),
                    ipeca: parseAmount(cellText(tr, 'montantVerseIpeca')),
                    reste: parseAmount(cellText(tr, 'rac')),
                };
                return;
            }

            // ── Ligne de détail (toujours dans le DOM, même repliée) ───────
            if (tr.classList.contains('example-detail-row')) {
                if (!pending) return; // détail vide sous un titre de mois
                const host = tr.querySelector('lib-my-refunds-table-details');
                let benef = '', dateSoin = '', count = 0;

                if (host) {
                    Array.from(host.children).forEach(el => {
                        if (el.classList.contains('title-detail-content')) {
                            const t = clean(el.textContent); // "Soin du 03/11/2026 pour LEIRE"
                            dateSoin = (t.match(/Soin du (\d{2}\/\d{2}\/\d{4})/) || [])[1] || '';
                            benef    = (t.match(/pour\s+(.+)$/) || [])[1] || '';
                            return;
                        }
                        el.querySelectorAll('.detail-content').forEach(dc => {
                            const sp = Array.from(dc.children);
                            if (sp.length < 6) return;
                            const frais = parseAmount(sp[1].textContent);
                            rows.push({
                                mois: currentMonth, dateRemb: pending.dateRemb,
                                destinataire: pending.destinataire,
                                beneficiaire: benef, dateSoin,
                                nature: clean(sp[0].textContent.replace(/€/g, '')),
                                fraisReels: frais,
                                rembSS: parseAmount(sp[2].textContent),
                                partOblig: parseAmount(sp[3].textContent),
                                ipeca: parseAmount(sp[4].textContent),
                                reste: parseAmount(sp[5].textContent),
                                regularisation: frais < 0,
                            });
                            count++;
                        });
                    });
                }
                if (count > 0) pending = null; else flushPending();
            }
        });
        flushPending();
        return rows;
    }

    function extractData() {
        return document.querySelector('app-reimbursement-data-table table')
            ? extractDataNg()
            : extractDataLegacy();
    }

    // ── Générateur XLSX natif (pas de dépendance CDN) ───────────────────────

    // Styles: index → (font, fill, border, numFmt, alignment)
    // 0 = texte normal       fillId=0 (none)
    // 1 = nombre normal      fillId=0 (none)
    // 2 = en-tête            fillId=3 (bleu foncé), font blanche
    // 3 = texte orange       fillId=2 (orange)
    // 4 = nombre orange      fillId=2 (orange)
    // 5 = total texte        fillId=4 (bleu clair), gras
    // 6 = total nombre       fillId=4 (bleu clair), gras
    const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1">
  <numFmt numFmtId="164" formatCode="#,##0.00 &quot;€&quot;"/>
</numFmts>
<fonts count="3">
  <font><sz val="10"/><name val="Arial"/></font>
  <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>
  <font><b/><sz val="10"/><name val="Arial"/></font>
</fonts>
<fills count="5">
  <fill><patternFill patternType="none"/></fill>
  <fill><patternFill patternType="gray125"/></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFFD7B0"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFBDD7EE"/></patternFill></fill>
</fills>
<borders count="2">
  <border><left/><right/><top/><bottom/><diagonal/></border>
  <border>
    <left style="thin"><color rgb="FFCCCCCC"/></left>
    <right style="thin"><color rgb="FFCCCCCC"/></right>
    <top style="thin"><color rgb="FFCCCCCC"/></top>
    <bottom style="thin"><color rgb="FFCCCCCC"/></bottom>
  </border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="7">
  <xf numFmtId="0"   fontId="0" fillId="0" borderId="1" xfId="0"><alignment horizontal="left"   vertical="center"/></xf>
  <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0"><alignment horizontal="right"  vertical="center"/></xf>
  <xf numFmtId="0"   fontId="1" fillId="3" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0"   fontId="0" fillId="2" borderId="1" xfId="0"><alignment horizontal="left"   vertical="center"/></xf>
  <xf numFmtId="164" fontId="0" fillId="2" borderId="1" xfId="0"><alignment horizontal="right"  vertical="center"/></xf>
  <xf numFmtId="0"   fontId="2" fillId="4" borderId="1" xfId="0"><alignment horizontal="left"   vertical="center"/></xf>
  <xf numFmtId="164" fontId="2" fillId="4" borderId="1" xfId="0"><alignment horizontal="right"  vertical="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

    function escXml(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function colLetter(n) { // 1-based → A, B, …, K
        let s = '';
        while (n > 0) { s = String.fromCharCode(65 + (n - 1) % 26) + s; n = Math.floor((n - 1) / 26); }
        return s;
    }

    function cellRef(row, col) { return colLetter(col) + row; }

    // Cellule texte inline
    function cStr(r, c, val, style) {
        return `<c r="${cellRef(r,c)}" s="${style}" t="inlineStr"><is><t>${escXml(val)}</t></is></c>`;
    }
    // Cellule nombre
    function cNum(r, c, val, style) {
        return `<c r="${cellRef(r,c)}" s="${style}"><v>${val}</v></c>`;
    }
    // Cellule formule
    function cForm(r, c, formula, style) {
        return `<c r="${cellRef(r,c)}" s="${style}"><f>${escXml(formula)}</f></c>`;
    }

    function buildSheetXml(data) {
        const COL_WIDTHS = [14, 22, 22, 14, 14, 28, 16, 16, 16, 16, 16];
        const HEADERS = [
            'Mois', 'Date de remboursement', 'Destinataire', 'Bénéficiaire',
            'Date du soin', 'Nature du soin',
            'Frais réels (€)', 'Remboursé SS (€)', 'Part obligatoire (€)',
            'Versé IPECA (€)', 'Reste à charge (€)',
        ];

        let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2"/></sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<cols>`;
        COL_WIDTHS.forEach((w, i) => {
            xml += `<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`;
        });
        xml += `</cols><sheetData>`;

        // Ligne 1 : en-têtes (style 2)
        xml += `<row r="1" ht="35" customHeight="1">`;
        HEADERS.forEach((h, i) => { xml += cStr(1, i+1, h, 2); });
        xml += `</row>`;

        // Regroupe les données par date de remboursement (en conservant l'ordre)
        const groups = [];
        data.forEach(row => {
            const last = groups[groups.length - 1];
            if (last && last.date === row.dateRemb) {
                last.rows.push(row);
            } else {
                groups.push({ date: row.dateRemb, rows: [row] });
            }
        });

        // Construit les lignes avec sous-totaux intercalés
        let r = 2; // numéro de ligne Excel courant
        const allDataRows = []; // référence aux lignes de données pour le TOTAL final

        groups.forEach(group => {
            const groupStart = r;

            group.rows.forEach(row => {
                const isReg = row.regularisation;
                const sText = isReg ? 3 : 0;
                const sNum  = isReg ? 4 : 1;
                xml += `<row r="${r}">`;
                xml += cStr(r, 1,  row.mois,         sText);
                xml += cStr(r, 2,  row.dateRemb,      sText);
                xml += cStr(r, 3,  row.destinataire,  sText);
                xml += cStr(r, 4,  row.beneficiaire,  sText);
                xml += cStr(r, 5,  row.dateSoin,      sText);
                xml += cStr(r, 6,  row.nature,        sText);
                xml += cNum(r, 7,  row.fraisReels,    sNum);
                xml += cNum(r, 8,  row.rembSS,        sNum);
                xml += cNum(r, 9,  row.partOblig,     sNum);
                xml += cNum(r, 10, row.ipeca,         sNum);
                xml += cNum(r, 11, row.reste,         sNum);
                xml += `</row>`;
                allDataRows.push(r);
                r++;
            });

            // Ligne de sous-total pour ce groupe de dates
            const groupEnd = r - 1;
            xml += `<row r="${r}">`;
            xml += cStr(r, 1, '', 5);
            xml += cStr(r, 2, `Sous-total ${group.date}`, 5);
            for (let c = 3; c <= 6; c++) xml += cStr(r, c, '', 5);
            for (let c = 7; c <= 11; c++) {
                xml += cForm(r, c, `SUM(${colLetter(c)}${groupStart}:${colLetter(c)}${groupEnd})`, 6);
            }
            xml += `</row>`;
            r++;
        });

        // Ligne de total général
        // Astuce : SUM(col) / 2 car chaque ligne de données est aussi comptée
        // dans son sous-total — diviser par 2 annule ce double-compte.
        const totalRow = r;
        const lastData = r - 1;
        xml += `<row r="${totalRow}">`;
        xml += cStr(totalRow, 1, 'TOTAL GÉNÉRAL', 5);
        for (let c = 2; c <= 6; c++) xml += cStr(totalRow, c, '', 5);
        for (let c = 7; c <= 11; c++) {
            xml += cForm(totalRow, c, `SUM(${colLetter(c)}2:${colLetter(c)}${lastData})/2`, 6);
        }
        xml += `</row>`;

        xml += `</sheetData></worksheet>`;
        if (DEBUG) {
            try {
                const parser = new DOMParser();
                const doc = parser.parseFromString(xml, 'application/xml');
                const err = doc.querySelector('parsererror');
                if (err) {
                    console.error('[IPECA Export] XML invalide :', err.textContent);
                    const lineMatch = err.textContent.match(/line (\d+)/i);
                    if (lineMatch) {
                        const lines = xml.split('\n');
                        const ln = parseInt(lineMatch[1]) - 1;
                        console.error('[IPECA Export] Contexte XML ligne', ln, ':', lines.slice(Math.max(0,ln-2), ln+3).join('\n'));
                    }
                } else {
                    console.log('[IPECA Export] XML valide');
                    const formules = [...xml.matchAll(/<f>([^<]+)<\/f>/g)].map(m => m[1]);
                    console.log('[IPECA Export] Formules :', formules);
                }
            } catch(e) { console.error('[IPECA Export] Erreur parse:', e); }
        }
        return xml;
    }

    // ── Compression ZIP en JS pur (deflate via CompressionStream) ───────────
    async function deflate(str) {
        const enc = new TextEncoder().encode(str);
        const cs  = new CompressionStream('deflate-raw');
        const writer = cs.writable.getWriter();
        writer.write(enc);
        writer.close();
        const chunks = [];
        const reader = cs.readable.getReader();
        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            chunks.push(value);
        }
        const total = chunks.reduce((s, c) => s + c.length, 0);
        const out = new Uint8Array(total);
        let offset = 0;
        for (const c of chunks) { out.set(c, offset); offset += c.length; }
        return out;
    }

    function crc32(data) {
        let crc = 0xFFFFFFFF;
        const table = crc32.table || (crc32.table = (() => {
            const t = new Uint32Array(256);
            for (let i = 0; i < 256; i++) {
                let c = i;
                for (let j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
                t[i] = c;
            }
            return t;
        })());
        for (let i = 0; i < data.length; i++) crc = table[(crc ^ data[i]) & 0xFF] ^ (crc >>> 8);
        return (crc ^ 0xFFFFFFFF) >>> 0;
    }

    function u16le(n) { return [n & 0xFF, (n >> 8) & 0xFF]; }
    function u32le(n) { return [n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF, (n >> 24) & 0xFF]; }

    async function buildZip(files) {
        // files: [{name, data (string)}]
        const enc = new TextEncoder();
        const entries = [];

        for (const f of files) {
            const raw     = enc.encode(f.data);
            const compressed = await deflate(f.data);
            const crc     = crc32(raw);
            const nameBytes = enc.encode(f.name);
            const offset  = 0; // sera corrigé
            entries.push({ name: nameBytes, raw, compressed, crc, size: raw.length });
        }

        // Calcule les offsets
        let pos = 0;
        const localOffsets = [];
        const parts = [];

        for (const e of entries) {
            localOffsets.push(pos);
            const header = [
                0x50, 0x4B, 0x03, 0x04,   // signature
                20, 0,                      // version needed
                0, 0,                       // flags
                8, 0,                       // compression (deflate)
                0, 0, 0, 0,                 // mod time/date
                ...u32le(e.crc),
                ...u32le(e.compressed.length),
                ...u32le(e.size),
                ...u16le(e.name.length),
                0, 0,                       // extra length
                ...e.name
            ];
            parts.push(new Uint8Array(header));
            parts.push(e.compressed);
            pos += header.length + e.compressed.length;
        }

        const cdStart = pos;
        for (let i = 0; i < entries.length; i++) {
            const e = entries[i];
            const cd = [
                0x50, 0x4B, 0x01, 0x02,
                20, 0, 20, 0,
                0, 0,
                8, 0,
                0, 0, 0, 0,
                ...u32le(e.crc),
                ...u32le(e.compressed.length),
                ...u32le(e.size),
                ...u16le(e.name.length),
                0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
                ...u32le(localOffsets[i]),
                ...e.name
            ];
            parts.push(new Uint8Array(cd));
            pos += cd.length;
        }

        const cdSize = pos - cdStart;
        const eocd = [
            0x50, 0x4B, 0x05, 0x06,
            0, 0, 0, 0,
            ...u16le(entries.length),
            ...u16le(entries.length),
            ...u32le(cdSize),
            ...u32le(cdStart),
            0, 0
        ];
        parts.push(new Uint8Array(eocd));

        const total = parts.reduce((s, p) => s + p.length, 0);
        const out = new Uint8Array(total);
        let offset = 0;
        for (const p of parts) { out.set(p, offset); offset += p.length; }
        return out;
    }

    // ── Export principal ─────────────────────────────────────────────────────
    async function exportToExcel() {
        const data = extractData();
        if (!data.length) {
            alert('Aucune donnée à exporter. Assurez-vous que les remboursements sont affichés.');
            return;
        }

        const sheetXml = buildSheetXml(data);

        const files = [
            { name: '[Content_Types].xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml"          ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml"            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>` },
            { name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>` },
            { name: 'xl/workbook.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Remboursements Santé" sheetId="1" r:id="rId1"/></sheets>
</workbook>` },
            { name: 'xl/_rels/workbook.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"   Target="styles.xml"/>
</Relationships>` },
            { name: 'xl/styles.xml',            data: STYLES_XML },
            { name: 'xl/worksheets/sheet1.xml', data: sheetXml },
        ];

        const zipBytes = await buildZip(files);
        const blob  = new Blob([zipBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url   = URL.createObjectURL(blob);
        const a     = document.createElement('a');
        const today = new Date().toISOString().slice(0, 10);
        a.href = url; a.download = `IPECA_Remboursements_${today}.xlsx`;
        a.click();
        URL.revokeObjectURL(url);
    }

    W.__ipecaExtractData = extractData; // utilisable depuis la console : __ipecaExtractData()

    // ── Lancement ────────────────────────────────────────────────────────────
    const observer = new MutationObserver(() => injectButton(false));
    observer.observe(document.body, { childList: true, subtree: true });
    injectButton(false);
    window.addEventListener('load', () => {
        injectButton(false);
        // Si rien n'a pu être injecté après 2 s, on utilise le bouton flottant
        setTimeout(() => injectButton(true), 2000);
    });

})();
