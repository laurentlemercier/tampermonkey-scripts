// ==UserScript==
// @name         IPECA - Export Remboursements Excel
// @namespace    http://tampermonkey.net/
// @version      1.3
// @description  Ajoute un bouton pour exporter les remboursements santé IPECA vers Excel
// @author       Claude
// @match        https://www.ipeca.fr/mes-remboursements*
// @grant        none
// @require      https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js
// @updateURL    https://raw.githubusercontent.com/laurentlemercier/tampermonkey-scripts/main/ipeca.user.js
// @downloadURL  https://raw.githubusercontent.com/laurentlemercier/tampermonkey-scripts/main/ipeca.user.js
// @license      GPL-3.0-only
// ==/UserScript==

(function () {
    'use strict';

    // ── Injection du bouton ──────────────────────────────────────────────────
    function injectButton() {
        const pagination = document.querySelector('.wrapper-refund-sante .pagination');
        if (!pagination || document.getElementById('ipeca-export-btn')) return;

        const btn = document.createElement('button');
        btn.id = 'ipeca-export-btn';
        btn.textContent = '⬇ Exporter vers Excel';
        btn.style.cssText = `
            margin-left: 16px;
            padding: 6px 14px;
            background: #1F4E79;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 13px;
            font-family: Arial, sans-serif;
            vertical-align: middle;
        `;
        btn.addEventListener('mouseenter', () => btn.style.background = '#2E75B6');
        btn.addEventListener('mouseleave', () => btn.style.background = '#1F4E79');
        btn.addEventListener('click', exportToExcel);

        pagination.querySelector('.content-pagination').appendChild(btn);
    }

    // ── Extraction des données ───────────────────────────────────────────────
    function extractData() {
        const rows = [];
        const mainTable = document.querySelector('.wrapper-refund-sante table.ipeca-table');
        if (!mainTable) return rows;

        const parseAmount = (str) =>
            parseFloat((str || '').replace(/\s/g, '').replace('€', '').replace(',', '.')) || 0;

        // Récupère uniquement les <tr> enfants directs de la table principale
        // (évite de remonter les tr imbriqués dans les sous-tableaux)
        const topLevelRows = Array.from(mainTable.children[0]?.children || mainTable.querySelectorAll(':scope > tbody > tr, :scope > tr'));

        let currentMonth = '';

        topLevelRows.forEach(tr => {
            // ── Ligne de titre de mois ──────────────────────────────────────
            const monthCell = tr.querySelector(':scope > td.ipeca-table-title-row');
            if (monthCell) {
                currentMonth = monthCell.textContent.trim();
                return;
            }

            // ── Ligne de remboursement (contient un td colspan=7) ───────────
            if (!tr.classList.contains('align-td')) return;

            // Lecture de la ligne résumé (date, destinataire, montants globaux)
            const summaryRow = tr.querySelector('table.ipeca-content tr');
            if (!summaryRow) return;

            const sCells = summaryRow.querySelectorAll('td');
            if (sCells.length < 5) return;

            // Supprime les <span class="mobile"> pour ne garder que le texte utile
            const getText = (cell) => {
                const clone = cell.cloneNode(true);
                clone.querySelectorAll('span.mobile').forEach(s => s.remove());
                return clone.textContent.trim();
            };

            const dateRemb = getText(sCells[0]);
            const destinat = getText(sCells[1]);
            // Reconstruit la date complète : "27 MAR." + " 2026" depuis le mois courant
            const year = (currentMonth.match(/\d{4}/) || [''])[0];
            const dateRembFull = dateRemb + (year ? ' ' + year : '');

            // ── Sous-décomptes (1 ou plusieurs blocs sub-details) ───────────
            const subDetails = tr.querySelectorAll('td.sub-details');

            if (subDetails.length === 0) {
                // Pas de détail disponible : ligne agrégée de secours
                rows.push({
                    mois: currentMonth, dateRemb: dateRembFull, destinataire: destinat,
                    beneficiaire: '', dateSoin: '', nature: '',
                    fraisReels: parseAmount(getText(sCells[2])),
                    rembSS: 0, partOblig: 0,
                    ipeca: parseAmount(getText(sCells[3])),
                    reste: parseAmount(getText(sCells[4])),
                });
                return;
            }

            subDetails.forEach(sub => {
                const header = sub.querySelector('.header-sub-details')?.textContent.trim() || '';
                // "Soin du 23/03/2026 pour  LEIRE"
                const dateSoin = (header.match(/Soin du (\d{2}\/\d{2}\/\d{4})/) || [])[1] || '';
                const benef    = (header.match(/pour\s+([A-ZÉÀÈ]+)/i) || [])[1] || '';

                // Chaque ligne de détail est un <tr> dans le sous-tableau le plus profond
                // Structure : sub > table > tr > td > table > tr (lignes de détail)
                sub.querySelectorAll('table > tbody > tr, table > tr').forEach(detailTr => {
                    // On ne veut que les tr qui contiennent exactement 6 cellules de données
                    const dCells = detailTr.querySelectorAll(':scope > td');
                    if (dCells.length !== 6) return;

                    // Exclut les lignes "spacer" (hauteur fixe, pas de texte utile)
                    if (dCells[0].hasAttribute('height')) return;

                    const fraisReels = parseAmount(dCells[1].textContent);
                    rows.push({
                        mois:         currentMonth,
                        dateRemb:     dateRembFull,
                        destinataire: destinat,
                        beneficiaire: benef,
                        dateSoin:     dateSoin,
                        nature:       dCells[0].textContent.replace(/€/g, '').trim(),
                        fraisReels:   fraisReels,
                        rembSS:       parseAmount(dCells[2].textContent),
                        partOblig:    parseAmount(dCells[3].textContent),
                        ipeca:        parseAmount(dCells[4].textContent),
                        reste:        parseAmount(dCells[5].textContent),
                        regularisation: fraisReels < 0,  // ligne d'annulation si montant négatif
                    });
                });
            });
        });

        return rows;
    }

    // ── Génération du fichier Excel ──────────────────────────────────────────
    async function exportToExcel() {
        const data = extractData();
        if (!data.length) {
            alert('Aucune donnée à exporter. Assurez-vous que les remboursements sont affichés.');
            return;
        }

        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Remboursements Santé');

        // ── Styles réutilisables ────────────────────────────────────────────
        const headerFill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
        const orangeFill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFD7B0' } };
        const totalFill    = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBDD7EE' } };
        const headerFont   = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Arial', size: 10 };
        const boldFont     = { bold: true, name: 'Arial', size: 10 };
        const normalFont   = { name: 'Arial', size: 10 };
        const moneyFmt     = '#,##0.00 "€"';
        const centerAlign  = { horizontal: 'center', vertical: 'middle', wrapText: true };
        const rightAlign   = { horizontal: 'right',  vertical: 'middle' };
        const leftAlign    = { horizontal: 'left',   vertical: 'middle' };
        const thinBorder   = { style: 'thin', color: { argb: 'FFCCCCCC' } };
        const allBorders   = { top: thinBorder, left: thinBorder, bottom: thinBorder, right: thinBorder };

        // ── Largeurs de colonnes ────────────────────────────────────────────
        ws.columns = [
            { width: 14 }, { width: 22 }, { width: 22 }, { width: 14 },
            { width: 14 }, { width: 28 },
            { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 },
        ];

        // ── Ligne d'en-tête ─────────────────────────────────────────────────
        const headers = [
            'Mois', 'Date de remboursement', 'Destinataire', 'Bénéficiaire',
            'Date du soin', 'Nature du soin',
            'Frais réels (€)', 'Remboursé SS (€)', 'Part obligatoire (€)',
            'Versé IPECA (€)', 'Reste à charge (€)',
        ];
        const headerRow = ws.addRow(headers);
        headerRow.height = 35;
        headerRow.eachCell(cell => {
            cell.fill      = headerFill;
            cell.font      = headerFont;
            cell.alignment = centerAlign;
            cell.border    = allBorders;
        });

        // ── Lignes de données ───────────────────────────────────────────────
        data.forEach(r => {
            const row = ws.addRow([
                r.mois, r.dateRemb, r.destinataire, r.beneficiaire,
                r.dateSoin, r.nature,
                r.fraisReels, r.rembSS, r.partOblig, r.ipeca, r.reste,
            ]);

            const fill = r.regularisation ? orangeFill : null;

            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                cell.font   = normalFont;
                cell.border = allBorders;
                if (fill) cell.fill = fill;

                if (colNumber >= 7) {
                    cell.numFmt    = moneyFmt;
                    cell.alignment = rightAlign;
                } else {
                    cell.alignment = leftAlign;
                }
            });
        });

        // ── Ligne de totaux ─────────────────────────────────────────────────
        const lastDataRow = ws.rowCount; // dernière ligne de données
        const totalRow = ws.addRow([
            'TOTAL', '', '', '', '', '',
            { formula: `SUM(G2:G${lastDataRow})` },
            { formula: `SUM(H2:H${lastDataRow})` },
            { formula: `SUM(I2:I${lastDataRow})` },
            { formula: `SUM(J2:J${lastDataRow})` },
            { formula: `SUM(K2:K${lastDataRow})` },
        ]);
        totalRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            cell.font   = boldFont;
            cell.fill   = totalFill;
            cell.border = allBorders;
            if (colNumber >= 7) {
                cell.numFmt    = moneyFmt;
                cell.alignment = rightAlign;
            } else {
                cell.alignment = leftAlign;
            }
        });

        // ── Téléchargement ──────────────────────────────────────────────────
        const buffer = await wb.xlsx.writeBuffer();
        const blob   = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url    = URL.createObjectURL(blob);
        const a      = document.createElement('a');
        const today  = new Date().toISOString().slice(0, 10);
        a.href     = url;
        a.download = `IPECA_Remboursements_${today}.xlsx`;
        a.click();
        URL.revokeObjectURL(url);
    }

    // ── Lancement ────────────────────────────────────────────────────────────
    // Attend que la page soit chargée (la table peut être injectée dynamiquement)
    const observer = new MutationObserver(() => {
        if (document.querySelector('.wrapper-refund-sante .pagination')) {
            injectButton();
        }
    });
    observer.observe(document.body, { childList: true, subtree: true });

    // Tentative immédiate si déjà présent
    if (document.readyState === 'complete') {
        injectButton();
    } else {
        window.addEventListener('load', injectButton);
    }

})();
