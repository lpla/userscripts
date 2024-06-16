// ==UserScript==
// @name         Fill Marks from Excel
// @namespace    https://lpla.github.io/
// @version      0.1
// @description  Fill student marks in Itaca form from Excel file with specific name and mark format using exact and fuzzy matching for names
// @author       lpla
// @match        https://docent.edu.gva.es/md-front/www/*
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        none
// @updateURL   https://raw.githubusercontent.com/lpla/userscripts/main/icata.user.js
// @downloadURL https://raw.githubusercontent.com/lpla/userscripts/main/itaca.user.js
// @supportURL  https://github.com/lpla/userscripts/issues
// ==/UserScript==

(function() {
    'use strict';

    // Create a file input element to upload the Excel file
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx';
    input.style.position = 'fixed';
    input.style.top = '10px';
    input.style.left = '10px';
    input.style.zIndex = 1000;
    document.body.appendChild(input);

    input.addEventListener('change', handleFile, false);

    function handleFile(e) {
        const file = e.target.files[0];
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});

            // Assuming the first row contains headers and the first column is names and the second column is marks
            const studentData = jsonData.slice(1);

            // Function to fill the marks in the HTML structure
            fillMarks(studentData);
        };
        reader.readAsArrayBuffer(file);
    }

    function normalizeName(name) {
        return name
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/-/g, ' ')
            .trim()
            .toLowerCase()
            .replace(/\s+/g, ' ');
    }

    function extractMark(mark) {
        if (typeof mark === 'string' || mark instanceof String) {
            const match = mark.match(/^\d+/);
            return match ? match[0] : '';
        }
        return '';
    }

    function levenshtein(a, b) {
        const tmp = [];

        if (a.length === 0) { return b.length; }
        if (b.length === 0) { return a.length; }

        for (let i = 0; i <= b.length; i++) { tmp[i] = [i]; }
        for (let j = 0; j <= a.length; j++) { tmp[0][j] = j; }

        for (let i = 1; i <= b.length; i++) {
            for (let j = 1; j <= a.length; j++) {
                tmp[i][j] = b[i - 1] === a[j - 1] ?
                    tmp[i - 1][j - 1] :
                    Math.min(tmp[i - 1][j - 1] + 1, Math.min(tmp[i][j - 1] + 1, tmp[i - 1][j] + 1));
            }
        }

        return tmp[b.length][a.length];
    }

    function fillMarks(studentData) {
        studentData.forEach(([name, mark]) => {
            if (name && mark) {
                const formattedName = formatExcelName(name);
                const normalizedExcelName = normalizeName(formattedName);
                const extractedMark = extractMark(mark);
                const nameElements = document.evaluate(`//div[@class='imc-nom']/p`, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);

                let exactMatch = false;
                let bestMatch = null;
                let bestMatchDistance = Infinity;

                // First try exact match
                for (let i = 0; i < nameElements.snapshotLength; i++) {
                    const nameElement = nameElements.snapshotItem(i);
                    const normalizedHtmlName = normalizeName(nameElement.textContent);

                    if (normalizedExcelName === normalizedHtmlName) {
                        exactMatch = true;
                        bestMatch = nameElement;
                        bestMatchDistance = 0;
                        break;
                    }
                }

                // If no exact match, try "Surname1, Name1 Name2"
                if (!exactMatch) {
                    const alternativeFormat = formatAlternativeExcelName(name);
                    const normalizedAlternativeExcelName = normalizeName(alternativeFormat);

                    for (let i = 0; i < nameElements.snapshotLength; i++) {
                        const nameElement = nameElements.snapshotItem(i);
                        const normalizedHtmlName = normalizeName(nameElement.textContent);

                        if (normalizedAlternativeExcelName === normalizedHtmlName) {
                            exactMatch = true;
                            bestMatch = nameElement;
                            bestMatchDistance = 0;
                            break;
                        }

                        const distance = levenshtein(normalizedAlternativeExcelName, normalizedHtmlName);
                        if (distance < bestMatchDistance) {
                            bestMatch = nameElement;
                            bestMatchDistance = distance;
                        }
                    }
                }

                // Fuzzy match with relaxed threshold
                if (!exactMatch) {
                    for (let i = 0; i < nameElements.snapshotLength; i++) {
                        const nameElement = nameElements.snapshotItem(i);
                        const normalizedHtmlName = normalizeName(nameElement.textContent);

                        const distance = levenshtein(normalizedExcelName, normalizedHtmlName);
                        if (distance < bestMatchDistance) {
                            bestMatch = nameElement;
                            bestMatchDistance = distance;
                        }
                    }
                }

                if (bestMatch && bestMatchDistance <= 15) { // Adjusted threshold for fuzzy matching
                    const markInputElement = bestMatch.parentNode.nextElementSibling.querySelector('input');
                    if (markInputElement) {
                        markInputElement.value = extractedMark;
                        markInputElement.dispatchEvent(new Event('input', { bubbles: true }));
                    }
                }
            }
        });
    }

    function formatExcelName(name) {
        const parts = name.trim().split(/\s+/);
        if (parts.length > 2) {
            const firstName = parts.slice(0, parts.length - 2).join(' ');
            const lastNames = parts.slice(parts.length - 2).join(' ');
            return `${lastNames}, ${firstName}`;
        } else {
            const firstName = parts[0];
            const lastNames = parts.slice(1).join(' ');
            return `${lastNames}, ${firstName}`;
        }
    }

    function formatAlternativeExcelName(name) {
        const parts = name.trim().split(/\s+/);
        const firstName = parts.slice(0, -1).join(' ');
        const lastName = parts.slice(-1).join(' ');
        return `${lastName}, ${firstName}`;
    }
})();
