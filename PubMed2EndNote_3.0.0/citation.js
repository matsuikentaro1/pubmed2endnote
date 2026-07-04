// citation.js
// PubMed XML → EndNote traveling-library field → Word-flavored HTML clipboard payload.

// --- Escaping helpers ---

function escapeXml(text) {
    return String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function escapeHtml(text) {
    return String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

// --- Formatting helpers ---

// Converts PubMed Initials field 'JA' to 'J. A.'
function formatInitials(initialsStr) {
    const clean = initialsStr.trim().replace(/[.\s]/g, '');
    if (!clean) return '';
    return clean.split('').join('. ') + '.';
}

function journalSentenceCase(journalName) {
    const name = (journalName || '').trim();
    if (!name) return '';
    return name[0].toUpperCase() + name.slice(1).toLowerCase();
}

// --- PubMed XML parsing ---

function parsePubMedXml(xmlData) {
    const doc = new DOMParser().parseFromString(xmlData, 'text/xml');
    if (doc.querySelector('parsererror')) {
        throw new Error('Failed to parse PubMed XML');
    }

    const text = (parent, selector) => {
        if (!parent) return '';
        const el = parent.querySelector(selector);
        return el ? el.textContent.trim() : '';
    };

    const article = doc.querySelector('PubmedArticle');
    if (!article) throw new Error('No PubmedArticle element found in XML');

    const medline = article.querySelector(':scope > MedlineCitation');
    const pubmedData = article.querySelector(':scope > PubmedData');
    if (!medline) throw new Error('No MedlineCitation element found');

    const art = medline.querySelector(':scope > Article');
    if (!art) throw new Error('No Article element found');

    const fields = {};

    // PMID must be MedlineCitation's direct child (CommentsCorrections also contain PMIDs)
    fields.PMID = text(medline, ':scope > PMID');

    // Article Title (textContent flattens nested tags like <i>, <sup>)
    const titleElem = art.querySelector(':scope > ArticleTitle');
    fields.TI = titleElem ? titleElem.textContent.trim().replace(/\.+$/, '') : '';

    // Journal
    fields.DP_year = '';
    fields.DP_month = '';
    const journalElem = art.querySelector(':scope > Journal');
    if (journalElem) {
        fields.JT = text(journalElem, ':scope > Title');
        fields.TA = text(journalElem, ':scope > ISOAbbreviation');

        for (const issnElem of journalElem.querySelectorAll(':scope > ISSN')) {
            const val = issnElem.textContent.trim();
            const issnType = issnElem.getAttribute('IssnType') || '';
            if (issnType === 'Electronic') fields.IS_Electronic = val;
            else if (issnType === 'Print') fields.IS_Print = val;
        }

        const journalIssue = journalElem.querySelector(':scope > JournalIssue');
        if (journalIssue) {
            fields.VI = text(journalIssue, ':scope > Volume');
            fields.IP = text(journalIssue, ':scope > Issue');

            const pubDate = journalIssue.querySelector(':scope > PubDate');
            if (pubDate) {
                const year = text(pubDate, ':scope > Year');
                const month = text(pubDate, ':scope > Month');
                const medlineDate = text(pubDate, ':scope > MedlineDate');
                if (year) {
                    fields.DP_year = year;
                    fields.DP_month = month;
                } else if (medlineDate) {
                    const parts = medlineDate.split(/\s+/);
                    fields.DP_year = parts[0] || '';
                    fields.DP_month = parts[1] || '';
                }
            }
        }
    }

    // Pagination
    fields.PG = text(art, ':scope > Pagination > MedlinePgn');

    // Abstract (may be structured with labeled sections)
    const abstractElem = art.querySelector(':scope > Abstract');
    if (abstractElem) {
        const parts = [];
        for (const textElem of abstractElem.querySelectorAll(':scope > AbstractText')) {
            const label = textElem.getAttribute('Label') || '';
            const t = textElem.textContent.trim();
            if (label && t) parts.push(`${label}: ${t}`);
            else if (t) parts.push(t);
        }
        fields.AB = parts.join(' ');
    } else {
        fields.AB = '';
    }

    // Authors
    const authors = [];
    const authorList = art.querySelector(':scope > AuthorList');
    if (authorList) {
        for (const authorElem of authorList.querySelectorAll(':scope > Author')) {
            const last = text(authorElem, ':scope > LastName');
            const fore = text(authorElem, ':scope > ForeName');
            const initials = text(authorElem, ':scope > Initials');
            const collective = text(authorElem, ':scope > CollectiveName');

            if (last) {
                const full = fore ? `${last}, ${fore}` : last;
                const formatted = initials ? `${last}, ${formatInitials(initials)}` : last;
                authors.push({ full, formatted, last });
            } else if (collective) {
                authors.push({ full: collective, formatted: collective, last: collective });
            }
        }
    }

    // Affiliations / Addresses (deduplicated, in order of appearance)
    const addresses = [];
    if (authorList) {
        const seen = new Set();
        for (const authorElem of authorList.querySelectorAll(':scope > Author')) {
            for (const aff of authorElem.querySelectorAll('Affiliation')) {
                const t = aff.textContent.trim();
                if (t && !seen.has(t)) {
                    addresses.push(t);
                    seen.add(t);
                }
            }
        }
    }
    fields.AD = addresses;

    // Language
    fields.LA = text(art, ':scope > Language');

    // Publication Types
    fields.PT = [...art.querySelectorAll('PublicationType')]
        .map(pt => pt.textContent.trim()).filter(Boolean);

    // Keywords (author keywords)
    fields.OT = [...medline.querySelectorAll('Keyword')]
        .map(kw => kw.textContent.trim()).filter(Boolean);

    // MeSH terms
    fields.MH = [...medline.querySelectorAll('DescriptorName')]
        .map(d => d.textContent.trim()).filter(Boolean);

    // Country
    fields.PL = text(medline, 'Country');

    // DOI — check PubmedData first, then ELocationID
    fields.DOI = '';
    if (pubmedData) {
        for (const articleId of pubmedData.querySelectorAll('ArticleId')) {
            if (articleId.getAttribute('IdType') === 'doi' && articleId.textContent.trim()) {
                fields.DOI = articleId.textContent.trim();
                break;
            }
        }
    }
    if (!fields.DOI) {
        for (const eloc of art.querySelectorAll('ELocationID')) {
            if (eloc.getAttribute('EIdType') === 'doi' && eloc.textContent.trim()) {
                fields.DOI = eloc.textContent.trim();
                break;
            }
        }
    }

    // Electronic publication date: ArticleDate takes priority over History/entrez
    fields.EDAT = '';
    for (const artDate of art.querySelectorAll(':scope > ArticleDate')) {
        if (artDate.getAttribute('DateType') === 'Electronic') {
            const y = text(artDate, ':scope > Year');
            const m = text(artDate, ':scope > Month');
            const d = text(artDate, ':scope > Day');
            if (y) fields.EDAT = `${y}/${m ? m.padStart(2, '0') : '00'}/${d ? d.padStart(2, '0') : '00'}`;
            break;
        }
    }
    if (!fields.EDAT && pubmedData) {
        for (const pubDate of pubmedData.querySelectorAll('PubMedPubDate')) {
            if (pubDate.getAttribute('PubStatus') === 'entrez') {
                const y = text(pubDate, ':scope > Year');
                const m = text(pubDate, ':scope > Month');
                const d = text(pubDate, ':scope > Day');
                if (y) fields.EDAT = `${y}/${m ? m.padStart(2, '0') : '00'}/${d ? d.padStart(2, '0') : '00'}`;
                break;
            }
        }
    }

    fields.OWN = 'NLM';

    return { fields, authors };
}

// --- EndNote XML construction ---

function buildEndNoteXml(fields, authors) {
    const pmid = fields.PMID || '';
    const year = fields.DP_year || '';
    const month = fields.DP_month || '';
    const title = fields.TI || '';
    const journalFull = journalSentenceCase(fields.JT || '');
    const journalAbbrev = (fields.TA || '').trim();
    const volume = fields.VI || '';
    const issue = fields.IP || '';
    const pages = fields.PG || '';
    const abstract = fields.AB || '';
    const place = fields.PL || '';
    const language = fields.LA || '';
    const doi = fields.DOI || '';
    const epubDate = fields.EDAT || '';
    const issnElectronic = fields.IS_Electronic || '';
    const issnLinking = fields.IS_Print || fields.IS_Electronic || '';
    const keywords = fields.OT || [];
    const meshTerms = fields.MH || [];
    const pubTypes = fields.PT || [];
    const addresses = fields.AD || [];

    const formattedAuthors = authors.map(a => a.formatted);
    const fullNameAuthors = authors.map(a => a.full);
    const firstAuthorLast = authors.length ? authors[0].last : '';

    const displayText = (firstAuthorLast && year) ? `(${firstAuthorLast}, ${year})` : `(${pmid})`;

    // Construct SO-equivalent source string for Notes
    const soParts = [];
    if (journalAbbrev) soParts.push(journalAbbrev + '.');
    let datePart = ` ${year}`;
    if (month) datePart += ` ${month}`;
    if (volume) {
        datePart += `;${volume}`;
        if (issue) datePart += `(${issue})`;
        if (pages) datePart += `:${pages}`;
    }
    datePart += '.';
    soParts.push(datePart.trim());
    const soField = soParts.join(' ');

    const xmlParts = [
        '<EndNote><Cite>',
        `<Author>${escapeXml(firstAuthorLast)}</Author>`,
        `<Year>${escapeXml(year)}</Year>`,
        `<RecNum>${escapeXml(pmid)}</RecNum>`,
        `<DisplayText>${escapeXml(displayText)}</DisplayText>`,
        '<record>',
        `<rec-number>${escapeXml(pmid)}</rec-number>`,
        '<ref-type name="Journal Article">17</ref-type>',
    ];

    if (formattedAuthors.length) {
        xmlParts.push('<contributors><authors>');
        for (const author of formattedAuthors) {
            xmlParts.push(`<author>${escapeXml(author)}</author>`);
        }
        xmlParts.push('</authors></contributors>');
    }

    if (title) {
        xmlParts.push(`<titles><title>${escapeXml(title)}</title></titles>`);
    }

    if (journalAbbrev) {
        xmlParts.push(`<secondary-title>${escapeXml(journalAbbrev)}</secondary-title>`);
    }
    if (journalFull) {
        xmlParts.push(`<alt-title>${escapeXml(journalFull)}</alt-title>`);
    }

    if (journalFull || journalAbbrev) {
        xmlParts.push('<periodical>');
        if (journalFull) xmlParts.push(`<full-title>${escapeXml(journalFull)}</full-title>`);
        if (journalAbbrev) xmlParts.push(`<abbr-1>${escapeXml(journalAbbrev)}</abbr-1>`);
        xmlParts.push('</periodical>');
    }

    if (volume) xmlParts.push(`<volume>${escapeXml(volume)}</volume>`);
    if (issue) xmlParts.push(`<number>${escapeXml(issue)}</number>`);
    if (pages) xmlParts.push(`<pages>${escapeXml(pages)}</pages>`);

    xmlParts.push('<dates>');
    if (year) xmlParts.push(`<year>${escapeXml(year)}</year>`);
    if (month) xmlParts.push(`<pub-dates><date>${escapeXml(month)}</date></pub-dates>`);
    xmlParts.push('</dates>');

    if (issnLinking) xmlParts.push(`<isbn>${escapeXml(issnLinking)}</isbn>`);
    if (doi) xmlParts.push(`<electronic-resource-num>${escapeXml(doi)}</electronic-resource-num>`);
    if (abstract) xmlParts.push(`<abstract>${escapeXml(abstract)}</abstract>`);

    const allKeywords = [...meshTerms, ...keywords];
    if (allKeywords.length) {
        xmlParts.push('<keywords>');
        for (const kw of allKeywords) {
            xmlParts.push(`<keyword>${escapeXml(kw)}</keyword>`);
        }
        xmlParts.push('</keywords>');
    }

    if (addresses.length) {
        xmlParts.push(`<author-address>${escapeXml(addresses.join(' '))}</author-address>`);
    }

    const notesContent = [];
    if (issnElectronic) notesContent.push(issnElectronic);
    notesContent.push(...fullNameAuthors);
    notesContent.push(...pubTypes);
    if (place) notesContent.push(place);
    if (soField) notesContent.push(soField);
    if (notesContent.length) {
        // Items are XML-escaped, then joined with a literal &#10; entity so EndNote
        // renders line breaks inside the Notes field.
        const notesText = notesContent.map(escapeXml).join('&#10;');
        xmlParts.push(`<notes>${notesText}</notes>`);
    }

    if (pmid) xmlParts.push(`<accession-num>${escapeXml(pmid)}</accession-num>`);
    if (epubDate) xmlParts.push(`<custom4>${escapeXml(epubDate)}</custom4>`);
    if (fields.OWN) xmlParts.push(`<custom3>${escapeXml(fields.OWN)}</custom3>`);
    if (language) xmlParts.push(`<language>${escapeXml(language)}</language>`);

    xmlParts.push('</record>', '</Cite></EndNote>');

    return { xmlStr: xmlParts.join(''), displayText };
}

// --- Word-flavored HTML clipboard payload ---

// Word reconstructs an { ADDIN EN.CITE ... } field from mso-element markers
// inside <!--[if supportFields]--> conditional comments on paste.
// fontFamily / fontSize (optional) let the pasted citation match the manuscript;
// they apply when Word's paste option is "Keep Source Formatting".
function buildWordFieldHtml(xmlStr, displayText, { fontFamily = '', fontSize = '' } = {}) {
    const fieldInstr = escapeHtml(` ADDIN EN.CITE ${xmlStr} `);
    let spanStyle = 'background:yellow;mso-highlight:yellow';
    if (fontFamily) spanStyle += `;font-family:"${fontFamily}"`;
    if (fontSize) spanStyle += `;font-size:${fontSize}pt`;
    // No <p> wrapper: an inline fragment pastes into the middle of a sentence without
    // splitting the paragraph or bringing its own paragraph spacing along.
    return '<html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
        'xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">' +
        '<head><meta charset="utf-8"></head><body>' +
        // The highlight span wraps the ENTIRE field (like \highlight7 around \field in RTF),
        // so the yellow survives when EndNote's instant formatting regenerates the field result.
        `<span style='${spanStyle}'>` +
        "<!--[if supportFields]><span style='mso-element:field-begin'></span>" +
        fieldInstr +
        "<span style='mso-element:field-separator'></span><![endif]-->" +
        escapeHtml(displayText) +
        "<!--[if supportFields]><span style='mso-element:field-end'></span><![endif]-->" +
        '</span></body></html>';
}

// --- Clipboard ---

async function copyCitationToClipboard(html, plainText) {
    try {
        await navigator.clipboard.write([new ClipboardItem({
            'text/html': new Blob([html], { type: 'text/html' }),
            'text/plain': new Blob([plainText], { type: 'text/plain' }),
        })]);
        return;
    } catch (err) {
        console.warn('navigator.clipboard.write failed, falling back to copy event:', err);
    }

    const handler = (ev) => {
        ev.clipboardData.setData('text/html', html);
        ev.clipboardData.setData('text/plain', plainText);
        ev.preventDefault();
    };
    document.addEventListener('copy', handler, { once: true });
    const ok = document.execCommand('copy');
    document.removeEventListener('copy', handler);
    if (!ok) throw new Error('Clipboard write failed');
}

// --- PubMed API fetch ---
// E-utilities sends Access-Control-Allow-Origin: *, so the content script can
// fetch directly and no host permission is needed.

async function fetchPubMedXml(pmid, userEmail) {
    const url = `https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed&id=${pmid}&rettype=xml&retmode=xml&email=${encodeURIComponent(userEmail)}`;
    const response = await fetch(url);
    if (!response.ok) {
        throw new Error(`PubMed API error: ${response.status}`);
    }
    const xmlData = await response.text();
    if (!xmlData || xmlData.trim() === '') {
        throw new Error(`No data found for PMID: ${pmid}`);
    }
    return xmlData;
}

// --- Entry point used by content.js ---

async function convertAndCopy(xmlData, formatOptions = {}) {
    const { fields, authors } = parsePubMedXml(xmlData);
    if (!authors.length) {
        throw new Error('No authors found in PubMed record');
    }
    const { xmlStr, displayText } = buildEndNoteXml(fields, authors);
    const html = buildWordFieldHtml(xmlStr, displayText, formatOptions);
    await copyCitationToClipboard(html, displayText);
    return displayText;
}
