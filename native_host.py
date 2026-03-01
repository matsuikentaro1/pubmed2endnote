import sys
import json
import struct
import logging
import os
import re
import xml.etree.ElementTree as ET

# --- Constants ---
LOG_FILE_NAME = 'native_host.log'

# --- Setup Logging ---
try:
    log_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), LOG_FILE_NAME)
    logging.basicConfig(filename=log_file_path, filemode='w', level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s',
                        encoding='utf-8')
except Exception as e:
    pass

# --- Helper Functions ---

def safe_escape_xml(text):
    """A safe, dependency-free XML escaper."""
    if not isinstance(text, str):
        text = str(text)
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&#39;")

def escape_rtf(text):
    """Escapes text for RTF with full Unicode support via \\uN? notation."""
    if not isinstance(text, str):
        text = str(text)
    result = []
    for char in text:
        if char == '\\':
            result.append('\\\\')
        elif char == '{':
            result.append('\\{')
        elif char == '}':
            result.append('\\}')
        elif ord(char) < 128:
            result.append(char)
        else:
            code = ord(char)
            if code > 32767:
                code -= 65536  # RTF uses signed 16-bit integers
            result.append(f'\\u{code}?')
    return ''.join(result)

def format_initials(initials_str):
    """Converts PubMed Initials field 'JA' to 'J. A.'"""
    clean = re.sub(r'[.\s]', '', initials_str.strip())
    if not clean:
        return ''
    return '. '.join(list(clean)) + '.'

def format_journal_sentence_case(journal_name):
    """Converts journal name to sentence case."""
    if not journal_name:
        return ''
    journal_name = journal_name.strip()
    if not journal_name:
        return ''
    return journal_name[0].upper() + journal_name[1:].lower()

def get_message():
    """Reads a single message from stdin."""
    try:
        raw_length = sys.stdin.buffer.read(4)
        if not raw_length:
            logging.info("Stdin pipe closed. Host is exiting.")
            return None

        message_length = struct.unpack('@I', raw_length)[0]
        logging.info(f"Attempting to read a message of length: {message_length}")

        message_body = sys.stdin.buffer.read(message_length)
        message_text = message_body.decode('utf-8')
        logging.info("Message received successfully.")
        return json.loads(message_text)

    except Exception as e:
        logging.error(f"CRITICAL: Failed to read or parse message from Chrome. Error: {e}", exc_info=True)
        return None

def send_message(message_dict):
    """Sends a single message to stdout."""
    try:
        message_text = json.dumps(message_dict)
        message_bytes = message_text.encode('utf-8')
        sys.stdout.buffer.write(struct.pack('@I', len(message_bytes)))
        sys.stdout.buffer.write(message_bytes)
        sys.stdout.buffer.flush()
        logging.info(f"Response sent to Chrome: {message_text}")
    except Exception as e:
        logging.error(f"CRITICAL: Failed to send message to Chrome. Error: {e}", exc_info=True)

# --- Core Logic ---

def parse_pubmed_xml(xml_data):
    """Parses PubMed XML and extracts bibliographic fields and author list."""
    root = ET.fromstring(xml_data)

    article = root.find('.//PubmedArticle')
    if article is None:
        logging.error("No PubmedArticle element found in XML.")
        return None, []

    medline = article.find('MedlineCitation')
    pubmed_data_elem = article.find('PubmedData')

    if medline is None:
        logging.error("No MedlineCitation element found.")
        return None, []

    art = medline.find('Article')
    if art is None:
        logging.error("No Article element found.")
        return None, []

    fields = {}

    # PMID
    fields['PMID'] = (medline.findtext('PMID') or '').strip()

    # Article Title (may contain nested tags like <i>, <sup>)
    title_elem = art.find('ArticleTitle')
    fields['TI'] = ''.join(title_elem.itertext()).strip().rstrip('.') if title_elem is not None else ''

    # Journal
    journal_elem = art.find('Journal')
    if journal_elem is not None:
        fields['JT'] = (journal_elem.findtext('Title') or '').strip()
        fields['TA'] = (journal_elem.findtext('ISOAbbreviation') or '').strip()

        for issn_elem in journal_elem.findall('ISSN'):
            issn_type = issn_elem.get('IssnType', '')
            val = (issn_elem.text or '').strip()
            if issn_type == 'Electronic':
                fields['IS_Electronic'] = val
            elif issn_type == 'Print':
                fields['IS_Print'] = val

        journal_issue = journal_elem.find('JournalIssue')
        if journal_issue is not None:
            fields['VI'] = (journal_issue.findtext('Volume') or '').strip()
            fields['IP'] = (journal_issue.findtext('Issue') or '').strip()

            pub_date = journal_issue.find('PubDate')
            if pub_date is not None:
                year = (pub_date.findtext('Year') or '').strip()
                month = (pub_date.findtext('Month') or '').strip()
                medline_date = (pub_date.findtext('MedlineDate') or '').strip()
                if year:
                    fields['DP_year'] = year
                    fields['DP_month'] = month
                elif medline_date:
                    parts = medline_date.split()
                    fields['DP_year'] = parts[0] if parts else ''
                    fields['DP_month'] = parts[1] if len(parts) > 1 else ''

    fields.setdefault('DP_year', '')
    fields.setdefault('DP_month', '')

    # Pagination
    fields['PG'] = (art.findtext('Pagination/MedlinePgn') or '').strip()

    # Abstract (may be structured with multiple AbstractText sections)
    abstract_elem = art.find('Abstract')
    if abstract_elem is not None:
        parts = []
        for text_elem in abstract_elem.findall('AbstractText'):
            label = text_elem.get('Label', '')
            text = ''.join(text_elem.itertext()).strip()
            if label and text:
                parts.append(f"{label}: {text}")
            elif text:
                parts.append(text)
        fields['AB'] = ' '.join(parts)
    else:
        fields['AB'] = ''

    # Authors
    authors = []
    author_list = art.find('AuthorList')
    if author_list is not None:
        for author_elem in author_list.findall('Author'):
            last = (author_elem.findtext('LastName') or '').strip()
            fore = (author_elem.findtext('ForeName') or '').strip()
            initials = (author_elem.findtext('Initials') or '').strip()
            collective = (author_elem.findtext('CollectiveName') or '').strip()

            if last:
                full = f"{last}, {fore}" if fore else last
                formatted = f"{last}, {format_initials(initials)}" if initials else last
                authors.append({'full': full, 'formatted': formatted, 'last': last})
            elif collective:
                authors.append({'full': collective, 'formatted': collective, 'last': collective})

    # Affiliations / Addresses
    addresses = []
    if author_list is not None:
        seen = set()
        for author_elem in author_list.findall('Author'):
            for aff in author_elem.findall('.//Affiliation'):
                text = (aff.text or '').strip()
                if text and text not in seen:
                    addresses.append(text)
                    seen.add(text)
    fields['AD'] = addresses

    # Language
    fields['LA'] = (art.findtext('Language') or '').strip()

    # Publication Types
    fields['PT'] = [(pt.text or '').strip() for pt in art.findall('.//PublicationType') if pt.text]

    # Keywords (Author keywords)
    fields['OT'] = [(kw.text or '').strip() for kw in medline.findall('.//Keyword') if kw.text]

    # MeSH Terms
    fields['MH'] = [(d.text or '').strip() for d in medline.findall('.//DescriptorName') if d.text]

    # Country
    fields['PL'] = (medline.findtext('.//Country') or '').strip()

    # DOI — check PubmedData first, then ELocationID
    fields['DOI'] = ''
    if pubmed_data_elem is not None:
        for article_id in pubmed_data_elem.findall('.//ArticleId'):
            if article_id.get('IdType') == 'doi' and article_id.text:
                fields['DOI'] = article_id.text.strip()
                break
    if not fields['DOI']:
        for eloc in art.findall('.//ELocationID'):
            if eloc.get('EIdType') == 'doi' and eloc.text:
                fields['DOI'] = eloc.text.strip()
                break

    # Electronic publication date: ArticleDate takes priority over History/entrez
    fields['EDAT'] = ''
    for art_date in art.findall('ArticleDate'):
        if art_date.get('DateType') == 'Electronic':
            y = art_date.findtext('Year') or ''
            m = art_date.findtext('Month') or ''
            d = art_date.findtext('Day') or ''
            if y:
                fields['EDAT'] = f"{y}/{m.zfill(2) if m else '00'}/{d.zfill(2) if d else '00'}"
            break
    if not fields['EDAT'] and pubmed_data_elem is not None:
        for pub_date in pubmed_data_elem.findall('.//PubMedPubDate'):
            if pub_date.get('PubStatus') == 'entrez':
                y = pub_date.findtext('Year') or ''
                m = pub_date.findtext('Month') or ''
                d = pub_date.findtext('Day') or ''
                if y:
                    fields['EDAT'] = f"{y}/{m.zfill(2) if m else '00'}/{d.zfill(2) if d else '00'}"
                break

    fields['OWN'] = 'NLM'

    return fields, authors


def process_and_copy(xml_data):
    """Parses PubMed XML, builds EndNote XML as RTF field, and copies to clipboard."""
    try:
        import win32clipboard
        import win32con
    except ImportError:
        logging.error("PYWIN32 NOT FOUND. This is a fatal error.")
        return False

    try:
        # 1. Parse PubMed XML
        logging.info("Parsing PubMed XML data.")
        fields, authors = parse_pubmed_xml(xml_data)

        if fields is None or not authors:
            logging.error("Parsing failed: could not parse XML or no authors found.")
            return False

        # 2. Extract fields
        pmid           = fields.get('PMID', '')
        year           = fields.get('DP_year', '')
        month          = fields.get('DP_month', '')
        title          = fields.get('TI', '')
        journal_full   = format_journal_sentence_case(fields.get('JT', ''))
        journal_abbrev = fields.get('TA', '').strip()
        volume         = fields.get('VI', '')
        issue          = fields.get('IP', '')
        pages          = fields.get('PG', '')
        abstract       = fields.get('AB', '')
        place          = fields.get('PL', '')
        language       = fields.get('LA', '')
        doi            = fields.get('DOI', '')
        epub_date      = fields.get('EDAT', '')
        issn_electronic = fields.get('IS_Electronic', '')
        issn_linking    = fields.get('IS_Print', '') or fields.get('IS_Electronic', '')
        keywords        = fields.get('OT', [])
        mesh_terms      = fields.get('MH', [])
        pub_types       = fields.get('PT', [])
        addresses       = fields.get('AD', [])

        formatted_authors   = [a['formatted'] for a in authors]
        full_name_authors   = [a['full'] for a in authors]
        first_author_last   = authors[0]['last'] if authors else ''

        display_text = f"({first_author_last}, {year})" if first_author_last and year else f"({pmid})"

        # Construct SO-equivalent source string for Notes
        so_parts = []
        if journal_abbrev:
            so_parts.append(journal_abbrev + '.')
        date_part = f" {year}"
        if month:
            date_part += f" {month}"
        if volume:
            date_part += f";{volume}"
            if issue:
                date_part += f"({issue})"
            if pages:
                date_part += f":{pages}"
        date_part += '.'
        so_parts.append(date_part.strip())
        so_field = ' '.join(so_parts)

        # 3. Build EndNote XML string
        logging.info("Building EndNote XML.")
        xml_parts = [
            "<EndNote><Cite>",
            f"<Author>{safe_escape_xml(first_author_last)}</Author>",
            f"<Year>{safe_escape_xml(year)}</Year>",
            f"<RecNum>{safe_escape_xml(pmid)}</RecNum>",
            f"<DisplayText>{safe_escape_xml(display_text)}</DisplayText>",
            "<record>",
            f"<rec-number>{safe_escape_xml(pmid)}</rec-number>",
            "<ref-type name=\"Journal Article\">17</ref-type>",
        ]

        if formatted_authors:
            xml_parts.append("<contributors><authors>")
            for author in formatted_authors:
                xml_parts.append(f"<author>{safe_escape_xml(author)}</author>")
            xml_parts.append("</authors></contributors>")

        if title:
            xml_parts.append(f"<titles><title>{safe_escape_xml(title)}</title></titles>")

        if journal_abbrev:
            xml_parts.append(f"<secondary-title>{safe_escape_xml(journal_abbrev)}</secondary-title>")
        if journal_full:
            xml_parts.append(f"<alt-title>{safe_escape_xml(journal_full)}</alt-title>")

        if journal_full or journal_abbrev:
            xml_parts.append("<periodical>")
            if journal_full:
                xml_parts.append(f"<full-title>{safe_escape_xml(journal_full)}</full-title>")
            if journal_abbrev:
                xml_parts.append(f"<abbr-1>{safe_escape_xml(journal_abbrev)}</abbr-1>")
            xml_parts.append("</periodical>")

        if volume:
            xml_parts.append(f"<volume>{safe_escape_xml(volume)}</volume>")
        if issue:
            xml_parts.append(f"<number>{safe_escape_xml(issue)}</number>")
        if pages:
            xml_parts.append(f"<pages>{safe_escape_xml(pages)}</pages>")

        xml_parts.append("<dates>")
        if year:
            xml_parts.append(f"<year>{safe_escape_xml(year)}</year>")
        if month:
            xml_parts.append(f"<pub-dates><date>{safe_escape_xml(month)}</date></pub-dates>")
        xml_parts.append("</dates>")

        if issn_linking:
            xml_parts.append(f"<isbn>{safe_escape_xml(issn_linking)}</isbn>")
        if doi:
            xml_parts.append(f"<electronic-resource-num>{safe_escape_xml(doi)}</electronic-resource-num>")
        if abstract:
            xml_parts.append(f"<abstract>{safe_escape_xml(abstract)}</abstract>")

        all_keywords = mesh_terms + keywords
        if all_keywords:
            xml_parts.append("<keywords>")
            for kw in all_keywords:
                xml_parts.append(f"<keyword>{safe_escape_xml(kw)}</keyword>")
            xml_parts.append("</keywords>")

        if addresses:
            full_address = ' '.join(addresses)
            xml_parts.append(f"<author-address>{safe_escape_xml(full_address)}</author-address>")

        notes_content = []
        if issn_electronic:
            notes_content.append(issn_electronic)
        notes_content.extend(full_name_authors)
        notes_content.extend(pub_types)
        if place:
            notes_content.append(place)
        if so_field:
            notes_content.append(so_field)
        if notes_content:
            notes_text = '&#10;'.join(notes_content)
            xml_parts.append(f"<notes>{notes_text}</notes>")

        if pmid:
            xml_parts.append(f"<accession-num>{safe_escape_xml(pmid)}</accession-num>")
        if epub_date:
            xml_parts.append(f"<custom4>{safe_escape_xml(epub_date)}</custom4>")
        if fields.get('OWN'):
            xml_parts.append(f"<custom3>{safe_escape_xml(fields['OWN'])}</custom3>")
        if language:
            xml_parts.append(f"<language>{safe_escape_xml(language)}</language>")

        xml_parts.extend(["</record>", "</Cite></EndNote>"])
        xml_str = "".join(xml_parts)

        logging.info(f"Generated XML length: {len(xml_str)} characters")

        # 4. Generate RTF field
        # escape_rtf converts non-ASCII chars to \uN? notation, so the final RTF is pure ASCII.
        logging.info("Generating RTF string.")
        rtf_full = (
            r"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}\pard\fs24 "
            + r"{\highlight7{\field{\*\fldinst { ADDIN EN.CITE " + escape_rtf(xml_str) + r" }}{\fldrslt " + escape_rtf(display_text) + r"}}}"
            + r"}"
        )

        logging.info(f"Generated RTF: {rtf_full[:300]}...")

        # 5. Copy to clipboard
        logging.info("Writing to clipboard.")
        rtf_format = win32clipboard.RegisterClipboardFormat("Rich Text Format")

        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()

        # RTF is pure ASCII after escape_rtf, so ASCII encoding is safe
        rtf_bytes = rtf_full.encode("ascii")

        win32clipboard.SetClipboardData(rtf_format, rtf_bytes)
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, display_text.encode("utf-16le"))
        win32clipboard.CloseClipboard()
        logging.info(f"Clipboard write successful. Display text: {display_text}")
        return True

    except Exception as e:
        logging.error(f"FATAL: Unhandled exception in process_and_copy. Error: {e}", exc_info=True)
        try:
            win32clipboard.CloseClipboard()
        except:
            pass
        return False

# --- Main Execution Loop ---

def main():
    logging.info("--- Python native host started successfully. ---")
    try:
        message = get_message()
        if message and 'xml_data' in message:
            if process_and_copy(message['xml_data']):
                send_message({"status": "success"})
            else:
                send_message({"status": "error", "message": "Processing failed. See log."})
        elif message:
            logging.error(f"Invalid message received: {message}")
            send_message({"status": "error", "message": "Invalid message format."})

    except Exception as e:
        logging.critical(f"CRITICAL: Unhandled exception in main(). Error: {e}", exc_info=True)
        send_message({"status": "error", "message": f"A critical error occurred: {e}"})
    finally:
        logging.info("--- Python native host is shutting down. ---")

if __name__ == '__main__':
    main()
