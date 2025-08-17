import sys
import json
import struct
import logging
import os
import re

# --- Constants ---
LOG_FILE_NAME = 'native_host.log'

# --- Setup Logging ---
# This is the first thing that runs.
# It clears the log file on every run to make debugging the latest attempt easy.
try:
    log_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), LOG_FILE_NAME)
    logging.basicConfig(filename=log_file_path, filemode='w', level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s',
                        encoding='utf-8')
except Exception as e:
    # This is a last resort if logging itself fails.
    pass

# --- Helper Functions ---

def safe_escape_xml(text):
    """A safe, dependency-free XML escaper to replace html.escape."""
    if not isinstance(text, str):
        text = str(text)
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&#39;")

def escape_rtf(text):
    """Escapes characters for RTF."""
    if not isinstance(text, str):
        text = str(text)
    return text.replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}')

def format_author_name(full_name):
    """Formats author name to Last, F. I."""
    try:
        parts = full_name.split(', ')
        if len(parts) == 2:
            last_name, first_names = parts
            # Split by space or hyphen for multiple first/middle names
            initials = ". ".join([name[0] for name in re.split('[ -]', first_names) if name])
            return f"{last_name}, {initials}."
        else:
            return full_name # Return original if format is unexpected
    except:
        return full_name

def get_message():
    """Reads a single message from stdin, with robust error logging."""
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
    """Sends a single message to stdout, with robust error logging."""
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

def process_and_copy(nbib_data):
    """Parses NBIB, builds RTF, and copies to the clipboard."""
    try:
        import win32clipboard
        import win32con
    except ImportError:
        logging.error("PYWIN32 NOT FOUND. This is a fatal error.")
        return False

    try:
        # 1. Parse NBIB data
        logging.info("Parsing NBIB data.")
        fields = {}
        authors = []
        
        # より堅牢なNBIBパース
        lines = nbib_data.split('\n')
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if not line:
                i += 1
                continue
                
            # タグとハイフンのパターンをチェック
            match = re.match(r'^([A-Z]{2,4})\s*-\s*(.*)', line)
            if match:
                tag, value = match.groups()
                
                # 継続行を読み込む
                i += 1
                while i < len(lines) and lines[i].startswith('      '):  # 6スペースで継続
                    value += ' ' + lines[i].strip()
                    i += 1
                
                # 値をクリーンアップ
                cleaned_value = ' '.join(value.strip().split())
                
                if tag == 'FAU':  # Full Author nameのみ使用
                    authors.append(cleaned_value)
                elif tag == 'AU':
                    # AUタグは処理するが別途保存（著者名の重複を避けるため）
                    pass  # FAUを優先するため、AUは無視
                else:
                    # 複数値を持つフィールドの場合はリストで保存
                    if tag in ['AD', 'OT', 'MH', 'PT', 'RN', 'AID', 'IS', 'LID']:
                        if tag not in fields:
                            fields[tag] = []
                        fields[tag].append(cleaned_value)
                    else:
                        fields[tag] = cleaned_value
            else:
                i += 1
        
        if not authors:
            logging.error("Parsing failed: No authors (FAU tag) found.")
            return False

        # 2. Build EndNote XML
        logging.info("Building EndNote XML.")
        
        # 基本フィールドの抽出とフォーマット
        pmid = fields.get('PMID', '')
        year = fields.get('DP', '')[:4] if fields.get('DP') else ''
        title = fields.get('TI', '').rstrip('.') # タイトル末尾のピリオドを削除
        
        # ジャーナル名の処理 (JTを優先、なければTAを使用)
        journal_full = fields.get('JT', '').title()
        journal_abbrev = fields.get('TA', '').title()
        journal = journal_full if journal_full else journal_abbrev

        volume = fields.get('VI', '')
        pages = fields.get('PG', '')
        abstract = fields.get('AB', '')
        issue = fields.get('IP', '')
        place = fields.get('PL', '')
        language = fields.get('LA', '')
        
        # 著者名のフォーマット
        formatted_authors = [format_author_name(name) for name in authors]
        first_author = formatted_authors[0].split(',')[0] if formatted_authors else ''

        # 日付関連の処理
        dp_field = fields.get('DP', '') #例: 2024 Dec
        month = ''
        if len(dp_field.split()) > 1:
            month = dp_field.split()[1]

        edat_field = fields.get('EDAT', '') #例: 2024/09/23 03:15
        epub_date = edat_field.split()[0] if edat_field else ''

        # 複数値フィールドの処理
        def get_list_field(field_name):
            field_value = fields.get(field_name, [])
            if isinstance(field_value, list):
                return [item for item in field_value if item.strip()]
            elif field_value and field_value.strip():
                return [field_value]
            else:
                return []

        addresses = get_list_field('AD')
        keywords = get_list_field('OT')
        mesh_terms = get_list_field('MH')
        pub_types = get_list_field('PT')
        
        # ISSN処理
        issn_list = get_list_field('IS')
        issn_linking = ''.join([issn.split()[0] for issn in issn_list if '(Linking)' in issn])
        issn_electronic = ''.join([issn.split()[0] for issn in issn_list if '(Electronic)' in issn])

        # DOI/AID処理
        lid_list = get_list_field('LID')
        doi = ''.join([lid.split()[0] for lid in lid_list if '[doi]' in lid])

        # その他のフィールド
        so_field = fields.get('SO', '')
        owner = fields.get('OWN', '') # Database Provider

        # 表示用テキストを生成
        display_text = f"({first_author}, {year})" if first_author and year else f"({pmid})"
        
        # XML構築開始
        xml_parts = [
            "<EndNote><Cite>",
            f"<Author>{safe_escape_xml(first_author)}</Author>",
            f"<Year>{safe_escape_xml(year)}</Year>",
            f"<RecNum>{safe_escape_xml(pmid)}</RecNum>",
            f"<DisplayText>{safe_escape_xml(display_text)}</DisplayText>",
            "<record>",
            f"<rec-number>{safe_escape_xml(pmid)}</rec-number>",
            "<ref-type name=\"Journal Article\">17</ref-type>",
        ]
        
        # 著者情報
        if formatted_authors:
            xml_parts.append("<contributors><authors>")
            for author in formatted_authors:
                xml_parts.append(f"<author>{safe_escape_xml(author)}</author>")
            xml_parts.append("</authors></contributors>")
        
        # タイトル
        if title:
            xml_parts.append(f"<titles><title>{safe_escape_xml(title)}</title></titles>")
        
        # 雑誌情報 (JTとTA両方に対応)
        if journal_full or journal_abbrev:
            xml_parts.append("<periodical>")
            if journal_full:
                xml_parts.append(f"<full-title>{safe_escape_xml(journal_full)}</full-title>")
            if journal_abbrev:
                xml_parts.append(f"<abbr-1>{safe_escape_xml(journal_abbrev)}</abbr-1>")
            xml_parts.append("</periodical>")

        # 補助的なジャーナル名フィールドも追加（互換性確保）
        secondary_title = journal_full if journal_full else journal_abbrev
        if secondary_title:
            xml_parts.append(f"<secondary-title>{safe_escape_xml(secondary_title)}</secondary-title>")
        
        # 巻・号・ページ
        if volume:
            xml_parts.append(f"<volume>{safe_escape_xml(volume)}</volume>")
        if issue:
            xml_parts.append(f"<number>{safe_escape_xml(issue)}</number>")
        if pages:
            xml_parts.append(f"<pages>{safe_escape_xml(pages)}</pages>")
        
        # 日付情報
        xml_parts.append("<dates>")
        if year:
            xml_parts.append(f"<year>{safe_escape_xml(year)}</year>")
        if month:
             xml_parts.append(f"<pub-dates><date>{safe_escape_xml(month)}</date></pub-dates>")
        xml_parts.append("</dates>")

        # ISSN (Linking)
        if issn_linking:
            xml_parts.append(f"<isbn>{safe_escape_xml(issn_linking)}</isbn>")

        # DOI
        if doi:
            xml_parts.append(f"<electronic-resource-num>{safe_escape_xml(doi)}</electronic-resource-num>")
        
        # アブストラクト
        if abstract:
            xml_parts.append(f"<abstract>{safe_escape_xml(abstract)}</abstract>")
        
        # キーワード (MeSH + OT)
        all_keywords = mesh_terms + keywords
        if all_keywords:
            xml_parts.append("<keywords>")
            for keyword in all_keywords:
                xml_parts.append(f"<keyword>{safe_escape_xml(keyword)}</keyword>")
            xml_parts.append("</keywords>")

        # Notesフィールドの構築
        notes_content = []
        if issn_electronic:
            notes_content.append(issn_electronic)
        notes_content.extend(authors) # 元のフルネームの著者リスト
        notes_content.extend(pub_types)
        if so_field:
            notes_content.append(so_field)
        if place:
            notes_content.append(place)
        if addresses:
            # Author AddressをNotesではなく専用フィールドに
            full_address = ' '.join(addresses)
            xml_parts.append(f"<author-address>{safe_escape_xml(full_address)}</author-address>")

        if notes_content:
            notes_text = '\n'.join(notes_content)
            xml_parts.append(f"<notes>{safe_escape_xml(notes_text)}</notes>")

        # Accession Number (PMID)
        if pmid:
            xml_parts.append(f"<accession-num>{safe_escape_xml(pmid)}</accession-num>")

        # Epub Date (Custom4)
        if epub_date:
            xml_parts.append(f"<custom4>{safe_escape_xml(epub_date)}</custom4>")

        # Database Provider (Custom3)
        if owner:
            xml_parts.append(f"<custom3>{safe_escape_xml(owner)}</custom3>")
        
        # 言語
        if language:
            xml_parts.append(f"<language>{safe_escape_xml(language)}</language>")

        # 終了タグ
        xml_parts.extend([
            "</record>",
            "</Cite></EndNote>"
        ])
        xml_str = "".join(xml_parts)
        
        logging.info(f"Generated XML length: {len(xml_str)} characters")

        # 3. Generate RTF field
        logging.info("Generating RTF string.")
        rtf_full = (
            r"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}\pard\fs24 "
            + r"{\field{\*\fldinst { ADDIN EN.CITE " + xml_str + r" }}{\fldrslt " + escape_rtf(display_text) + r"}}"
            + r"}"
        )
        
        logging.info(f"Generated RTF: {rtf_full[:300]}...")

        # 4. Copy to clipboard
        logging.info("Writing to clipboard.")
        rtf_format = win32clipboard.RegisterClipboardFormat("Rich Text Format")
        
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        
        try:
            rtf_bytes = rtf_full.encode("cp1252", errors="replace")
        except:
            rtf_bytes = rtf_full.encode("ascii", errors="replace")
        
        win32clipboard.SetClipboardData(rtf_format, rtf_bytes)
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, display_text.encode("utf-16le"))
        win32clipboard.CloseClipboard()
        logging.info(f"Clipboard write successful. Display text: {display_text}")
        return True

    except Exception as e:
        logging.error(f"FATAL: Unhandled exception in process_and_copy. Error: {e}", exc_info=True)
        try: win32clipboard.CloseClipboard() # Attempt to close clipboard on error
        except: pass
        return False

# --- Main Execution Loop ---

def main():
    logging.info("--- Python native host started successfully. ---")
    try:
        message = get_message()
        if message and 'nbib_data' in message:
            if process_and_copy(message['nbib_data']):
                send_message({"status": "success"})
            else:
                send_message({"status": "error", "message": "Processing failed. See log."})
        elif message:
            logging.error(f"Invalid message received: {message}")
            send_message({"status": "error", "message": "Invalid message format."})
        # If message is None, it means stdin closed, so we just exit gracefully.

    except Exception as e:
        logging.critical(f"CRITICAL: Unhandled exception in main(). Error: {e}", exc_info=True)
        # Attempt to send a final error message if possible
        send_message({"status": "error", "message": f"A critical error occurred: {e}"})
    finally:
        logging.info("--- Python native host is shutting down. ---")

if __name__ == '__main__':
    main()
