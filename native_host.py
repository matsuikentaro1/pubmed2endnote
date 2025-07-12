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
            logging.error("Parsing failed: No authors (AU tag) found.")
            return False

        # 2. Build EndNote XML
        logging.info("Building EndNote XML.")
        logging.info(f"Parsed fields: {fields}")
        logging.info(f"Authors found: {authors}")
        
        # 基本フィールドの抽出
        pmid = fields.get('PMID', '')
        year = fields.get('DP', '')[:4] if fields.get('DP') else ''
        title = fields.get('TI', '')
        journal = fields.get('JT', '')  # The Tohoku journal of experimental medicine
        volume = fields.get('VI', '')
        pages = fields.get('PG', '')
        abstract = fields.get('AB', '')
        issue = fields.get('IP', '')
        place = fields.get('PL', '')
        journal_abbrev = fields.get('TA', '')  # Tohoku J Exp Med
        language = fields.get('LA', '')
        copyright_info = fields.get('CI', '')
        
        # デバッグ: 基本フィールドをログ出力
        logging.info(f"Raw JT field: '{fields.get('JT', 'NOT_FOUND')}'")
        logging.info(f"Raw TA field: '{fields.get('TA', 'NOT_FOUND')}'")
        logging.info(f"Extracted journal: '{journal}'")
        logging.info(f"Extracted journal_abbrev: '{journal_abbrev}'")
        
        # 日付情報
        edat = fields.get('EDAT', '')
        dep_date = fields.get('DEP', '')  # 電子出版日
        publication_status = fields.get('PST', '')
        
        # 複数値フィールドの処理（空値を除外）
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
        registry_numbers = get_list_field('RN')
        
        # ISSN処理
        issn_list = get_list_field('IS')
        issn_electronic = [issn.split()[0] for issn in issn_list if '(Electronic)' in issn]
        issn_linking = [issn.split()[0] for issn in issn_list if '(Linking)' in issn]
        
        # DOI/AID処理
        lid_list = get_list_field('LID')
        aid_list = get_list_field('AID')
        
        doi = ''
        pii = ''
        for lid in lid_list:
            if '[doi]' in lid:
                doi = lid.split()[0]
            elif '[pii]' in lid:
                pii = lid.split()[0]
        
        # AIDからも同様の情報を取得
        for aid in aid_list:
            if '[doi]' in aid and not doi:
                doi = aid.split()[0]
            elif '[pii]' in aid and not pii:
                pii = aid.split()[0]
        
        # メールアドレス抽出（アドレスフィールドから）
        emails = []
        for addr in addresses:
            email_matches = re.findall(r'Electronic address:\s*([\w\.-]+@[\w\.-]+)', addr)
            emails.extend(email_matches)
        
        # 著者情報
        first_author = authors[0].split(', ')[0] if authors else ''
        
        # デバッグ情報を追加
        logging.info(f"Authors: {authors}")
        logging.info(f"Journal: '{journal}', Journal Abbrev: '{journal_abbrev}'")
        logging.info(f"Volume: '{volume}', Issue: '{issue}', Pages: '{pages}'")
        logging.info(f"Addresses found: {len(addresses)} items")
        logging.info(f"Keywords found: {len(keywords)} items")
        logging.info(f"MeSH terms found: {len(mesh_terms)} items")
        logging.info(f"Publication types: {pub_types}")
        logging.info(f"Registry numbers: {len(registry_numbers)} items")
        logging.info(f"DOI: {doi}, PII: {pii}")
        logging.info(f"ISSN Electronic: {issn_electronic}, Linking: {issn_linking}")
        
        # NBIBの全フィールドをEndNote XMLに完全マッピング
        xml_parts = [
            "<EndNote><Cite>",
            f"<Author>{safe_escape_xml(first_author)}</Author>",
            f"<Year>{safe_escape_xml(year)}</Year>",
            f"<RecNum>{safe_escape_xml(pmid)}</RecNum>",
            "<DisplayText>(1)</DisplayText>",
            "<record>",
            f"<rec-number>{safe_escape_xml(pmid)}</rec-number>",
            "<ref-type name=\"Journal Article\">17</ref-type>",
        ]
        
        # 著者情報（FAUとAUの両方を保持）
        if authors:
            xml_parts.append("<contributors><authors>")
            for author in authors:
                xml_parts.append(f"<author>{safe_escape_xml(author)}</author>")
            xml_parts.append("</authors></contributors>")
        
        # タイトル
        if title:
            xml_parts.append(f"<titles><title>{safe_escape_xml(title)}</title></titles>")
        
        # 雑誌情報（JTとTA両方保持、ジャーナル名を確実に配置）
        if journal or journal_abbrev:
            xml_parts.append("<periodical>")
            if journal:
                xml_parts.append(f"<full-title>{safe_escape_xml(journal)}</full-title>")
            if journal_abbrev:
                xml_parts.append(f"<abbr-1>{safe_escape_xml(journal_abbrev)}</abbr-1>")
            xml_parts.append("</periodical>")
        
        # 補助的なジャーナル名フィールドも追加（互換性確保）
        if journal:
            xml_parts.append(f"<secondary-title>{safe_escape_xml(journal)}</secondary-title>")
        elif journal_abbrev:
            xml_parts.append(f"<secondary-title>{safe_escape_xml(journal_abbrev)}</secondary-title>")
        
        # 巻・号・ページ
        if volume:
            xml_parts.append(f"<volume>{safe_escape_xml(volume)}</volume>")
        if issue:
            xml_parts.append(f"<number>{safe_escape_xml(issue)}</number>")
        if pages:
            xml_parts.append(f"<pages>{safe_escape_xml(pages)}</pages>")
        
        # 日付情報
        if year:
            xml_parts.append(f"<dates><year>{safe_escape_xml(year)}</year></dates>")
        
        # 出版地
        if place:
            xml_parts.append(f"<pub-location>{safe_escape_xml(place)}</pub-location>")
        
        # DOI/LID
        if doi:
            xml_parts.append(f"<electronic-resource-num>{safe_escape_xml(doi)}</electronic-resource-num>")
        
        # アブストラクト
        if abstract:
            xml_parts.append(f"<abstract>{safe_escape_xml(abstract)}</abstract>")
        
        # 言語
        if language:
            xml_parts.append(f"<language>{safe_escape_xml(language)}</language>")
        
        # 著者所属（AD）
        if addresses:
            for addr in addresses:
                xml_parts.append(f"<author-address>{safe_escape_xml(addr)}</author-address>")
        
        # 論文タイプ（PT）- カスタムフィールドに移動してジャーナル名との混同を避ける
        if pub_types:
            for pt in pub_types:
                xml_parts.append(f"<custom2>{safe_escape_xml('Publication Type: ' + pt)}</custom2>")
        
        # MeSH用語（MH）
        if mesh_terms:
            for mesh in mesh_terms:
                xml_parts.append(f"<keyword>{safe_escape_xml(mesh)}</keyword>")
        
        # その他のキーワード（OT）
        if keywords:
            for keyword in keywords:
                xml_parts.append(f"<keyword>{safe_escape_xml(keyword)}</keyword>")
        
        # 化学物質登録番号（RN）
        if registry_numbers:
            for rn in registry_numbers:
                xml_parts.append(f"<research-notes>{safe_escape_xml(rn)}</research-notes>")
        
        # ISSN情報（IS）
        for issn in issn_list:
            xml_parts.append(f"<isbn>{safe_escape_xml(issn)}</isbn>")
        
        # NBIBの他の重要なフィールドを追加
        # PMID
        if pmid:
            xml_parts.append(f"<accession-num>{safe_escape_xml(pmid)}</accession-num>")
        
        # 雑誌ID（JID）
        jid = fields.get('JID', '')
        if jid:
            xml_parts.append(f"<call-num>{safe_escape_xml(jid)}</call-num>")
        
        # 状態情報（STAT, OWN, PST等）をカスタムフィールドとして保持
        for field_name in ['STAT', 'OWN', 'DCOM', 'LR', 'SB', 'EDAT', 'MHDA', 'CRDT', 'PHST', 'AID', 'PST', 'SO']:
            field_value = fields.get(field_name, '')
            if field_value:
                if isinstance(field_value, list):
                    for value in field_value:
                        xml_parts.append(f"<custom1>{safe_escape_xml(field_name + ': ' + value)}</custom1>")
                else:
                    xml_parts.append(f"<custom1>{safe_escape_xml(field_name + ': ' + field_value)}</custom1>")
        
        # 終了タグ
        xml_parts.extend([
            "</record>",
            "</Cite></EndNote>"
        ])
        xml_str = "".join(xml_parts)
        
        logging.info(f"Generated XML length: {len(xml_str)} characters")
        logging.info(f"Generated XML (first 1000 chars): {xml_str[:1000]}...")
        logging.info(f"Generated XML (last 500 chars): ...{xml_str[-500:]}")

        # 3. Generate RTF field
        logging.info("Generating RTF string.")
        rtf_full = (
            r"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}\pard\fs24 "
            + r"{\field{\*\fldinst { ADDIN EN.CITE " + xml_str + r" }}{\fldrslt (1)}}"
            + r"}"
        )
        
        logging.info(f"Generated RTF: {rtf_full[:300]}...")

        # 4. Copy to clipboard
        logging.info("Writing to clipboard.")
        rtf_format = win32clipboard.RegisterClipboardFormat("Rich Text Format")
        
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        
        # RTFをUTF-8でエンコードして、cp1252でデコードエラーを回避
        try:
            rtf_bytes = rtf_full.encode("cp1252", errors="replace")
        except:
            # フォールバック: ASCII文字のみを使用
            rtf_bytes = rtf_full.encode("ascii", errors="replace")
        
        win32clipboard.SetClipboardData(rtf_format, rtf_bytes)
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, "(1)".encode("utf-16le"))
        win32clipboard.CloseClipboard()
        logging.info("Clipboard write successful. Display text: (1)")
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
