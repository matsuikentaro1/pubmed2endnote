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
                
                if tag == 'FAU':  # Full Author nameを優先
                    authors.append(cleaned_value)
                elif tag == 'AU' and cleaned_value not in [a.replace(', ', ' ') for a in authors]:
                    # FAUがない場合のみAUを使用
                    authors.append(cleaned_value)
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
        
        pmid = fields.get('PMID', '')
        year = fields.get('DP', '')[:4]
        title = fields.get('TI', '')
        journal = fields.get('JT', '')
        volume = fields.get('VI', '')
        pages = fields.get('PG', '')
        abstract = fields.get('AB', '')
        aid = fields.get('AID', '')
        
        # SendClipboard.pyと同じ正確なXML構造
        first_author = authors[0].split(', ')[0] if authors else ''
        
        # 追加データの抽出（SendClipboard.pyと同様）
        keywords = re.findall(r'OT  - ([^\n]+)', nbib_data)
        addresses = re.findall(r'AD  - ([^\n]+)', nbib_data)
        emails = re.findall(r'Electronic address:\s*([\w\.-]+@[\w\.-]+)', nbib_data)
        
        xml_parts = [
            "<EndNote><Cite>",
            f"<Author>{safe_escape_xml(first_author)}</Author>",
            f"<Year>{safe_escape_xml(year)}</Year>",
            f"<RecNum>{safe_escape_xml(pmid)}</RecNum>",
            "<DisplayText>(1)</DisplayText>",
            "<record>",
            f"<rec-number>{safe_escape_xml(pmid)}</rec-number>",
            "<ref-type name=\"Journal Article\">17</ref-type>",
            "<contributors><authors>",
            *[f"<author>{safe_escape_xml(a)}</author>" for a in authors],
            "</authors></contributors>",
            f"<titles><title>{safe_escape_xml(title)}</title></titles>",
            f"<periodical><full-title>{safe_escape_xml(journal)}</full-title></periodical>",
            f"<pages>{safe_escape_xml(pages)}</pages>",
            f"<volume>{safe_escape_xml(volume)}</volume>",
            f"<dates><year>{safe_escape_xml(year)}</year></dates>",
            f"<electronic-resource-num>{safe_escape_xml(aid)}</electronic-resource-num>",
            f"<abstract>{safe_escape_xml(abstract)}</abstract>",
            *[f"<keywords><keyword>{safe_escape_xml(k)}</keyword></keywords>" for k in keywords],
            *[f"<author-address>{safe_escape_xml(addr)}</author-address>" for addr in addresses],
            *[f"<email>{e}</email>" for e in emails],
            "</record>",
            "</Cite></EndNote>"
        ]
        xml_str = "".join(xml_parts)
        
        logging.info(f"Generated XML: {xml_str[:500]}...")

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
