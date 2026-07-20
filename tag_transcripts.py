"""
tag_transcripts.py -- applies Codebook.xlsx highlight colors to matching
keywords across every Google Doc transcript in the Line-By-Line folder.

Setup:
    pip install openpyxl google-api-python-client google-auth-oauthlib

Usage:
    python tag_transcripts.py                # whole folder
    python tag_transcripts.py --doc-id XXXXX # single doc
"""

import argparse
import re

import openpyxl
from googleapiclient.discovery import build

CODEBOOK_PATH = "Codebook.xlsx"
LINE_BY_LINE_FOLDER_ID = "INSERT_FOLDER_ID"
SCOPES = ["https://www.googleapis.com/auth/documents",
          "https://www.googleapis.com/auth/drive.readonly"]


def load_codebook(path=CODEBOOK_PATH):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Sheet1"]
    codes = []
    for row in ws.iter_rows(min_row=9, max_row=300):
        code_cell, name_cell = row[0], row[1]
        if code_cell.value is None or name_cell.value is None:
            continue
        fill = code_cell.fill.fgColor.rgb if code_cell.fill else None
        if not fill or fill == "00000000":
            continue
        codes.append({"code_name": str(name_cell.value).strip(),
                       "rgb": tuple(int(fill[i:i + 2], 16) for i in (2, 4, 6))})
    return codes


def build_pattern(keyword):
    parts = keyword.split()
    last = parts[-1].lower()
    for suf in ("ing", "edly", "ed", "es", "s"):
        if last.endswith(suf) and len(last) - len(suf) >= 3:
            last = last[: -len(suf)]
            break
    variant = re.escape(last) + r"(?:s|es|d|ed|ing)?"
    if len(parts) > 1:
        prefix = r"\s+".join(re.escape(p) for p in parts[:-1])
        return r"\b" + prefix + r"\s+" + variant + r"\b"
    return r"\b" + variant + r"\b"


def get_credentials(credentials_path="credentials.json", token_path="token.json"):
    import os
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from google_auth_oauthlib.flow import InstalledAppFlow

    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w") as f:
            f.write(creds.to_json())
    return creds


def get_docs_in_folder(folder_id, creds):
    service = build("drive", "v3", credentials=creds)
    query = f"'{folder_id}' in parents and trashed=false and mimeType='application/vnd.google-apps.document'"
    items, page_token = [], None
    while True:
        results = service.files().list(
            q=query, spaces="drive", fields="nextPageToken, files(id, name)",
            pageToken=page_token, supportsAllDrives=True,
            includeItemsFromAllDrives=True, corpora="allDrives",
        ).execute()
        items.extend(results.get("files", []))
        page_token = results.get("nextPageToken")
        if not page_token:
            break
    return items


def tag_doc(doc_id, codes, docs_service):
    doc = docs_service.documents().get(documentId=doc_id).execute()
    text, offsets = "", []
    for element in doc.get("body", {}).get("content", []):
        for run in element.get("paragraph", {}).get("elements", []):
            tr = run.get("textRun")
            if not tr:
                continue
            content = tr.get("content", "")
            offsets.append((run.get("startIndex"), text.__len__(), content))
            text += content

    def to_doc_index(plain_idx):
        for start_idx, plain_start, content in offsets:
            if plain_start <= plain_idx < plain_start + len(content):
                return start_idx + (plain_idx - plain_start)
        return None

    requests = []
    for code in codes:
        pattern = build_pattern(code["code_name"])
        r, g, b = code["rgb"]
        for m in re.finditer(pattern, text, flags=re.IGNORECASE):
            start = to_doc_index(m.start())
            end = to_doc_index(m.end() - 1)
            if start is None or end is None:
                continue
            requests.append({
                "updateTextStyle": {
                    "range": {"startIndex": start, "endIndex": end + 1},
                    "textStyle": {"backgroundColor": {"color": {"rgbColor": {
                        "red": r / 255, "green": g / 255, "blue": b / 255}}}},
                    "fields": "backgroundColor",
                }
            })

    if requests:
        docs_service.documents().batchUpdate(
            documentId=doc_id, body={"requests": requests}).execute()
    return len(requests)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--folder-id", default=LINE_BY_LINE_FOLDER_ID)
    parser.add_argument("--doc-id", default=None)
    parser.add_argument("--codebook", default=CODEBOOK_PATH)
    args = parser.parse_args()

    codes = load_codebook(args.codebook)
    creds = get_credentials()
    docs_service = build("docs", "v1", credentials=creds)

    if args.doc_id:
        targets = [{"id": args.doc_id, "name": args.doc_id}]
    else:
        targets = get_docs_in_folder(args.folder_id, creds)
        print(f"Found {len(targets)} doc(s).")

    for d in targets:
        n = tag_doc(d["id"], codes, docs_service)
        print(f"{d['name']}: highlighted {n} span(s).")


if __name__ == "__main__":
    main()