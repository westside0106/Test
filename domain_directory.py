#!/usr/bin/env python3
"""Simple domain directory management tool.

Allows adding domains manually or from CSV/TXT/XLSX files. Attempts to
populate information via WHOIS lookup. If network access is
unavailable the fields remain empty so they can be filled manually.
"""

import argparse
import csv
import json
import os
import sys
from dataclasses import dataclass, asdict
from datetime import datetime
from typing import List

# Optional: tiny XLSX reader using only the standard library
import zipfile
import xml.etree.ElementTree as ET

DATA_FILE = "domains.json"

@dataclass
class DomainInfo:
    domain: str
    appraisal_value: str = ""
    renewal_status: str = ""
    expiration_date: str = ""
    name_server: str = ""
    registration_date: str = ""
    timestamp: str = ""
    admin_email: str = ""  # remains empty
    domain_note: str = ""   # remains empty
    total_search_30d: str = ""
    total_search_180d: str = ""
    unique_search_30d: str = ""
    unique_search_180d: str = ""

class DomainDirectory:
    def __init__(self, path: str = DATA_FILE) -> None:
        self.path = path
        self.domains: List[DomainInfo] = []
        self.load()

    def load(self) -> None:
        if os.path.exists(self.path):
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)
                self.domains = [DomainInfo(**d) for d in data]

    def save(self) -> None:
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump([asdict(d) for d in self.domains], f, indent=2)

    def add(self, info: DomainInfo) -> None:
        self.domains.append(info)
        self.save()


def read_txt(file_path: str) -> List[str]:
    with open(file_path, "r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]


def read_csv(file_path: str) -> List[str]:
    domains = []
    with open(file_path, newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        for row in reader:
            if row:
                domains.append(row[0])
    return domains


def read_xlsx(file_path: str) -> List[str]:
    """Very small XLSX reader returning values of the first column."""
    domains = []
    with zipfile.ZipFile(file_path) as z:
        # read shared strings
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            with z.open('xl/sharedStrings.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                for si in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'):
                    shared.append(si.text or '')
        # read first worksheet
        sheet_name = 'xl/worksheets/sheet1.xml'
        if sheet_name not in z.namelist():
            return domains
        with z.open(sheet_name) as f:
            tree = ET.parse(f)
            root = tree.getroot()
            for row in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
                cells = row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')
                if cells:
                    cell = cells[0]
                    value = ''
                    if cell.get('t') == 's':
                        idx_elem = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                        if idx_elem is not None:
                            idx = int(idx_elem.text)
                            value = shared[idx]
                    else:
                        v_elem = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                        if v_elem is not None:
                            value = v_elem.text or ''
                    if value:
                        domains.append(value.strip())
    return domains


def fetch_whois_data(domain: str) -> DomainInfo:
    """Attempt to fetch WHOIS data. Returns DomainInfo with fields filled when possible."""
    info = DomainInfo(domain=domain)
    info.timestamp = datetime.utcnow().isoformat()
    try:
        import socket

        # find whois server via IANA
        s = socket.create_connection(('whois.iana.org', 43), timeout=5)
        s.sendall((domain + "\r\n").encode())
        response = b""
        while True:
            data = s.recv(4096)
            if not data:
                break
            response += data
        s.close()
        response_text = response.decode(errors='ignore')
        whois_server = None
        for line in response_text.splitlines():
            if line.lower().startswith('whois:'):
                whois_server = line.split(':', 1)[1].strip()
                break
        if whois_server:
            s = socket.create_connection((whois_server, 43), timeout=5)
            s.sendall((domain + "\r\n").encode())
            resp = b""
            while True:
                data = s.recv(4096)
                if not data:
                    break
                resp += data
            s.close()
            text = resp.decode(errors='ignore')
            for line in text.splitlines():
                if 'Creation Date' in line or 'Registered On' in line:
                    info.registration_date = line.split(':', 1)[1].strip()
                elif 'Expiry Date' in line or 'Expiration Date' in line:
                    info.expiration_date = line.split(':', 1)[1].strip()
                elif line.lower().startswith('name server'):
                    ns = line.split(':', 1)[1].strip()
                    if info.name_server:
                        info.name_server += ',' + ns
                    else:
                        info.name_server = ns
    except Exception:
        # network might be unavailable
        pass
    return info


def add_domains(dir_obj: DomainDirectory, domains: List[str]) -> None:
    for domain in domains:
        info = fetch_whois_data(domain)
        # Allow manual adjustment of other fields
        print(f"\nAdding domain: {domain}")
        def prompt(field_name: str, current: str) -> str:
            placeholder = f"Enter {field_name}"
            prompt_text = f"{field_name} [{current or placeholder}]: "
            value = input(prompt_text).strip()
            if value:
                return value
            return current
        info.appraisal_value = prompt('Appraisal Value', info.appraisal_value)
        info.renewal_status = prompt('Renewal Status', info.renewal_status)
        info.expiration_date = prompt('Expiration Date', info.expiration_date)
        info.name_server = prompt('Name Server', info.name_server)
        info.registration_date = prompt('Registration Date', info.registration_date)
        # timestamp already set
        # admin_email and domain_note intentionally left blank
        info.total_search_30d = prompt('Total Search Last 30 Days', info.total_search_30d)
        info.total_search_180d = prompt('Total Search Last 180 Days', info.total_search_180d)
        info.unique_search_30d = prompt('Unique Search Last 30 Days', info.unique_search_30d)
        info.unique_search_180d = prompt('Unique Search Last 180 Days', info.unique_search_180d)
        dir_obj.add(info)
        print("Domain stored.\n")


def parse_file(file_path: str) -> List[str]:
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.txt':
        return read_txt(file_path)
    if ext == '.csv':
        return read_csv(file_path)
    if ext in ('.xls', '.xlsx'):
        try:
            return read_xlsx(file_path)
        except Exception as e:
            print(f"Failed to read XLSX file: {e}")
            return []
    raise ValueError(f"Unsupported file type: {ext}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Manage domain directory")
    parser.add_argument('--add', nargs='*', help='Add domains manually')
    parser.add_argument('--file', help='Load domains from CSV/TXT/XLSX file')
    parser.add_argument('--list', action='store_true', help='List stored domains')
    args = parser.parse_args()

    directory = DomainDirectory()

    domains_to_add: List[str] = []
    if args.add:
        domains_to_add.extend(args.add)
    if args.file:
        domains_to_add.extend(parse_file(args.file))

    if domains_to_add:
        add_domains(directory, domains_to_add)

    if args.list:
        for d in directory.domains:
            print(asdict(d))

if __name__ == '__main__':
    main()
