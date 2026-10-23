"""Read saved Excel calculation values, never pivot tables or execute formulas."""
from io import BytesIO
from math import isfinite
from decimal import Decimal
import re
import zipfile
import xml.etree.ElementTree as ET

import openpyxl

GROUPS = ('FREZ', 'LASER', 'MANUALNE', 'SZTANCA', 'GILOTYNA', 'FORMIERKA')
NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def read_workbook(data, sheet=None):
    with zipfile.ZipFile(BytesIO(data)) as archive:
        if sum(i.file_size for i in archive.infolist()) > 50 * 1024 * 1024:
            raise ValueError('Rozpakowany plik jest zbyt duży (limit 50 MB).')
    book = openpyxl.load_workbook(BytesIO(data), data_only=True, read_only=True)
    try:
        candidates = []
        for ws in book:
            if ws.title.upper() == 'TAB':
                continue
            for row in ws.iter_rows(max_row=min(ws.max_row, 100), max_col=13, values_only=True):
                if str(row[7]).strip().upper() == 'OPIS PROCESU' and 'RBH' in str(row[12]).upper():
                    candidates.append(ws.title)
                    break
        if not candidates:
            raise ValueError('Nie znaleziono arkusza wyceny z nagłówkami OPIS PROCESU (H) i ilość RBH (M).')
        if sheet is None and len(candidates) > 1:
            return {'sheets': candidates, 'selection_required': True}
        selected = sheet or candidates[0]
        if selected not in candidates:
            raise ValueError('Wybrany arkusz nie jest arkuszem wyceny.')
        ws = book[selected]
        quantity = ws['K4'].value
        if isinstance(quantity, bool) or not isinstance(quantity, (int, float)) or not isfinite(quantity) or quantity <= 0 or quantity != int(quantity):
            raise ValueError('Nakład w K4 musi być dodatnią całkowitą liczbą sztuk. Zapisz przeliczony plik w Excelu.')
        # Read raw numeric day fractions: datetime conversion rounds subsecond values.
        with zipfile.ZipFile(BytesIO(data)) as archive:
            root = ET.fromstring(archive.read(ws._worksheet_path))
        raw = {c.attrib['r']: c.find('m:v', NS).text for c in root.findall('.//m:c', NS)
               if c.find('m:v', NS) is not None and c.attrib.get('t', 'n') == 'n'}
        totals = dict.fromkeys(GROUPS, Decimal(0))
        details = []
        started = False
        setup = Decimal(0)
        for row in ws.iter_rows(max_col=13):
            label = str(row[7].value or '').strip()
            if label.upper() == 'OPIS PROCESU':
                started = True
                continue
            if not started:
                continue
            if not label:
                break
            group = re.split(r'\s*[-–—]\s*', label, maxsplit=1)[0].strip().upper()
            is_setup = group == 'RBH NA ROZRUCH PRODUKCJI'
            if is_setup:
                group = 'MANUALNE'
            if group not in totals:
                raise ValueError(f'Nieznana grupa procesu w H{row[7].row}: {label}. Wynik nie został obliczony.')
            address = row[12].coordinate
            try:
                hours = Decimal(raw[address]) * 24
                if not hours.is_finite() or hours < 0:
                    raise ValueError()
            except (KeyError, ValueError, ArithmeticError, TypeError):
                raise ValueError(f'Brak poprawnego czasu w {address}. Otwórz plik w Excelu, przelicz formuły i zapisz ponownie.') from None
            totals[group] += hours
            if is_setup:
                setup += hours
            details.append({'row': row[7].row, 'label': label, 'group': group, 'hours': float(hours)})
        if not details:
            raise ValueError('Tabela procesów jest pusta.')
        return {'sheets': candidates, 'sheet': selected, 'quantity': int(quantity),
                'project': str(ws['H4'].value or ''), 'setup_hours': float(setup),
                'processes': [{'name': g, 'hours': float(h), 'per_item': float(h / Decimal(str(quantity)))} for g, h in totals.items()],
                'details': details}
    finally:
        book.close()
