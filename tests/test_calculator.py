import unittest
from io import BytesIO
import openpyxl
from calculator import read_workbook


def workbook(quantity=10, label='FREZ - cięcie', time=1.5, extra=False):
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = 'Wycena'
    sheet['K4'] = quantity
    sheet['H9'] = 'OPIS PROCESU'
    sheet['M9'] = 'ilość RBH'
    sheet['H10'] = label
    sheet['M10'] = time
    sheet['H11'] = 'RBH NA ROZRUCH PRODUKCJI'
    sheet['M11'] = 2 / 24
    if extra:
        book.copy_worksheet(sheet).title = 'Druga'
    out = BytesIO()
    book.save(out)
    return out.getvalue()


class CalculatorTests(unittest.TestCase):
    def test_days_over_24_hours_and_setup(self):
        result = read_workbook(workbook())
        values = {p['name']: p for p in result['processes']}
        self.assertAlmostEqual(values['FREZ']['hours'], 36)
        self.assertAlmostEqual(values['FREZ']['per_item'], 3.6)
        self.assertAlmostEqual(values['MANUALNE']['per_item'], .2)
        self.assertEqual(len(result['details']), 2)

    def test_invalid_quantity(self):
        for quantity in (0, -1, 'dziesięć', 1.5, True):
            with self.subTest(quantity=quantity), self.assertRaises(ValueError):
                read_workbook(workbook(quantity=quantity))

    def test_unknown_process_blocks_results(self):
        with self.assertRaisesRegex(ValueError, 'Nieznana grupa'):
            read_workbook(workbook(label='INNY - opis'))

    def test_missing_formula_cache_blocks_results(self):
        with self.assertRaisesRegex(ValueError, 'M10'):
            read_workbook(workbook(time='=1/24'))

    def test_negative_time(self):
        with self.assertRaises(ValueError):
            read_workbook(workbook(time=-1))

    def test_multiple_sheets(self):
        data = workbook(extra=True)
        self.assertTrue(read_workbook(data)['selection_required'])
        self.assertEqual(read_workbook(data, 'Druga')['sheet'], 'Druga')

    def test_dash_variants(self):
        for dash in ('-', '–', '—'):
            self.assertEqual(read_workbook(workbook(label='FREZ '+dash+' opis'))['processes'][0]['per_item'], 3.6)


if __name__ == '__main__':
    unittest.main()
