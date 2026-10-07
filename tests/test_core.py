import os
import tempfile
import unittest
from datetime import date, timedelta

from openpyxl import Workbook

from core import (
    compute_status,
    import_excel_file,
    parse_date,
    prepare_record,
)


class CoreLogicTests(unittest.TestCase):
    def test_parse_date_accepts_day_only(self):
        today = date.today()
        parsed = parse_date("25")
        self.assertEqual(parsed.day, 25)
        self.assertEqual(parsed.month, today.month)
        self.assertEqual(parsed.year, today.year)

    def test_parse_date_accepts_full_date(self):
        self.assertEqual(parse_date("25/08/2026"), date(2026, 8, 25))

    def test_compute_status_for_delivered_and_late(self):
        delivered = prepare_record(
            {
                "COD": "A01",
                "DATA ENTRADA": "10/08/2026",
                "CLIENTES": "Clínica A",
                "PACIENTE": "Ana",
                "SUP/INF": "Superior",
                "SAIDA": "20/08/2026",
                "OBSERVAÇÕES": "",
                "ENTREGUE": True,
            }
        )
        self.assertEqual(compute_status(delivered), "Entregue")

        late = prepare_record(
            {
                "COD": "A02",
                "DATA ENTRADA": "10/08/2026",
                "CLIENTES": "Clínica B",
                "PACIENTE": "Beto",
                "SUP/INF": "Inferior",
                "SAIDA": (date.today() - timedelta(days=2)).strftime("%d/%m/%Y"),
                "OBSERVAÇÕES": "",
                "ENTREGUE": False,
            }
        )
        self.assertEqual(compute_status(late), "Atrasado")

        ok = prepare_record(
            {
                "COD": "A03",
                "DATA ENTRADA": "10/08/2026",
                "CLIENTES": "Clínica C",
                "PACIENTE": "Cida",
                "SUP/INF": "Ambas",
                "SAIDA": (date.today() + timedelta(days=5)).strftime("%d/%m/%Y"),
                "OBSERVAÇÕES": "",
                "ENTREGUE": False,
            }
        )
        self.assertEqual(compute_status(ok), "No prazo")

    def test_import_excel_detects_header_after_title_row(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            path = os.path.join(tmp_dir, "TRABALHOS.xlsx")
            wb = Workbook()
            ws = wb.active
            ws.title = "Planilha"
            ws.append(["TRABALHOS"])
            ws.append(["COD", "DATA ENTRADA", "CLIENTES", "PACIENTE", "SUP/INF", "SAIDA"])
            ws.append(["T-001", "12/08/2026", "Clinica X", "Paciente 1", "Superior", "20/08/2026"])
            wb.save(path)

            records = import_excel_file(path)
            self.assertEqual(len(records), 1)
            self.assertEqual(records[0]["COD"], "T-001")
            self.assertEqual(records[0]["CLIENTES"], "Clinica X")


if __name__ == "__main__":
    unittest.main()
