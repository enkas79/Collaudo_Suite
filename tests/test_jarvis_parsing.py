import unittest

from collaudo_suite.checklist.jarvis_api import COMMERCIAL_CODE_FIELD, _ticket_codes, _ticket_number


class JarvisParsingTests(unittest.TestCase):
    def test_commercial_code_read_from_omnisearch_field(self):
        ticket = {
            "stringProperties": [
                {"key": "propertydefinition_1288/jarvisformfield_3168", "value": ["DatasetElement_706033"]},
                {"key": COMMERCIAL_CODE_FIELD.upper(), "value": ["NC30000068", ""]},
            ]
        }
        self.assertEqual(_ticket_codes(ticket), ["NC30000068"])

    def test_opaque_property_values_are_not_used_as_codes(self):
        # Regressione: i riferimenti tipo DatasetElement_xxx non devono diventare codici commerciali.
        ticket = {"properties": {"propertydefinition_999": "DatasetElement_706033"}}
        self.assertEqual(_ticket_codes(ticket), [])

    def test_named_property_still_supported(self):
        ticket = {"properties": {"Commessa S Codice commerciale": "GENYA600"}}
        self.assertEqual(_ticket_codes(ticket), ["GENYA600"])

    def test_ticket_number_normalization(self):
        self.assertEqual(_ticket_number({"id": "Ticket_12345"}), "12345")
        self.assertEqual(_ticket_number({"number": 678.0}), "678")
        self.assertEqual(_ticket_number({"id": ""}), "")


if __name__ == "__main__":
    unittest.main()
