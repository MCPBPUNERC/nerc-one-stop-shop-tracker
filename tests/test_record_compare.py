import unittest
from nerc_tracker import compare_records, severity

def record(key="abc", row=10, standard="PRC-005", status="Inactive", tracked=True):
    return {
        "key":key,
        "identity_text":f"{standard}|standard={standard}|requirement=r1",
        "sheet":"One Stop Shop",
        "source_row":row,
        "standards":[standard],
        "standard":standard,
        "family":standard.split("-")[0],
        "tracked":tracked,
        "fields":{"Standard":standard,"Requirement":"R1","Status":status},
    }

class SemanticComparisonTests(unittest.TestCase):
    def test_row_movement_is_not_a_change(self):
        old = record(row=100)
        new = record(row=225)
        self.assertEqual(compare_records([old],[new]), [])

    def test_status_change_is_detected(self):
        old = record(status="Inactive")
        new = record(status="Mandatory Subject to Enforcement")
        changes = compare_records([old],[new])
        self.assertEqual(len(changes),1)
        self.assertEqual(changes[0]["field"],"Status")
        self.assertEqual(changes[0]["severity"],"high")

    def test_multiple_row_reorder_is_not_change(self):
        old = [
            record(key="a", row=100, standard="PRC-005", status="Inactive"),
            record(key="b", row=101, standard="FAC-008", status="Inactive"),
            record(key="c", row=102, standard="TOP-001", status="Inactive"),
        ]
        new = [
            record(key="c", row=20, standard="TOP-001", status="Inactive"),
            record(key="a", row=250, standard="PRC-005", status="Inactive"),
            record(key="b", row=88, standard="FAC-008", status="Inactive"),
        ]
        self.assertEqual(compare_records(old, new), [])

    def test_real_change_survives_row_reorder(self):
        old = record(key="a", row=100, standard="FAC-008", status="Inactive")
        new = record(key="a", row=250, standard="FAC-008",
                     status="Mandatory Subject to Enforcement")
        changes = compare_records([old], [new])
        self.assertEqual(len(changes), 1)
        self.assertEqual(changes[0]["standard"], "FAC-008")
        self.assertEqual(changes[0]["old"], "Inactive")
        self.assertEqual(changes[0]["new"], "Mandatory Subject to Enforcement")

    def test_untracked_change_is_information_only(self):
        item = {"tracked":False,"field":"Status","old":"Inactive","new":"Mandatory Subject to Enforcement"}
        self.assertEqual(severity(item),"info")

if __name__ == "__main__":
    unittest.main()
