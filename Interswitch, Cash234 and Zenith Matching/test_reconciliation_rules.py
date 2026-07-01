import unittest

import pandas as pd

from reconciliation import (
    build_interswitch_zenith_description_links,
    build_interswitch_zenith_carryover_sheets,
    build_presence_table,
    build_required_report_sheets,
    reconcile,
)
from tran_description_matcher import build_tran_description_comparison


class ReconciliationRuleTests(unittest.TestCase):
    def test_main_reconciliation_uses_rrn_only_for_updated_export(self):
        raw_frames = {
            "INTERSWITCH": pd.DataFrame(
                [
                    {
                        "Retrieval_Reference_Nr": "RRN001",
                        "Local_Date_Time": "2026-06-01",
                        "Amount": "100",
                        "Tran_ID": "123456789012|123456",
                    }
                ]
            ),
            "CASH234": pd.DataFrame(
                [
                    {
                        "R R N": "RRN001",
                        "Transaction Date": "2026-06-01",
                        "Amount": "100",
                    }
                ]
            ),
            "ZENITH": pd.DataFrame(
                [
                    {
                        "RRN": "",
                        "EffectiveDate": "01/06/2026",
                        "Amount": "100",
                        "Description": "Transfer|ZIB|123456789012|123456",
                    },
                    {
                        "RRN": "",
                        "EffectiveDate": "01/06/2026",
                        "Amount": "100",
                        "Description": "Transfer|ZIB|999999999999|999999",
                    },
                ]
            ),
        }

        results = reconcile(raw_frames)

        updated_zenith = results["updated_exports"]["ZENITH"]
        self.assertEqual(len(updated_zenith), 2)
        self.assertEqual(len(results["sets"]["INTERSWITCH_AND_CASH234_ONLY"]), 1)
        self.assertEqual(len(results["sets"]["MATCHED_IN_ALL_THREE"]), 0)

    def test_zenith_unsettled_report_uses_rrn_only(self):
        raw_frames = {
            "INTERSWITCH": pd.DataFrame(
                [
                    {
                        "Retrieval_Reference_Nr": "910400365506",
                        "Local_Date_Time": "2026-06-01",
                        "Amount": "1400",
                        "Tran_ID": "CR|FBN|ZIB|010126000831|539276",
                    }
                ]
            ),
            "CASH234": pd.DataFrame(
                columns=["R R N", "Transaction Date", "Amount"]
            ),
            "ZENITH": pd.DataFrame(
                [
                    {
                        "RRN": "",
                        "EffectiveDate": "01/06/2026",
                        "Amount": "1400",
                        "Description": "CR|FBN|ZIB|010126000831|539276",
                    }
                ]
            ),
        }

        results = reconcile(raw_frames)
        sheets = build_required_report_sheets(results["report_frames"], "Zenith_Unsettled.xlsx")

        self.assertEqual(len(sheets["Zenith Unsettled"]), 1)

    def test_separate_tran_description_matcher_finds_token_match(self):
        result = build_tran_description_comparison(
            pd.DataFrame(
                [
                    {
                        "Retrieval_Reference_Nr": "910400365506",
                        "Local_Date_Time": "2026-06-01",
                        "Amount": "1400",
                        "Tran_ID": "CR|ECO|ZIB|290526105743|141840",
                    }
                ]
            ),
            pd.DataFrame(
                [
                    {
                        "RRN": "",
                        "EffectiveDate": "01/06/2026",
                        "Amount": "1400",
                        "Description": "ISW TRF-CR|ECO|ZIB|290526105743|141840-2026-05-29",
                    }
                ]
            ),
        )

        self.assertEqual(len(result["Matched"]), 1)
        self.assertEqual(result["Matched"].iloc[0]["MATCH_TOKEN"], "290526105743|141840")

    def test_separate_tran_description_matcher_handles_slants_in_zenith_description(self):
        result = build_tran_description_comparison(
            pd.DataFrame(
                [
                    {
                        "Retrieval_Reference_Nr": "910400365506",
                        "Local_Date_Time": "2026-06-01",
                        "Amount": "1400",
                        "Tran_ID": "CR|FBN|ZIB|010126000831|539276",
                    }
                ]
            ),
            pd.DataFrame(
                [
                    {
                        "RRN": "",
                        "EffectiveDate": "01/06/2026",
                        "Amount": "1400",
                        "Description": "ISW TRF/CR/FBN/ZIB/010126000831/539276-2026-05-29",
                    }
                ]
            ),
        )

        self.assertEqual(len(result["Matched"]), 1)
        self.assertEqual(result["Matched"].iloc[0]["MATCH_TOKEN"], "010126000831|539276")

    def test_interswitch_zenith_carryover_contains_rrn_remainders_for_second_matcher(self):
        raw_frames = {
            "INTERSWITCH": pd.DataFrame(
                [
                    {
                        "Retrieval_Reference_Nr": "RRNMATCHED",
                        "Local_Date_Time": "2026-06-01",
                        "Amount": "100",
                        "Tran_ID": "CR|FBN|ZIB|010126000831|539276",
                    },
                    {
                        "Retrieval_Reference_Nr": "RRNONLYI",
                        "Local_Date_Time": "2026-06-01",
                        "Amount": "200",
                        "Tran_ID": "CR|ECO|ZIB|290526105743|141840",
                    },
                ]
            ),
            "CASH234": pd.DataFrame(columns=["R R N", "Transaction Date", "Amount"]),
            "ZENITH": pd.DataFrame(
                [
                    {
                        "RRN": "RRNMATCHED",
                        "EffectiveDate": "01/06/2026",
                        "Amount": "100",
                        "Description": "RRN matched row",
                    },
                    {
                        "RRN": "",
                        "EffectiveDate": "01/06/2026",
                        "Amount": "200",
                        "Description": "ISW TRF-CR|ECO|ZIB|290526105743|141840-2026-05-29",
                    },
                ]
            ),
        }

        results = reconcile(raw_frames)
        carryover = build_interswitch_zenith_carryover_sheets(results["report_frames"])

        self.assertEqual(carryover["Interswitch Carryover"]["Retrieval_Reference_Nr"].tolist(), ["RRNONLYI"])
        self.assertEqual(len(carryover["Zenith Carryover"]), 1)
        self.assertEqual(
            carryover["Zenith Carryover"].iloc[0]["Description"],
            "ISW TRF-CR|ECO|ZIB|290526105743|141840-2026-05-29",
        )

    def test_blank_zenith_rrns_are_not_counted_as_unique_rrns(self):
        presence = build_presence_table(
            {
                "INTERSWITCH": pd.DataFrame({"NORMALIZED_RRN": pd.Series(["ABC123"], dtype="string")}),
                "CASH234": pd.DataFrame({"NORMALIZED_RRN": pd.Series([], dtype="string")}),
                "ZENITH": pd.DataFrame({"NORMALIZED_RRN": pd.Series([""], dtype="string")}),
            }
        )

        self.assertEqual(presence["NORMALIZED_RRN"].tolist(), ["ABC123"])

    def test_description_links_use_normalized_indexed_tokens(self):
        primary_frame = pd.DataFrame(
            {
                "INTERSWITCH_Tran_ID": [" 123456789012 / 123456 "],
                "NORMALIZED_RRN": ["RRN001"],
            },
            index=[10],
        )
        zenith_frame = pd.DataFrame(
            {
                "ZENITH_Description": ["Transfer | ZIB | 123456789012 | 123456"],
                "NORMALIZED_RRN": [""],
            },
            index=[20],
        )

        links = build_interswitch_zenith_description_links(
            primary_frame=primary_frame,
            zenith_frame=zenith_frame,
            already_settled_mask=pd.Series(False, index=primary_frame.index),
            already_matched_zenith_mask=pd.Series(False, index=zenith_frame.index),
        )

        self.assertEqual(links, [(10, 20)])

    def test_description_links_do_not_match_partial_tran_id_tokens(self):
        primary_frame = pd.DataFrame(
            {
                "INTERSWITCH_Tran_ID": ["CR|FBN|ZIB|010126000831|53927699"],
                "NORMALIZED_RRN": ["RRN001"],
            },
            index=[10],
        )
        zenith_frame = pd.DataFrame(
            {
                "ZENITH_Description": ["CR|FBN|ZIB|010126000831|539276"],
                "NORMALIZED_RRN": [""],
            },
            index=[20],
        )

        links = build_interswitch_zenith_description_links(
            primary_frame=primary_frame,
            zenith_frame=zenith_frame,
            already_settled_mask=pd.Series(False, index=primary_frame.index),
            already_matched_zenith_mask=pd.Series(False, index=zenith_frame.index),
        )

        self.assertEqual(links, [])

    def test_description_links_are_one_to_one_after_rrn_matching(self):
        primary_frame = pd.DataFrame(
            {
                "INTERSWITCH_Tran_ID": [
                    "CR|FBN|ZIB|010126000831|539276",
                    "CR|FBN|ZIB|010126000831|539276",
                ],
                "NORMALIZED_RRN": ["RRN001", "RRN002"],
            },
            index=[10, 11],
        )
        zenith_frame = pd.DataFrame(
            {
                "ZENITH_Description": ["CR|FBN|ZIB|010126000831|539276"],
                "NORMALIZED_RRN": [""],
            },
            index=[20],
        )

        links = build_interswitch_zenith_description_links(
            primary_frame=primary_frame,
            zenith_frame=zenith_frame,
            already_settled_mask=pd.Series(False, index=primary_frame.index),
            already_matched_zenith_mask=pd.Series(False, index=zenith_frame.index),
        )

        self.assertEqual(links, [(10, 20)])


if __name__ == "__main__":
    unittest.main()
