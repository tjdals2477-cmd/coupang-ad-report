"""광고 보고서 페이지: 광고보고서 xlsm(raw 시트)·상품 전환 보고서 형식 읽기."""
import io
import os
import sys
import unittest

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.data import guess_meta, load_report  # noqa: E402


def book(sheets: dict) -> io.BytesIO:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for r in rows:
            ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


class AdFormatTest(unittest.TestCase):
    def test_xlsm_report_uses_raw_sheet_and_aliases(self):
        f = book({
            "일자별": [["날짜", "노출", "광고비용"], ["(비어 있음)", 1, 2]],          # 요약 시트가 먼저
            "raw": [["날짜", "캠페인 ID", "캠페인", "광고집행 상품명", "광고집행 옵션ID", "노출수", "클릭수", "광고비",
                     "총 주문수(14일)", "총 전환매출액(14일)"],
                    [None, 1, "AI", "상품A", 11, 100, 10, 5000, 2, 30000],
                    [None, 1, "AI", "상품B", 12, 50, 5, 1000, 0, 0]],
        })
        df = load_report(f, "안트_9월 광고 보고서.xlsm")
        self.assertEqual(df["cost"].sum(), 6000)
        self.assertEqual(df["revenue"].sum(), 30000)
        self.assertTrue(df["date"].isna().all())  # 날짜 없는 기간 합계 파일

    def test_conversion_report_headers_without_campaign_id(self):
        f = book({"Sheet1": [["캠페인 이름", "광고집행 상품명", "광고 집행 옵션 ID", "노출수", "클릭수", "광고비(원)",
                              "총 주문수 (14일)", "총 전환 매출액 (14일)(원)"],
                             ["AI스마트광고", "상품A", 11, 100, 10, 5000, 2, 30000]]})
        df = load_report(f, "안트_AI_상품전환_0901-0909.xlsx")
        self.assertEqual(df["캠페인 ID"].iloc[0], "AI스마트광고")
        self.assertEqual(df["orders"].sum(), 2)

    def test_guess_meta_xlsm(self):
        self.assertEqual(guess_meta("안트_9월 광고 보고서.xlsm")["month"], 9)


if __name__ == "__main__":
    unittest.main()
