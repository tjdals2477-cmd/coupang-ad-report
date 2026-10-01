"""실행: py -m unittest discover -s tests -v   (coupang_report 폴더에서)"""
import io
import os
import sys
import unittest

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.analysis import Settings, best_campaign, build_context, delta, issue_campaign  # noqa: E402
from core.data import (ReportFileError, campaign_table, clean_product_name, guess_meta,  # noqa: E402
                       load_report, product_table, summarize)
from core.render import extract_report_data, merge_reports, render_report  # noqa: E402

COLS = ["캠페인 ID", "캠페인명", "광고집행 상품명", "광고집행 옵션ID", "노출수", "클릭수", "광고비",
        "총 주문수(1일)", "총 전환매출액(1일)", "총 주문수(14일)", "총 전환매출액(14일)"]


def make_df(rows):
    """rows: (캠페인ID, 캠페인명, 상품명, 옵션ID, 클릭, 광고비, 주문14, 매출14)"""
    data = [[cid, name, prod, opt, clicks * 10, clicks, cost, orders, rev, orders, rev]
            for cid, name, prod, opt, clicks, cost, orders, rev in rows]
    buf = io.BytesIO()
    pd.DataFrame(data, columns=COLS).to_excel(buf, index=False)
    buf.seek(0)
    return load_report(buf, "x.xlsx")


PREV = make_df([
    (1, "AI스마트광고", "상품A", 11, 100, 100_000, 5, 600_000),
    (2, "수동캠페인", "상품B", 22, 50, 80_000, 2, 160_000),
    (3, "옛이름", "상품C", 33, 40, 60_000, 4, 600_000),
])
CUR = make_df([
    (1, "AI스마트광고", "상품A", 11, 100, 150_000, 6, 750_000),
    (1, "AI스마트광고", "상품D", 44, 20, 20_000, 2, 300_000),
    (3, "새이름", "상품C", 33, 30, 60_000, 3, 480_000),
    (9, "신규캠페인", "상품E", 55, 50, 300_000, 1, 100_000),   # 광고비 1위지만 직전월 없음
    (8, "소액캠페인", "상품F", 66, 2, 1_000, 1, 90_000),       # ROAS 9000% 이지만 소액
    (1, "AI스마트광고", "-", 0, 5, 5_000, 0, 0),               # 배너 행
])


class DataTests(unittest.TestCase):
    def test_summarize_totals_and_rates(self):
        s = summarize(PREV)
        self.assertEqual(s["cost"], 240_000)
        self.assertEqual(s["revenue"], 1_360_000)
        self.assertAlmostEqual(s["roas"], 1_360_000 / 240_000 * 100)
        self.assertAlmostEqual(s["cvr"], 11 / 190 * 100)

    def test_zero_division_is_zero(self):
        df = make_df([(1, "a", "p", 1, 0, 0, 0, 0)])
        s = summarize(df)
        self.assertEqual((s["roas"], s["cvr"]), (0.0, 0.0))
        self.assertEqual(campaign_table(df).loc[0, "roas"], 0)

    def test_campaign_table_groups_by_id_sorted_by_cost(self):
        t = campaign_table(CUR)
        self.assertEqual(list(t["id"]), ["9", "1", "3", "8"])
        self.assertEqual(t.set_index("id").loc["1", "cost"], 175_000)

    def test_product_table_excludes_banner_rows(self):
        p = product_table(CUR, "1")
        self.assertEqual(sorted(p["option_id"]), ["11", "44"])

    def test_missing_columns_raise_clear_error(self):
        buf = io.BytesIO()
        pd.DataFrame({"캠페인명": ["a"]}).to_excel(buf, index=False)
        buf.seek(0)
        with self.assertRaises(ReportFileError) as cm:
            load_report(buf, "bad.xlsx")
        self.assertIn("광고비", str(cm.exception))

    def test_comma_strings_become_numbers_csv(self):
        csv = ",".join(COLS) + "\n1,a,p,1,\"1,000\",10,\"5,000\",1,\"20,000\",1,\"20,000\"\n"
        df = load_report(io.BytesIO(csv.encode("utf-8-sig")), "x.csv")
        self.assertEqual(summarize(df)["cost"], 5_000)
        self.assertEqual(summarize(df)["revenue"], 20_000)

    def test_clean_product_name(self):
        n = '안트 LED 일자등 50W,{"색상":"화이트"},5개,5개,주광색,주광색'
        self.assertEqual(clean_product_name(n), "안트 LED 일자등 50W · 5개 · 주광색")

    def test_guess_meta(self):
        m = guess_meta("안트_9월 광고보고서.xlsx")
        self.assertEqual((m["account"], m["month"], m["range"]), ("안트", 9, None))
        self.assertEqual(guess_meta("report.xlsx")["month"], None)

    def test_guess_meta_date_range(self):
        for name, label, days in [("9월16~9월22일.xlsx", "9/16~22", 7), ("9월 23~9월 29.xlsx", "9/23~29", 7),
                                  ("고네오_9.1~9.14 광고.xlsx", "9/1~14", 14), ("8월25일-9월7일.xlsx", "8/25~9/7", 14),
                                  ("12월29~1월4일.xlsx", "12/29~1/4", 7)]:
            m = guess_meta(name)
            self.assertEqual((m["label"], m["days"]), (label, days), name)


class AnalysisTests(unittest.TestCase):
    def test_delta_directions(self):
        self.assertEqual(delta(120, 100)["dir"], "up")
        self.assertEqual(delta(80, 100)["text"], "▼ 20.0%")
        self.assertEqual(delta(5.5, 5.0, "pp")["text"], "▲ 0.5%p")
        self.assertIsNone(delta(10, 0))

    def test_issue_requires_campaign_in_previous_month(self):
        prev_c, cur_c = campaign_table(PREV), campaign_table(CUR)
        issue = issue_campaign(prev_c, cur_c, summarize(CUR)["cost"])
        self.assertEqual(issue["cur"]["id"], "1")  # 9번(신규)은 금월 1위지만 제외

    def test_issue_none_when_all_new(self):
        cur = make_df([(77, "신규", "p", 1, 10, 10_000, 1, 50_000)])
        self.assertIsNone(issue_campaign(campaign_table(PREV), campaign_table(cur), 10_000))

    def test_issue_detects_rename_by_id(self):
        cur = make_df([(3, "새이름", "상품C", 33, 30, 60_000, 3, 480_000)])
        issue = issue_campaign(campaign_table(PREV), campaign_table(cur), 60_000)
        self.assertTrue(issue["renamed"])
        self.assertEqual(issue["prev"]["name"], "옛이름")

    def test_best_same_campaign_and_min_cost_filter(self):
        best = best_campaign(campaign_table(PREV), campaign_table(CUR), 50_000)
        self.assertEqual(best["cur"]["id"], "3")      # 소액캠페인(8) 제외, ROAS 800%
        self.assertEqual(best["mode"], "same")
        self.assertEqual(best["prev"]["name"], "옛이름")

    def test_best_fallback_to_prev_best(self):
        cur = make_df([(9, "신규", "p", 1, 10, 100_000, 3, 900_000)])
        best = best_campaign(campaign_table(PREV), campaign_table(cur), 50_000)
        self.assertEqual(best["mode"], "fallback")
        self.assertEqual(best["prev"]["id"], "3")      # 직전월 ROAS 1위 (1000%)

    def test_top_products_sorting_and_thresholds(self):
        s = Settings(account="t", min_clicks=25, min_product_cost=10_000)
        ctx = build_context(PREV, CUR, s)
        ai = next(c for c in ctx["campaigns"] if c["id"] == "1")
        self.assertEqual([r["name"] for r in ai["top"]["roas"]], ["상품D", "상품A"])  # 1500% > 500%
        self.assertEqual([r["name"] for r in ai["top"]["cvr"]], ["상품A"])          # 클릭 20 < 25 제외

    def test_campaign_cards_targets_and_new_flag(self):
        s = Settings(account="t", targets={"1": {"target_roas": 500, "daily_budget": 30000}})
        cards = {c["id"]: c for c in build_context(PREV, CUR, s)["campaigns"]}
        self.assertTrue(cards["1"]["hit"])
        self.assertEqual(cards["1"]["budget"], "30,000원")
        self.assertIsNone(cards["3"]["achieve"])
        self.assertTrue(cards["9"]["is_new"])
        self.assertFalse(cards["1"]["is_new"])

    def test_short_campaign_name(self):
        from core.analysis import short
        self.assertEqual(short("[리뷰 점수 복구 필요]다목적 주차장등 60W"), "다목적 주차장등 60W")
        self.assertEqual(short("(09.16) LED 사각 천장등"), "LED 사각 천장등")
        self.assertEqual(short("안트브랜드_시스템시리즈_그로스전용제품"), "안트브랜드 시스템시리즈 그로스전용제품")

    def test_trend_chart_requires_all_three_values(self):
        self.assertIsNone(build_context(PREV, CUR, Settings(account="t"))["trend"])
        self.assertIsNone(build_context(PREV, CUR, Settings(account="t", prev_total_sales=1_000_000))["trend"])
        t = build_context(PREV, CUR, Settings(account="t", prev_total_sales=1_000_000,
                                              cur_total_sales=2_000_000))["trend"]
        self.assertEqual(t["granularity"], "month")
        self.assertEqual(t["unit"], "기간 합계")
        self.assertEqual(t["n"], 2)
        self.assertEqual([sr["key"] for sr in t["series"]], ["sales", "revenue", "cost"])
        self.assertEqual(t["series"][0]["values"], [1_000_000, 2_000_000])
        ys = sorted(sr["end_label_y"] for sr in t["series"])  # 끝점 라벨이 서로 겹치지 않음
        self.assertTrue(all(b - a >= 15.9 for a, b in zip(ys, ys[1:])), ys)

    def test_short_won(self):
        from core.analysis import short_won
        self.assertEqual(short_won(90_989_740), "9,099만")
        self.assertEqual(short_won(120_000_000), "1.2억")
        self.assertEqual(short_won(100_000_000), "1억")

    def test_total_sales_missing_shows_pending(self):
        ctx = build_context(PREV, CUR, Settings(account="t"))
        sales = ctx["kpis"][1]
        self.assertIsNone(sales["prev"])
        self.assertIsNone(sales["delta"])

    def test_three_summary_steps_are_short(self):
        steps = build_context(PREV, CUR, Settings(account="t"))["steps"]
        self.assertEqual(len(steps), 3)
        for s in steps:
            lines = s["text"].splitlines()
            self.assertTrue(1 <= len(lines) <= 4, s["text"])
            self.assertTrue(all(len(l) <= 40 for l in lines), s["text"])


class RenderTests(unittest.TestCase):
    def setUp(self):
        self.html = render_report(build_context(PREV, CUR, Settings(account="안트", prev_label="8월", cur_label="9월")))

    def test_report_has_all_sections_and_toggle(self):
        for marker in ["① 8월 vs 9월", "② 9월 최대 이슈", "③ 9월 최고 효율", "④ 개선점", "⑤ 9월 캠페인",
                       'data-sort="roas"', 'data-sort="cvr"']:
            self.assertIn(marker, self.html)

    def test_embedded_data_roundtrip(self):
        d = extract_report_data(self.html)
        self.assertEqual(d["account"], "안트")
        self.assertEqual(d["cur"]["cost"], summarize(CUR)["cost"])

    def test_user_text_is_escaped(self):
        ctx = build_context(PREV, CUR, Settings(account="<b>x</b>"))
        ctx["steps"][0]["text"] = "<script>alert(1)</script>"
        html = render_report(ctx)
        self.assertNotIn("<script>alert(1)</script>", html)

    def test_summary_deltas_are_colored(self):
        self.assertIn('<span class="d-up">▲', self.html)
        self.assertIn('<span class="d-down">▼', self.html)
        self.assertIn('class="step-conc"', self.html)

    def test_colorize_escapes_user_text(self):
        from core.render import colorize_deltas
        out = str(colorize_deltas("<b>x</b> ▲ 12.5% ▼3.0%p"))
        self.assertIn("&lt;b&gt;", out)
        self.assertIn('<span class="d-up">▲ 12.5%</span>', out)
        self.assertIn('<span class="d-down">▼ 3.0%p</span>', out)

    def test_merge_escapes_script_and_adds_summary(self):
        merged, skipped = merge_reports([{"label": "A", "html": self.html}, {"label": "B", "html": self.html}],
                                        title="통합")
        self.assertEqual(merged.count("</script>"), 2)  # 내장 보고서의 </script>가 이스케이프됨
        self.assertIn('data-report="summary"', merged)
        self.assertEqual(skipped, [])

    def test_merge_skips_summary_for_foreign_html(self):
        merged, skipped = merge_reports([{"label": "A", "html": self.html},
                                         {"label": "외부", "html": "<html><title>x</title></html>"}])
        self.assertEqual(skipped, ["외부"])
        self.assertNotIn('data-report="summary"', merged)

    def test_trend_chart_rendered_only_with_sales(self):
        self.assertNotIn('class="trend-card"', self.html)
        ctx = build_context(PREV, CUR, Settings(account="t", prev_total_sales=1_000_000, cur_total_sales=2_000_000))
        html = render_report(ctx)
        self.assertIn('class="trend-card"', html)
        self.assertEqual(html.count('class="hit"'), 2)

    def test_report_has_editor(self):
        self.assertIn("data-edit-kit", self.html)
        self.assertIn("coupang-report:update", self.html)

    def test_merged_shell_receives_edits(self):
        merged, _ = merge_reports([{"label": "A", "html": self.html}], add_summary=False)
        self.assertIn("data-shell-save", merged)
        self.assertIn("coupang-shell:hello", merged)

    def test_merge_requires_reports(self):
        with self.assertRaises(ValueError):
            merge_reports([])


def make_daily_df(start, days, seed=1):
    """가상 일별 보고서 (테스트용): 하루 1행, 광고비·매출이 날마다 다름."""
    import random
    rnd = random.Random(seed)
    rows = []
    for i in range(days):
        d = pd.Timestamp(start) + pd.Timedelta(days=i)
        cost = rnd.randint(50, 150) * 1000
        rows.append([1, "AI스마트광고", "상품A", 11, 1000, 50, cost, 3, cost * 5, 3, cost * 5, d.strftime("%Y%m%d")])
    buf = io.BytesIO()
    pd.DataFrame(rows, columns=COLS + ["날짜"]).to_excel(buf, index=False)
    buf.seek(0)
    return load_report(buf, "daily.xlsx")


def fake_sales(df_prev, df_cur, gran):
    """가상 총매출: 구간마다 광고 전환 매출의 1.4배."""
    from core.analysis import make_buckets
    out = {}
    for side, df in (("prev", df_prev), ("cur", df_cur)):
        for b in make_buckets(df, side, gran):
            out[b["key"]] = round(b["revenue"] * 1.4)
    return out


class DailyTrendTests(unittest.TestCase):
    def test_dates_detected_from_yyyymmdd(self):
        from core.data import date_span
        df = make_daily_df("2026-08-01", 31)
        self.assertEqual(date_span(df), (pd.Timestamp("2026-08-01"), pd.Timestamp("2026-08-31")))
        self.assertIsNone(date_span(PREV))  # 월 합계 파일은 날짜 없음

    def test_month_files_use_7_day_buckets(self):
        from core.analysis import choose_granularity, make_buckets
        prev, cur = make_daily_df("2026-08-01", 31), make_daily_df("2026-09-01", 30, seed=2)
        self.assertEqual(choose_granularity(prev, cur), "week")
        b = make_buckets(prev, "prev", "week")
        self.assertEqual([x["label"] for x in b], ["8/1~7", "8/8~14", "8/15~21", "8/22~28", "8/29~31"])
        self.assertEqual(sum(x["cost"] for x in b), summarize(prev)["cost"])  # 구간 합 = 월 합계

    def test_7_and_14_day_files_use_daily_points(self):
        from core.analysis import choose_granularity
        self.assertEqual(choose_granularity(make_daily_df("2026-09-01", 7), make_daily_df("2026-09-08", 7)), "day")
        self.assertEqual(choose_granularity(make_daily_df("2026-08-18", 14), make_daily_df("2026-09-01", 14)), "day")
        # 한쪽이라도 14일을 넘으면 둘 다 7일 단위 (같은 단위로 비교)
        self.assertEqual(choose_granularity(make_daily_df("2026-08-01", 31), make_daily_df("2026-09-01", 14)), "week")
        # 한쪽에 날짜가 없으면 월 합계 2점
        self.assertIsNone(choose_granularity(PREV, make_daily_df("2026-09-01", 14)))

    def test_missing_day_filled_with_zero(self):
        from core.analysis import make_buckets
        df = make_daily_df("2026-09-01", 7)
        df = df[df["date"] != pd.Timestamp("2026-09-03")]
        b = make_buckets(df, "cur", "day")
        self.assertEqual(len(b), 7)
        self.assertEqual(b[2]["cost"], 0)

    def test_daily_chart_with_fake_sales(self):
        prev, cur = make_daily_df("2026-09-01", 7), make_daily_df("2026-09-08", 7, seed=3)
        s = Settings(account="t", prev_label="1주차", cur_label="2주차",
                     bucket_sales=fake_sales(prev, cur, "day"))
        t = build_context(prev, cur, s)["trend"]
        self.assertEqual((t["granularity"], t["n"]), ("day", 14))
        self.assertEqual(len(t["dividers"]), 1)                     # 주차 사이 구분선
        self.assertEqual([h["label"] for h in t["headers"]], ["1주차", "2주차"])
        self.assertEqual(len(t["series"][0]["segs"]), 2)             # 기간 사이는 선을 끊음
        html = render_report(build_context(prev, cur, s))
        self.assertEqual(html.count('class="hit"'), 14)

    def test_weekly_chart_and_blank_cell_blocks_chart(self):
        prev, cur = make_daily_df("2026-08-01", 31), make_daily_df("2026-09-01", 30, seed=2)
        sales = fake_sales(prev, cur, "week")
        t = build_context(prev, cur, Settings(account="t", bucket_sales=sales))["trend"]
        self.assertEqual((t["granularity"], t["n"]), ("week", 10))
        self.assertEqual(t["cols"][4]["label"], "8/29~31 (3일)")      # 7일 미만 구간 표시
        self.assertTrue(t["cols"][4]["partial"] and t["has_partial"])
        self.assertFalse(t["cols"][0]["partial"])
        sales.pop(next(iter(sales)))                                  # 한 칸 비움
        self.assertIsNone(build_context(prev, cur, Settings(account="t", bucket_sales=sales))["trend"])

    def test_x_labels_thinned_without_collision(self):
        prev, cur = make_daily_df("2026-09-01", 14), make_daily_df("2026-09-15", 14, seed=3)
        t = build_context(prev, cur, Settings(account="t", bucket_sales=fake_sales(prev, cur, "day")))["trend"]
        shown = [i for i, c in enumerate(t["cols"]) if c["show"]]
        self.assertIn(0, shown); self.assertIn(13, shown); self.assertIn(14, shown); self.assertIn(27, shown)
        same_group_gaps = [b - a for a, b in zip(shown, shown[1:]) if not (a <= 13 < b)]
        self.assertTrue(all(g >= 2 for g in same_group_gaps), shown)

    def test_tooltip_delta_vs_previous_bucket(self):
        prev, cur = make_daily_df("2026-09-01", 7), make_daily_df("2026-09-08", 7, seed=3)
        t = build_context(prev, cur, Settings(account="t", bucket_sales=fake_sales(prev, cur, "day")))["trend"]
        self.assertIsNone(t["tooltip"][0]["deltas"][0])   # 첫 구간
        self.assertIsNotNone(t["tooltip"][1]["deltas"][0])
        self.assertIsNone(t["tooltip"][7]["deltas"][0])   # 금월 첫 구간은 직전월과 이어지지 않음


# 실데이터 회귀 테스트 (로컬 전용): 계정 수치는 저장소에 올리지 않도록 별도 파일에 둔다.
# tests/real_data_local.json 예시:
# {"prev": "C:/.../브랜드_8월.xlsx", "cur": "C:/.../브랜드_9월.xlsx",
#  "prev_revenue": 0, "prev_cvr": 0.0, "issue": "캠페인명"}
_LOCAL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "real_data_local.json")
REAL = {}
if os.path.exists(_LOCAL):
    import json
    with open(_LOCAL, encoding="utf-8") as f:
        REAL = json.load(f)


@unittest.skipUnless(REAL and os.path.exists(REAL.get("prev", "")) and os.path.exists(REAL.get("cur", "")),
                     "실데이터 설정 없음 (tests/real_data_local.json)")
class RealDataRegression(unittest.TestCase):
    """기존 보고서 수치와 대조."""

    def test_previous_month_matches_existing_report(self):
        s = summarize(load_report(REAL["prev"], os.path.basename(REAL["prev"])))
        self.assertEqual(round(s["revenue"]), REAL["prev_revenue"])
        self.assertAlmostEqual(s["cvr"], REAL["prev_cvr"], places=2)

    def test_sections(self):
        prev = load_report(REAL["prev"], os.path.basename(REAL["prev"]))
        cur = load_report(REAL["cur"], os.path.basename(REAL["cur"]))
        ctx = build_context(prev, cur, Settings(account="t"))
        self.assertEqual(ctx["issue"]["cur"]["name"], REAL["issue"])
        self.assertEqual(ctx["best"]["mode"], "same")


if __name__ == "__main__":
    unittest.main()
