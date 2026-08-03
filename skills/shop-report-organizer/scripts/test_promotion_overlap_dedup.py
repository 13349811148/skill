"""Regression checks for overlapping near-30-day promotion downloads."""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path
from unittest.mock import patch


SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import organize_reports as organizer


PDD_HEADERS = [
    "归档日期",
    "归档商品ID",
    "归档推广类型",
    "总花费(元)",
    "点击量",
]
PDD_SIGNATURE = organizer.Signature(
    report_type="推广数据",
    target_folder="推广数据",
    platform="拼多多",
    required_headers=[],
    product_id_headers=["归档商品ID"],
    period_headers=["归档日期"],
    period_mode="column_range",
    promotion_type="拼多多-商品推广",
    notes="test",
)
PDD_DAILY_HEADERS = [
    "日期",
    "曝光量",
    "点击量",
    "总花费(元)",
    "净实际投产比",
]
PDD_DAILY_SIGNATURE = organizer.Signature(
    report_type="推广数据",
    target_folder="推广数据",
    platform="拼多多",
    required_headers=PDD_DAILY_HEADERS,
    product_id_headers=[],
    period_headers=["日期"],
    period_mode="column_range",
    promotion_type="拼多多-商品推广",
    notes="test daily breakdown",
)


def promotion_row(row_number: int, spend: str, clicks: str) -> organizer.PromotionRow:
    return organizer.PromotionRow(
        sheet_name="推广数据",
        row_number=row_number,
        headers=PDD_HEADERS,
        values=["2026-07-01", "1000001", "拼多多-商品推广", spend, clicks],
        signature=PDD_SIGNATURE,
    )


def test_today_overlap_replaces_yesterday_record() -> None:
    yesterday = Path("昨天下载的近30天推广.xlsx")
    today = Path("今天下载的近30天推广.xlsx")
    rows_by_path = {
        yesterday: [promotion_row(2, "10.00", "5")],
        today: [promotion_row(2, "10", "8")],
    }
    with (
        patch.object(organizer, "scan_promotion_file_full"),
        patch.object(
            organizer,
            "iter_promotion_rows_full",
            side_effect=lambda path, _max_cols: iter(rows_by_path[path]),
        ),
    ):
        _headers, records, duplicate_rows = organizer.promotion_records(
            [yesterday, today],
            120,
            preferred_paths={today},
        )

    assert duplicate_rows == 1
    assert len(records) == 1
    assert records[0]["点击量"] == "8"
    assert records[0]["总花费(元)"] == "10"


def test_different_spend_is_not_deduplicated() -> None:
    source = Path("同一来源.xlsx")
    rows_by_path = {
        source: [
            promotion_row(2, "10.00", "5"),
            promotion_row(3, "10.01", "6"),
        ]
    }
    with (
        patch.object(organizer, "scan_promotion_file_full"),
        patch.object(
            organizer,
            "iter_promotion_rows_full",
            side_effect=lambda path, _max_cols: iter(rows_by_path[path]),
        ),
    ):
        _headers, records, duplicate_rows = organizer.promotion_records(
            [source], 120
        )

    assert duplicate_rows == 0
    assert len(records) == 2


def test_promotion_type_is_part_of_the_safe_scope() -> None:
    pdd = {
        "归档日期": "2026-07-01",
        "归档商品ID": "1000001",
        "归档推广类型": "拼多多-商品推广",
        "总花费(元)": "10.00",
    }
    wanxiang = {
        "归档日期": "2026-07-01",
        "归档商品ID": "1000001",
        "归档推广类型": "万相台",
        "花费": "10",
    }
    assert organizer.promotion_record_key(pdd) != organizer.promotion_record_key(
        wanxiang
    )


def test_all_supported_spend_headers_are_normalized() -> None:
    records = [
        {
            "归档日期": "2026-07-01",
            "归档商品ID": "1000001",
            "归档推广类型": "万相台",
            "花费": "￥1,234.50",
        },
        {
            "归档日期": "2026-07-01",
            "归档商品ID": "1000001",
            "归档推广类型": "新客加速",
            "预估新客加速费用": "1234.500",
        },
        {
            "归档日期": "2026-07-01",
            "归档商品ID": "1000001",
            "归档推广类型": "新客加速",
            "预估抽佣金额": "1234.5元",
        },
        {
            "归档日期": "2026-07-01",
            "归档商品ID": "1000001",
            "归档推广类型": "老客加速",
            "预估老客加速费用": 1234.5,
        },
    ]
    assert {
        organizer.promotion_record_key(record)[-1] for record in records
    } == {"1234.5"}


def test_transactional_merge_keeps_current_source_once() -> None:
    headers = [
        "归档日期",
        "归档商品ID",
        "归档推广类型",
        "商品名称",
        "商品ID",
        "推广名称",
        "总花费(元)",
        "净实际投产比",
        "点击量",
    ]
    with tempfile.TemporaryDirectory(prefix="promotion-overlap-test-") as temp_name:
        root = Path(temp_name)
        target = root / "推广数据" / "测试店" / "测试店—推广数据.xlsx"
        source = root / "下载未分类" / "今天近30天推广.xlsx"
        target.parent.mkdir(parents=True)
        source.parent.mkdir(parents=True)
        organizer.write_xlsx_table(
            target,
            headers,
            [
                [
                    "2026-07-01",
                    "1000001",
                    "拼多多-商品推广",
                    "测试商品",
                    "1000001",
                    "商品推广",
                    "10.00",
                    "2.0",
                    "5",
                ]
            ],
            "推广数据",
        )
        organizer.write_xlsx_table(
            source,
            headers,
            [
                [
                    "2026-07-01",
                    "1000001",
                    "拼多多-商品推广",
                    "测试商品",
                    "1000001",
                    "商品推广",
                    "10",
                    "2.1",
                    "8",
                ]
            ],
            "推广数据",
        )
        analysis = organizer.Analysis(
            source=source,
            status="ready",
            reason="ok",
            report_type="推广数据",
            target_folder="推广数据",
            shop_name="测试店",
            shop_source="test",
            period="2026-06-02 至 2026-07-01",
            period_source="test",
            download_date="2026-07-02",
            target_path=target,
            product_counts="",
            matched_headers="",
            missing_fields="",
            map_updates=[],
            product_ids=["1000001"],
            new_product_ids=[],
        )

        final_path, source_count, row_count, duplicate_rows = (
            organizer.merge_promotion_data(analysis, 120)
        )
        _headers, records, remaining_duplicates = organizer.promotion_records(
            [final_path], 120
        )

        assert source_count == 2
        assert row_count == 1
        assert duplicate_rows == 1
        assert remaining_duplicates == 0
        assert records[0]["点击量"] == "8"
        assert records[0]["净实际投产比"] == "2.1"
        assert not source.exists()


def test_strict_daily_breakdown_uses_validated_filename_id() -> None:
    source = Path(
        "商品推广_分天数据_20260703至20260801(测试旗舰店)(测试商品_1000001).xls"
    )
    metadata = organizer.pdd_daily_breakdown_filename_metadata(source)
    assert metadata is not None
    assert metadata.shop_text == "测试旗舰店"
    assert metadata.filename_product_id == "1000001"

    rows = [PDD_DAILY_HEADERS, ["2026-08-01", "100", "5", "10.00", "2.0"]]
    assert organizer.promotion_preview_missing_fields(
        rows,
        0,
        PDD_DAILY_HEADERS,
        PDD_DAILY_SIGNATURE,
        fallback_product_id="1000001",
    ) == []

    promotion_rows = [
        organizer.PromotionRow(
            sheet_name="推广数据",
            row_number=2,
            headers=PDD_DAILY_HEADERS,
            values=rows[1],
            signature=PDD_DAILY_SIGNATURE,
        )
    ]
    with (
        patch.object(organizer, "scan_promotion_file_full"),
        patch.object(
            organizer,
            "iter_promotion_rows_full",
            side_effect=lambda path, _max_cols: iter(promotion_rows),
        ),
    ):
        _headers, records, duplicate_rows = organizer.promotion_records(
            [source],
            120,
            fallback_product_ids={source: "1000001"},
        )

    assert duplicate_rows == 0
    assert records[0]["归档商品ID"] == "1000001"
    assert records[0]["归档日期"] == "2026-08-01"


def test_matching_operation_log_content_corrects_malformed_suffix() -> None:
    daily_source = Path(
        "商品推广_分天数据_20260703至20260801(庄品健粮油旗舰店)"
        "(新大米_97770676548).xls"
    )
    log_source = Path(
        "商品推广操作日志_260802(庄品健粮油旗舰店)"
        "(新大米_97770676548).xls"
    )
    common = dict(
        status="ready",
        reason="ok",
        target_folder="推广数据",
        shop_name="庄品健粮油旗舰店",
        shop_source="test",
        period="2026-07-03至2026-08-01",
        period_source="date_column",
        download_date="2026-08-02",
        target_path=Path("target.xlsx"),
        product_counts="",
        matched_headers="",
        missing_fields="",
        map_updates=[],
        new_product_ids=[],
    )
    daily = organizer.Analysis(
        source=daily_source,
        report_type="推广数据",
        product_ids=["97770676548"],
        **common,
    )
    log = organizer.Analysis(
        source=log_source,
        report_type="推广调整日志",
        product_ids=["977770676548"],
        **common,
    )
    organizer.apply_pdd_daily_breakdown_metadata_evidence(
        [daily, log],
        {"977770676548": "庄品健粮油旗舰店"},
        ["庄品健粮油旗舰店"],
    )

    assert daily.status == "ready"
    assert daily.product_ids == ["977770676548"]
    assert daily.new_product_ids == []
    assert "id_source=operation_log_content" in daily.product_counts


def main() -> None:
    test_today_overlap_replaces_yesterday_record()
    test_different_spend_is_not_deduplicated()
    test_promotion_type_is_part_of_the_safe_scope()
    test_all_supported_spend_headers_are_normalized()
    test_transactional_merge_keeps_current_source_once()
    test_strict_daily_breakdown_uses_validated_filename_id()
    test_matching_operation_log_content_corrects_malformed_suffix()
    print("PASS: 近30天推广重叠去重回归测试全部通过")


if __name__ == "__main__":
    main()
