#!/usr/bin/env python3
"""Add deterministic E-number evidence mappings to an existing summary run.

This is a local migration utility. It does not call a model, search engine, or any
network service. By default it writes a sibling directory and leaves the original
summary untouched.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import shutil
import sys
from pathlib import Path
from typing import Any

from summarize_inquiry_evidence import (
    SCRIPT_VERSION,
    SUMMARY_FIELDS,
    enrich_summary_with_citation_details,
    parse_evidence_pack,
    validate_summary,
    write_results,
)


def load_json(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(value, dict):
        raise ValueError(f"JSON 顶层必须是对象：{path}")
    return value


def backfill(summary_dir: Path, output_dir: Path) -> int:
    metadata = load_json(summary_dir / "summary_run_metadata.json")
    input_dir = Path(str(metadata.get("input_dir") or ""))
    pack_path = input_dir / "llm_evidence_pack.md"
    if not pack_path.exists():
        raise FileNotFoundError(f"未找到原证据包：{pack_path}")

    sections = parse_evidence_pack(pack_path.read_text(encoding="utf-8"))
    by_cas = {section.cas: section for section in sections}
    source_rows = [
        json.loads(line)
        for line in (summary_dir / "summary_results.jsonl").read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
    results: list[dict[str, Any]] = []
    for source_row in source_rows:
        cas = str(source_row.get("cas") or "")
        section = by_cas.get(cas)
        if section is None:
            raise ValueError(f"原证据包中找不到 CAS {cas}")
        model_fields = {field: source_row[field] for field in SUMMARY_FIELDS}
        validated = validate_summary(model_fields, section)
        results.append(enrich_summary_with_citation_details(validated, section))

    usage_path = summary_dir / "summary_token_usage.json"
    usages = load_json(usage_path).get("products", []) if usage_path.exists() else []
    errors = metadata.get("errors", [])
    metadata = dict(metadata)
    metadata.update(
        {
            "output_dir": str(output_dir.resolve()),
            "citation_mapping_backfilled_at": dt.datetime.now(dt.timezone.utc).isoformat(
                timespec="seconds"
            ),
            "citation_mapping_source_dir": str(summary_dir.resolve()),
            "citation_mapping_summarizer_version": SCRIPT_VERSION,
            "citation_mapping_network_calls": 0,
            "citation_mapping_model_calls": 0,
        }
    )
    write_results(output_dir, results, usages, errors, metadata)

    raw_source = summary_dir / "raw_responses"
    raw_target = output_dir / "raw_responses"
    if raw_source.exists() and raw_source.resolve() != raw_target.resolve():
        shutil.copytree(raw_source, raw_target, dirs_exist_ok=True)
    return len(results)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="为历史总结结果回填 E 编号对应的证据、专利号和 URL（完全离线）"
    )
    parser.add_argument("summary_dir", help="现有 summary 输出目录")
    parser.add_argument("--output-dir", default="", help="新输出目录")
    args = parser.parse_args()
    summary_dir = Path(args.summary_dir)
    output_dir = (
        Path(args.output_dir)
        if args.output_dir
        else summary_dir.with_name(summary_dir.name + "-evidence-mapped")
    )
    if output_dir.resolve() == summary_dir.resolve():
        raise ValueError("请使用不同的输出目录，避免覆盖原历史结果")
    count = backfill(summary_dir, output_dir)
    print(f"已离线回填 {count} 个产品：{output_dir.resolve()}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"错误：{type(exc).__name__}: {exc}", file=sys.stderr)
        raise SystemExit(2)
