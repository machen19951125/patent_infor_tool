#!/usr/bin/env python3
"""Summarize a collected inquiry evidence pack with one model call per product.

The script never performs web search.  It sends only the already-collected product
section to an OpenAI Responses-compatible endpoint and requires strict JSON output.
Use --dry-run to export prompts and validate the full local pipeline without a key
or network connection.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import math
import os
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import requests


SCRIPT_VERSION = "0.2.3"
PROVIDER_DEFAULTS = {
    "openai": {
        "base_url": "https://api.openai.com/v1",
        "api_key_env": "OPENAI_API_KEY",
        "api_family": "responses",
    },
    "deepseek": {
        "base_url": "https://api.deepseek.com",
        "api_key_env": "DEEPSEEK_API_KEY",
        "api_family": "chat_completions",
        "model": "deepseek-v4-flash",
    },
}
SECTION_RE = re.compile(
    r"^##\s+\d+\.\s+(?P<bd>[^|\r\n]+?)\s+\|\s+CAS\s+(?P<cas>[^\r\n]+?)\s*$",
    re.MULTILINE,
)
EVIDENCE_ID_RE = re.compile(r"^####\s+(E\d+)\s+\|", re.MULTILINE)
EVIDENCE_HEADER_RE = re.compile(
    r"^####\s+(?P<evidence_id>E\d+)\s+\|\s+(?P<source>[^|\r\n]+?)\s+\|\s+"
    r"(?P<source_type>[^|\r\n]+?)(?:\s+\|[^\r\n]*)?$",
    re.MULTILINE,
)
PATENT_NUMBER_RE = re.compile(
    r"\b(?:AT|AU|BE|BR|CA|CH|CN|DE|DK|EP|ES|FI|FR|GB|HK|IL|IN|IT|JP|KR|MX|NL|NO|RU|SE|SG|TW|US|WO|ZA)"
    r"(?:[-\s]?\d){4,}(?:[-\s]?[A-Z]\d?)?\b",
    re.IGNORECASE,
)
MISSING_BD_LABEL = "无BD"


@dataclass(frozen=True)
class EvidenceReference:
    evidence_id: str
    source: str
    source_type: str
    title: str = ""
    identifiers: str = ""
    url: str = ""


@dataclass(frozen=True)
class ProductSection:
    bd: str
    cas: str
    text: str
    evidence_ids: tuple[str, ...]
    evidence_references: tuple[EvidenceReference, ...] = ()
    bd_missing: bool = False


SUMMARY_FIELDS = [
    "bd",
    "cas",
    "chemical_name_en",
    "category",
    "associated_drug_or_project",
    "inquiry_increase_reason",
    "evidence_grade",
    "confidence",
    "business_priority",
    "follow_up_and_quote_points",
    "evidence_rationale",
    "cited_evidence_ids",
    "needs_human_review",
]

CSV_FIELDS = SUMMARY_FIELDS + [
    "cited_evidence_references",
    "cited_verified_patent_numbers",
    "cited_candidate_patent_numbers",
    "cited_patent_urls",
]

SUMMARY_SCHEMA: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "bd": {"type": "string"},
        "cas": {"type": "string"},
        "chemical_name_en": {"type": "string"},
        "category": {"type": "string"},
        "associated_drug_or_project": {"type": "string"},
        "inquiry_increase_reason": {"type": "string"},
        "evidence_grade": {"type": "string", "enum": ["A", "B", "C"]},
        "confidence": {"type": "string", "enum": ["高", "中", "低"]},
        "business_priority": {"type": "string", "enum": ["高", "中", "低"]},
        "follow_up_and_quote_points": {"type": "string"},
        "evidence_rationale": {"type": "string"},
        "cited_evidence_ids": {
            "type": "array",
            "items": {"type": "string"},
        },
        "needs_human_review": {"type": "boolean"},
    },
    "required": SUMMARY_FIELDS,
}

SYSTEM_PROMPT = """你是医药与精细化工询单证据分析员。仅使用用户给出的单个产品证据包，不使用未提供的知识，不联网。最终只输出一个 JSON 对象，不得输出 Markdown 代码块或 JSON 之外的解释。

严格区分事实与推断：商业目录可支持名称、市售形式和可得性，不能单独证明下游药物、合成路线或需求上升原因。PubChem 专利索引只是候选专利号集合，不等于原始专利正文证据。只有 source_type=verified_original_patent 且提供正文命中范围和片段的条目可作为专利正文证据；名称/同义词命中不等于 CAS 命中，相关化学形式也不等于同一物料。如没有直接的市场、监管、临床、公司或论文事件，必须明说“未见直接的近期需求事件证据”，只做保守业务判断。

证据等级：A=监管机构、原始专利/论文或权威结构化数据库直接支持关键身份/用途；B=多家独立商业来源一致，且结构、路线或事件逻辑吻合；C=仅有商业目录或结构推断。等级表示核心业务结论的证据强度，不是 PubChem 身份记录本身的可靠性。

只引用证据包中存在且真正支持结论的 E 编号。中文简洁输出，不得伪造药物、项目、市场事件、时间或数字。""".strip()


def _evidence_field(block: str, label: str) -> str:
    match = re.search(rf"^-\s+{re.escape(label)}：(.*)$", block, re.MULTILINE)
    return match.group(1).strip() if match else ""


def parse_evidence_references(section_text: str) -> tuple[EvidenceReference, ...]:
    matches = list(EVIDENCE_HEADER_RE.finditer(section_text))
    references: list[EvidenceReference] = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(section_text)
        block = section_text[match.start() : end]
        references.append(
            EvidenceReference(
                evidence_id=match.group("evidence_id").strip(),
                source=match.group("source").strip(),
                source_type=match.group("source_type").strip(),
                title=_evidence_field(block, "标题"),
                identifiers=_evidence_field(block, "标识符"),
                url=_evidence_field(block, "URL"),
            )
        )
    return tuple(references)


def parse_evidence_pack(text: str) -> list[ProductSection]:
    matches = list(SECTION_RE.finditer(text))
    sections: list[ProductSection] = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        section_text = text[match.start() : end].strip()
        ids = tuple(EVIDENCE_ID_RE.findall(section_text))
        references = parse_evidence_references(section_text)
        bd = match.group("bd").strip()
        sections.append(
            ProductSection(
                bd=bd,
                cas=match.group("cas").strip(),
                text=section_text,
                evidence_ids=ids,
                evidence_references=references,
                bd_missing=bd == MISSING_BD_LABEL,
            )
        )
    if not sections:
        raise ValueError("证据包中未找到产品章节")
    return sections


def extract_patent_numbers(reference: EvidenceReference) -> list[str]:
    """Return normalized publication numbers without treating every E item as a patent."""
    if "patent" not in reference.source_type.lower():
        return []
    numbers: list[str] = []
    for value in (reference.identifiers, reference.title, reference.url):
        for match in PATENT_NUMBER_RE.findall(value):
            normalized = re.sub(r"[-\s]", "", match).upper()
            if normalized not in numbers:
                numbers.append(normalized)
    return numbers


def enrich_summary_with_citation_details(
    summary: dict[str, Any], section: ProductSection
) -> dict[str, Any]:
    """Attach deterministic E-number mappings after model/schema validation."""
    by_id = {item.evidence_id: item for item in section.evidence_references}
    details: list[dict[str, Any]] = []
    verified_patents: list[str] = []
    candidate_patents: list[str] = []
    patent_urls: list[str] = []
    for evidence_id in summary["cited_evidence_ids"]:
        reference = by_id.get(evidence_id)
        if reference is None:
            continue
        patent_numbers = extract_patent_numbers(reference)
        if reference.source_type == "verified_original_patent":
            patent_status = "verified"
            target = verified_patents
        elif patent_numbers:
            patent_status = "candidate"
            target = candidate_patents
        else:
            patent_status = "not_patent"
            target = []
        for patent_number in patent_numbers:
            if patent_number not in target:
                target.append(patent_number)
        if patent_numbers and reference.url and reference.url not in patent_urls:
            patent_urls.append(reference.url)
        details.append(
            {
                "evidence_id": reference.evidence_id,
                "source": reference.source,
                "source_type": reference.source_type,
                "title": reference.title,
                "identifiers": reference.identifiers,
                "url": reference.url,
                "patent_status": patent_status,
                "patent_numbers": patent_numbers,
            }
        )
    enriched = dict(summary)
    enriched["cited_evidence_details"] = details
    enriched["cited_verified_patent_numbers"] = verified_patents
    enriched["cited_candidate_patent_numbers"] = candidate_patents
    enriched["cited_patent_urls"] = patent_urls
    return enriched


def build_user_prompt(section: ProductSection, include_schema: bool = False) -> str:
    schema_instruction = ""
    if include_schema:
        schema_instruction = (
            "\n\n必须严格使用以下 JSON Schema 的所有字段，不得增删字段：\n"
            + json.dumps(SUMMARY_SCHEMA, ensure_ascii=False, separators=(",", ":"))
        )
    return f"""请分析以下唯一产品，输出符合 JSON Schema 的结果。

字段口径：
- bd / cas：必须使用证据包标题中的产品标识。标题为“无BD”时，bd 填“无BD”；CAS 必须原样返回。
- chemical_name_en：使用证据包中的 PubChem preferred name 或 IUPAC English name。保留英文原文、立体化学标记和标点，不翻译成中文，不自行缩写。
- category：商业/研发分类，如原料药中间体、药化砌块、杂质对照品、分析标准、工艺试剂或通用无机盐。
- associated_drug_or_project：只写有证据的关联；否则写“未锁定单一药物/项目”。
- inquiry_increase_reason：先写证据能证明的需求场景；没有近期事件时必须明示保留。
- evidence_rationale：说明等级的核心依据及限制，最多 100 个汉字。
- needs_human_review：下游归属、立体形式、商业需求原因或身份存在实质不确定性时为 true。
- MedChemExpress (MCE) 产品页视为人工整理的二级药理资料，可支持靶点、机制、通路和实验背景判断；但不取代其引用的原始论文/专利，也不单独证明近期询单增长。
{schema_instruction}

产品证据包：
{section.text}
""".strip()


def build_request_payload(
    section: ProductSection,
    model: str,
    max_output_tokens: int = 1400,
    reasoning_effort: str = "",
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "model": model,
        "store": False,
        "max_output_tokens": max_output_tokens,
        "input": [
            {"role": "developer", "content": SYSTEM_PROMPT},
            {"role": "user", "content": build_user_prompt(section)},
        ],
        "text": {
            "format": {
                "type": "json_schema",
                "name": "inquiry_product_summary",
                "strict": True,
                "schema": SUMMARY_SCHEMA,
            }
        },
    }
    if reasoning_effort:
        payload["reasoning"] = {"effort": reasoning_effort}
    return payload


def build_deepseek_request_payload(
    section: ProductSection,
    model: str,
    max_output_tokens: int = 1400,
    thinking_mode: str = "enabled",
    reasoning_effort: str = "high",
    temperature: float = 0.2,
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": build_user_prompt(section, include_schema=True)},
        ],
        "response_format": {"type": "json_object"},
        "max_tokens": max_output_tokens,
        "temperature": temperature,
        "stream": False,
        "thinking": {"type": thinking_mode},
    }
    if thinking_mode == "enabled" and reasoning_effort:
        payload["reasoning_effort"] = reasoning_effort
    return payload


def build_provider_payload(
    provider: str,
    section: ProductSection,
    model: str,
    max_output_tokens: int,
    reasoning_effort: str,
    thinking_mode: str,
    temperature: float,
) -> dict[str, Any]:
    if provider == "deepseek":
        return build_deepseek_request_payload(
            section,
            model,
            max_output_tokens=max_output_tokens,
            thinking_mode=thinking_mode,
            reasoning_effort=reasoning_effort or "high",
            temperature=temperature,
        )
    return build_request_payload(
        section,
        model,
        max_output_tokens=max_output_tokens,
        reasoning_effort=reasoning_effort,
    )


def extract_output_text(response: dict[str, Any]) -> str:
    if isinstance(response.get("output_text"), str) and response["output_text"].strip():
        return response["output_text"].strip()
    refusals: list[str] = []
    texts: list[str] = []
    for item in response.get("output") or []:
        if not isinstance(item, dict):
            continue
        for content in item.get("content") or []:
            if not isinstance(content, dict):
                continue
            if content.get("type") == "output_text" and isinstance(content.get("text"), str):
                texts.append(content["text"])
            if content.get("type") == "refusal" and isinstance(content.get("refusal"), str):
                refusals.append(content["refusal"])
    if refusals:
        raise ValueError("模型拒绝输出：" + " ".join(refusals))
    if not texts:
        raise ValueError("模型响应中未找到 output_text")
    return "\n".join(texts).strip()


def extract_deepseek_output_text(response: dict[str, Any]) -> str:
    choices = response.get("choices") or []
    if not choices or not isinstance(choices[0], dict):
        raise ValueError("DeepSeek 响应中未找到 choices[0]")
    choice = choices[0]
    if choice.get("finish_reason") == "length":
        raise ValueError("DeepSeek 输出因 max_tokens 不足被截断")
    message = choice.get("message") or {}
    content = message.get("content") if isinstance(message, dict) else None
    if not isinstance(content, str) or not content.strip():
        raise ValueError("DeepSeek 返回了空的 JSON 内容")
    return content.strip()


def extract_provider_output_text(provider: str, response: dict[str, Any]) -> str:
    return extract_deepseek_output_text(response) if provider == "deepseek" else extract_output_text(response)


def validate_summary(value: Any, section: ProductSection) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ValueError("总结结果不是 JSON 对象")
    missing = [field for field in SUMMARY_FIELDS if field not in value]
    extras = [field for field in value if field not in SUMMARY_FIELDS]
    if missing or extras:
        raise ValueError(f"字段不匹配；缺失={missing}；额外={extras}")
    if value["cas"] != section.cas:
        raise ValueError("模型输出的 CAS 与当前产品不一致")
    if section.bd_missing:
        # BD is absent from single-CAS input.  Do not require the model to
        # reproduce a display-only placeholder; write the canonical value
        # deterministically so validation retries do not waste API tokens.
        value["bd"] = MISSING_BD_LABEL
    elif value["bd"] != section.bd:
        raise ValueError("模型输出的 BD 与当前产品不一致")
    if value["evidence_grade"] not in {"A", "B", "C"}:
        raise ValueError("证据等级必须是 A/B/C")
    if value["confidence"] not in {"高", "中", "低"}:
        raise ValueError("置信度必须是高/中/低")
    if value["business_priority"] not in {"高", "中", "低"}:
        raise ValueError("业务优先级必须是高/中/低")
    if not re.search(r"[A-Za-z]", str(value["chemical_name_en"])):
        raise ValueError("chemical_name_en 必须保留英文化学名称")
    for field in SUMMARY_FIELDS:
        if field in {"cited_evidence_ids", "needs_human_review"}:
            continue
        if not isinstance(value[field], str) or not value[field].strip():
            raise ValueError(f"字段 {field} 不能为空")
    if not isinstance(value["needs_human_review"], bool):
        raise ValueError("needs_human_review 必须是布尔值")
    citations = value["cited_evidence_ids"]
    if not isinstance(citations, list) or not all(isinstance(item, str) for item in citations):
        raise ValueError("cited_evidence_ids 必须是字符串数组")
    invalid = sorted(set(citations) - set(section.evidence_ids))
    if invalid:
        raise ValueError(f"引用了不存在的证据编号：{invalid}")
    if not citations:
        raise ValueError("至少引用一条证据")
    return value


def estimate_tokens(text: str) -> int:
    cjk = sum(1 for character in text if "\u3400" <= character <= "\u9fff")
    return cjk + math.ceil((len(text) - cjk) / 4)


class ModelApiClient:
    def __init__(
        self,
        provider: str,
        base_url: str,
        api_key: str,
        timeout: float,
        retries: int,
    ) -> None:
        suffix = "/chat/completions" if provider == "deepseek" else "/responses"
        self.endpoint = base_url.rstrip("/") + suffix
        self.provider = provider
        self.api_key = api_key
        self.timeout = timeout
        self.retries = max(0, retries)
        self.session = requests.Session()

    def create(self, payload: dict[str, Any]) -> dict[str, Any]:
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "User-Agent": f"InquiryEvidenceSummarizer/{SCRIPT_VERSION}",
        }
        last_error: Exception | None = None
        for attempt in range(self.retries + 1):
            try:
                response = self.session.post(
                    self.endpoint,
                    headers=headers,
                    json=payload,
                    timeout=self.timeout,
                )
                if response.status_code == 429 or response.status_code >= 500:
                    raise requests.HTTPError(
                        f"HTTP {response.status_code}: {response.text[:800]}",
                        response=response,
                    )
                response.raise_for_status()
                data = response.json()
                if not isinstance(data, dict):
                    raise ValueError("API 返回值不是 JSON 对象")
                return data
            except (requests.RequestException, ValueError) as exc:
                last_error = exc
                if attempt >= self.retries:
                    break
                time.sleep(min(2**attempt, 8))
        assert last_error is not None
        raise last_error


def usage_row(
    section: ProductSection,
    response: dict[str, Any],
    provider: str = "openai",
    attempt: int = 1,
) -> dict[str, Any]:
    usage = response.get("usage") or {}
    if provider == "deepseek":
        input_tokens = int(usage.get("prompt_tokens") or 0)
        cached_input_tokens = int(usage.get("prompt_cache_hit_tokens") or 0)
        output_tokens = int(usage.get("completion_tokens") or 0)
        output_details = usage.get("completion_tokens_details") or {}
        reasoning_tokens = int(output_details.get("reasoning_tokens") or 0)
    else:
        input_tokens = int(usage.get("input_tokens") or 0)
        input_details = usage.get("input_tokens_details") or {}
        cached_input_tokens = int(input_details.get("cached_tokens") or 0)
        output_tokens = int(usage.get("output_tokens") or 0)
        output_details = usage.get("output_tokens_details") or {}
        reasoning_tokens = int(output_details.get("reasoning_tokens") or 0)
    return {
        "bd": section.bd,
        "cas": section.cas,
        "provider": provider,
        "model": str(response.get("model") or ""),
        "api_calls": 1,
        "input_tokens": input_tokens,
        "cached_input_tokens": cached_input_tokens,
        "output_tokens": output_tokens,
        "reasoning_tokens": reasoning_tokens,
        "total_tokens": int(usage.get("total_tokens") or (input_tokens + output_tokens)),
        "response_ids": [str(response.get("id") or "")],
        "attempts": [attempt],
    }


def aggregate_usage_rows(rows: list[dict[str, Any]]) -> dict[str, Any]:
    if not rows:
        raise ValueError("无可汇总的 token 用量")
    first = rows[0]
    return {
        "bd": first["bd"],
        "cas": first["cas"],
        "provider": first["provider"],
        "model": next((row["model"] for row in reversed(rows) if row["model"]), ""),
        "api_calls": sum(int(row["api_calls"]) for row in rows),
        "input_tokens": sum(int(row["input_tokens"]) for row in rows),
        "cached_input_tokens": sum(int(row["cached_input_tokens"]) for row in rows),
        "output_tokens": sum(int(row["output_tokens"]) for row in rows),
        "reasoning_tokens": sum(int(row["reasoning_tokens"]) for row in rows),
        "total_tokens": sum(int(row["total_tokens"]) for row in rows),
        "response_ids": [item for row in rows for item in row["response_ids"]],
        "attempts": [item for row in rows for item in row["attempts"]],
    }


def format_evidence_reference(detail: dict[str, Any]) -> str:
    patent_numbers = detail.get("patent_numbers") or []
    status = detail.get("patent_status")
    if status == "verified":
        kind = f"已核验专利 {'/'.join(patent_numbers)}"
    elif status == "candidate":
        kind = f"候选专利 {'/'.join(patent_numbers)}"
    else:
        kind = str(detail.get("source_type") or "非专利证据")
    parts = [str(detail.get("evidence_id") or ""), kind]
    for value in (detail.get("source"), detail.get("title"), detail.get("url")):
        if value:
            parts.append(str(value))
    return " | ".join(parts)


def write_results(
    output_dir: Path,
    results: list[dict[str, Any]],
    usages: list[dict[str, Any]],
    errors: list[dict[str, str]],
    metadata: dict[str, Any],
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    jsonl_path = output_dir / "summary_results.jsonl"
    jsonl_path.write_text(
        "".join(json.dumps(item, ensure_ascii=False) + "\n" for item in results),
        encoding="utf-8",
    )
    with (output_dir / "summary_results.csv").open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_FIELDS)
        writer.writeheader()
        for item in results:
            row = {field: item.get(field, "") for field in SUMMARY_FIELDS}
            row["cited_evidence_ids"] = ";".join(item["cited_evidence_ids"])
            row["cited_evidence_references"] = " || ".join(
                format_evidence_reference(detail) for detail in item.get("cited_evidence_details", [])
            )
            row["cited_verified_patent_numbers"] = ";".join(
                item.get("cited_verified_patent_numbers", [])
            )
            row["cited_candidate_patent_numbers"] = ";".join(
                item.get("cited_candidate_patent_numbers", [])
            )
            row["cited_patent_urls"] = ";".join(item.get("cited_patent_urls", []))
            writer.writerow(row)

    report_lines = ["# 询单产品大模型总结", ""]
    for index, item in enumerate(results, start=1):
        report_lines.extend(
            [
                f"## {index}. {item['bd']} | CAS {item['cas']}",
                "",
                f"- Chemical name: {item['chemical_name_en']}",
                f"- 类别：{item['category']}",
                f"- 可能关联药物/项目：{item['associated_drug_or_project']}",
                f"- 询单增加原因判断：{item['inquiry_increase_reason']}",
                f"- 证据等级 / 置信度 / 业务优先级：{item['evidence_grade']} / {item['confidence']} / {item['business_priority']}",
                f"- 建议追问/报价要点：{item['follow_up_and_quote_points']}",
                f"- 证据说明：{item['evidence_rationale']}",
                f"- 引用证据：{', '.join(item['cited_evidence_ids'])}",
                f"- 需人工复核：{'是' if item['needs_human_review'] else '否'}",
                "",
                "### 引用证据索引",
                "",
            ]
        )
        details = item.get("cited_evidence_details", [])
        if details:
            report_lines.extend(f"- {format_evidence_reference(detail)}" for detail in details)
        else:
            report_lines.append("- 无可展开的证据元数据")
        verified = item.get("cited_verified_patent_numbers", [])
        candidates = item.get("cited_candidate_patent_numbers", [])
        report_lines.extend(
            [
                "",
                f"- 已核验专利号：{', '.join(verified) if verified else '无'}",
                f"- 候选专利号（未经正文核验）：{', '.join(candidates) if candidates else '无'}",
                "",
            ]
        )
    (output_dir / "summary_report.md").write_text("\n".join(report_lines), encoding="utf-8")

    total_usage = {
        "method": "exact API usage returned by the model provider",
        "products": usages,
        "total_input_tokens": sum(item["input_tokens"] for item in usages),
        "total_cached_input_tokens": sum(item["cached_input_tokens"] for item in usages),
        "total_output_tokens": sum(item["output_tokens"] for item in usages),
        "total_reasoning_tokens": sum(item["reasoning_tokens"] for item in usages),
        "total_tokens": sum(item["total_tokens"] for item in usages),
    }
    (output_dir / "summary_token_usage.json").write_text(
        json.dumps(total_usage, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    metadata["errors"] = errors
    (output_dir / "summary_run_metadata.json").write_text(
        json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def select_sections(
    sections: list[ProductSection], limit: int, only_cas: set[str]
) -> list[ProductSection]:
    selected = [section for section in sections if not only_cas or section.cas in only_cas]
    if limit > 0:
        selected = selected[:limit]
    if not selected:
        raise ValueError("没有产品符合当前筛选条件")
    return selected


def run(args: argparse.Namespace) -> int:
    input_dir = Path(args.input_dir)
    pack_path = input_dir / "llm_evidence_pack.md"
    if not pack_path.exists():
        raise FileNotFoundError(f"未找到证据包：{pack_path}")
    sections = parse_evidence_pack(pack_path.read_text(encoding="utf-8"))
    only_cas = {item.strip() for item in args.only_cas.split(",") if item.strip()}
    sections = select_sections(sections, args.limit, only_cas)

    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    output_dir = Path(args.output_dir) if args.output_dir else input_dir / f"summary-{stamp}"
    output_dir.mkdir(parents=True, exist_ok=True)

    provider = args.provider
    defaults = PROVIDER_DEFAULTS[provider]
    model = args.model.strip()
    if not model and provider == "openai":
        model = os.getenv("OPENAI_MODEL", "").strip()
    if not model:
        model = str(defaults.get("model") or "")
    base_url = args.base_url.strip() or str(defaults["base_url"])
    api_key_env = args.api_key_env.strip() or str(defaults["api_key_env"])
    if not 0 <= args.temperature <= 2:
        raise ValueError("temperature 必须介于 0 和 2 之间")
    metadata: dict[str, Any] = {
        "script_version": SCRIPT_VERSION,
        "started_at": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "input_dir": str(input_dir.resolve()),
        "output_dir": str(output_dir.resolve()),
        "product_count": len(sections),
        "dry_run": bool(args.dry_run),
        "web_search_enabled": False,
        "provider": provider,
        "api_family": defaults["api_family"],
        "structured_output_mode": (
            "json_object_plus_local_schema_validation"
            if provider == "deepseek"
            else "strict_json_schema"
        ),
        "model": model,
        "base_url": base_url,
        "max_output_tokens": args.max_output_tokens,
        "reasoning_effort": args.reasoning_effort,
        "thinking_mode": args.thinking_mode if provider == "deepseek" else None,
        "temperature": args.temperature if provider == "deepseek" else None,
        "validation_retries": args.validation_retries,
    }

    if args.dry_run:
        prompt_dir = output_dir / "prompts"
        prompt_dir.mkdir(parents=True, exist_ok=True)
        manifest: list[dict[str, Any]] = []
        for section in sections:
            payload = build_provider_payload(
                provider,
                section,
                model or "__MODEL_REQUIRED_FOR_LIVE_RUN__",
                args.max_output_tokens,
                args.reasoning_effort,
                args.thinking_mode,
                args.temperature,
            )
            prompt_path = prompt_dir / f"{section.bd}__{section.cas}.request.json"
            prompt_path.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            input_text = SYSTEM_PROMPT + "\n" + build_user_prompt(
                section, include_schema=provider == "deepseek"
            )
            manifest.append(
                {
                    "bd": section.bd,
                    "cas": section.cas,
                    "provider": provider,
                    "model": model or "__MODEL_REQUIRED_FOR_LIVE_RUN__",
                    "evidence_items": len(section.evidence_ids),
                    "estimated_input_tokens": estimate_tokens(input_text),
                    "request_file": str(prompt_path.relative_to(output_dir)),
                }
            )
        (output_dir / "prompt_manifest.json").write_text(
            json.dumps(
                {
                    "method": "CJK characters + ceil(other characters / 4); estimate only",
                    "products": manifest,
                    "estimated_total_input_tokens": sum(
                        item["estimated_input_tokens"] for item in manifest
                    ),
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        metadata["finished_at"] = dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")
        metadata["status"] = "prompts_exported"
        (output_dir / "summary_run_metadata.json").write_text(
            json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"已离线导出 {len(sections)} 个产品的模型请求：{output_dir.resolve()}")
        return 0

    api_key = os.getenv(api_key_env, "").strip()
    if not api_key:
        raise ValueError(f"未配置环境变量 {api_key_env}")
    if not model:
        raise ValueError("真实调用必须通过 --model 指定模型")

    client = ModelApiClient(provider, base_url, api_key, args.timeout, args.retries)
    raw_dir = output_dir / "raw_responses"
    raw_dir.mkdir(parents=True, exist_ok=True)
    results: list[dict[str, Any]] = []
    usages: list[dict[str, Any]] = []
    errors: list[dict[str, str]] = []

    for index, section in enumerate(sections, start=1):
        print(
            f"[{index}/{len(sections)}] {provider} 总结 {section.bd} | {section.cas}",
            flush=True,
        )
        payload = build_provider_payload(
            provider,
            section,
            model,
            args.max_output_tokens,
            args.reasoning_effort,
            args.thinking_mode,
            args.temperature,
        )
        attempt_usage: list[dict[str, Any]] = []
        try:
            summary: dict[str, Any] | None = None
            last_validation_error: Exception | None = None
            for validation_attempt in range(args.validation_retries + 1):
                response = client.create(payload)
                attempt_number = validation_attempt + 1
                attempt_usage.append(
                    usage_row(section, response, provider=provider, attempt=attempt_number)
                )
                raw_name = (
                    f"{section.bd}__{section.cas}.json"
                    if attempt_number == 1
                    else f"{section.bd}__{section.cas}.attempt-{attempt_number}.json"
                )
                (raw_dir / raw_name).write_text(
                    json.dumps(response, ensure_ascii=False, indent=2), encoding="utf-8"
                )
                try:
                    output_text = extract_provider_output_text(provider, response)
                    summary = validate_summary(json.loads(output_text), section)
                    break
                except (ValueError, json.JSONDecodeError) as exc:
                    last_validation_error = exc
                    if validation_attempt < args.validation_retries:
                        print(
                            f"  JSON/业务 Schema 校验失败，正在重试 "
                            f"({attempt_number}/{args.validation_retries + 1})：{exc}",
                            flush=True,
                        )
            if summary is None:
                assert last_validation_error is not None
                raise last_validation_error
            results.append(enrich_summary_with_citation_details(summary, section))
        except Exception as exc:
            errors.append(
                {
                    "bd": section.bd,
                    "cas": section.cas,
                    "error": f"{type(exc).__name__}: {exc}",
                }
            )
            print(f"  失败：{type(exc).__name__}: {exc}", flush=True)
        finally:
            if attempt_usage:
                usages.append(aggregate_usage_rows(attempt_usage))

    metadata["finished_at"] = dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")
    metadata["successful_products"] = len(results)
    metadata["failed_products"] = len(errors)
    write_results(output_dir, results, usages, errors, metadata)
    print(
        f"总结完成：成功 {len(results)}，失败 {len(errors)}，输出 {output_dir.resolve()}",
        flush=True,
    )
    return 1 if errors else 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="将已采集的询单证据包逐产品交给大模型做严格 JSON 总结"
    )
    parser.add_argument("input_dir", help="包含 llm_evidence_pack.md 的采集输出目录")
    parser.add_argument("--output-dir", default="", help="总结输出目录")
    parser.add_argument("--dry-run", action="store_true", help="只离线导出请求，不调用模型")
    parser.add_argument("--provider", choices=["openai", "deepseek"], default="openai")
    parser.add_argument("--model", default="", help="模型名称")
    parser.add_argument("--base-url", default="", help="API 根地址；留空使用提供商官方地址")
    parser.add_argument("--api-key-env", default="", help="API 密钥环境变量名")
    parser.add_argument("--limit", type=int, default=0, help="只处理前 N 个产品，0 为全部")
    parser.add_argument("--only-cas", default="", help="只处理指定 CAS，多个用逗号分隔")
    parser.add_argument("--max-output-tokens", type=int, default=1400)
    parser.add_argument(
        "--reasoning-effort",
        choices=["", "none", "low", "medium", "high", "xhigh", "max"],
        default="",
    )
    parser.add_argument("--thinking-mode", choices=["enabled", "disabled"], default="enabled")
    parser.add_argument("--temperature", type=float, default=0.2)
    parser.add_argument("--validation-retries", type=int, default=1)
    parser.add_argument("--timeout", type=float, default=180.0)
    parser.add_argument("--retries", type=int, default=2)
    return parser


def main() -> int:
    try:
        return run(build_parser().parse_args())
    except Exception as exc:
        print(f"错误：{type(exc).__name__}: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
