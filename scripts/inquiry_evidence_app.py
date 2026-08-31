#!/usr/bin/env python3
"""Streamlit UI for the inquiry evidence collector."""

from __future__ import annotations

import datetime as dt
import hashlib
import io
import json
import os
import queue
import re
import subprocess
import sys
import threading
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import streamlit as st


PROJECT_DIR = Path(__file__).resolve().parents[1]
COLLECTOR = PROJECT_DIR / "scripts" / "collect_inquiry_evidence.py"
SUMMARIZER = PROJECT_DIR / "scripts" / "summarize_inquiry_evidence.py"
EXAMPLE_FILE = PROJECT_DIR / "询单范例.xlsx"
UPLOAD_DIR = PROJECT_DIR / ".cache" / "gui_uploads"
SETTINGS_FILE = PROJECT_DIR / ".cache" / "gui_settings.json"
CAS_RE = re.compile(r"^\d{2,7}-\d{2}-\d$")
SOURCE_OPTIONS = {
    "pubchem": "PubChem 身份与专利索引",
    "googlepatents": "Google Patents 专利正文",
    "chembl": "ChEMBL 药物与活性",
    "europepmc": "Europe PMC 文献",
    "pubmed": "PubMed 文献",
    "clinicaltrials": "ClinicalTrials.gov 临床试验",
    "tavily": "Tavily 开放网页检索",
    "mce": "MedChemExpress (MCE) 药理与靶点资料（需 Tavily Key）",
    "brave": "Brave 开放网页检索（备用）",
}
DEFAULT_SOURCES = [name for name in SOURCE_OPTIONS if name not in {"brave", "mce"}]
PERSISTED_SETTING_KEYS = {
    "input_mode",
    "single_cas",
    "use_example",
    "last_input_path",
    "max_workers",
    "web_results",
    "pack_items",
    "patent_verification",
    "mce_max_products",
    "offline",
    "include_recent_events",
    "resume_selected_dir",
    "tavily_key",
    "brave_key",
    "selected_sources",
    "summary_input_text",
    "summary_provider",
    "summary_dry_run",
    "openai_model_choice",
    "openai_custom_model",
    "openai_reasoning_effort",
    "openai_summary_key",
    "openai_base_url",
    "deepseek_model_choice",
    "deepseek_custom_model",
    "deepseek_thinking_mode",
    "deepseek_reasoning_effort",
    "deepseek_temperature",
    "deepseek_summary_key",
    "deepseek_base_url",
    "max_output_tokens",
    "validation_retries",
}


@dataclass
class RunState:
    process: subprocess.Popen[str]
    output_dir: Path
    events: queue.Queue[tuple[str, object]] = field(default_factory=queue.Queue)
    logs: list[str] = field(default_factory=list)
    return_code: int | None = None
    error: str = ""
    stopped: bool = False


def load_gui_settings() -> dict[str, object]:
    try:
        value = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except (FileNotFoundError, OSError, json.JSONDecodeError):
        return {}
    return value if isinstance(value, dict) else {}


def save_gui_settings(values: dict[str, object]) -> None:
    SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
    payload = {key: values[key] for key in sorted(PERSISTED_SETTING_KEYS) if key in values}
    temporary = SETTINGS_FILE.with_suffix(".tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    temporary.replace(SETTINGS_FILE)


def persisted_value(settings: dict[str, object], key: str, default: object) -> object:
    return settings.get(key, default)


def persist_current_session(**extra: object) -> None:
    values = load_gui_settings()
    values.update({
        key: st.session_state[key]
        for key in PERSISTED_SETTING_KEYS
        if key in st.session_state
    })
    values.update(extra)
    save_gui_settings(values)


def cas_is_valid(value: str) -> bool:
    value = value.strip()
    if not CAS_RE.fullmatch(value):
        return False
    body, check = value.rsplit("-", 1)
    digits = body.replace("-", "")
    total = sum(int(character) * multiplier for multiplier, character in enumerate(reversed(digits), start=1))
    return total % 10 == int(check)


def save_upload(uploaded_file: object) -> Path:
    name = Path(getattr(uploaded_file, "name", "inquiry.xlsx")).name
    suffix = Path(name).suffix.lower()
    if suffix not in {".xlsx", ".xlsm", ".csv"}:
        raise ValueError("仅支持 .xlsx、.xlsm 或 .csv 文件")
    content = uploaded_file.getvalue()  # type: ignore[attr-defined]
    digest = hashlib.sha256(content).hexdigest()[:16]
    safe_stem = "".join(
        character if character.isalnum() or character in "._-" else "_"
        for character in Path(name).stem
    )
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    target = UPLOAD_DIR / f"{safe_stem or 'inquiry'}-{digest}{suffix}"
    if not target.exists():
        target.write_bytes(content)
    return target


def new_output_dir() -> Path:
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S-%f")[:-3]
    return PROJECT_DIR / "outputs" / f"streamlit-{stamp}"


def new_summary_output_dir() -> Path:
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S-%f")[:-3]
    return PROJECT_DIR / "outputs" / f"summary-streamlit-{stamp}"


def latest_evidence_output_dir() -> Path | None:
    output_root = PROJECT_DIR / "outputs"
    if not output_root.exists():
        return None
    candidates = [
        path
        for path in output_root.iterdir()
        if path.is_dir() and (path / "llm_evidence_pack.md").exists()
    ]
    return max(candidates, key=lambda path: path.stat().st_mtime) if candidates else None


def resumable_evidence_output_dirs() -> list[tuple[Path, dict[str, object]]]:
    output_root = PROJECT_DIR / "outputs"
    if not output_root.exists():
        return []
    candidates: list[tuple[Path, dict[str, object]]] = []
    for path in output_root.iterdir():
        state_path = path / "resume_state.json"
        if not path.is_dir() or not state_path.exists():
            continue
        try:
            state = json.loads(state_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        if not isinstance(state, dict):
            continue
        remaining = int(state.get("remaining_products") or 0)
        if state.get("status") != "complete" and remaining > 0:
            candidates.append((path, state))
    return sorted(candidates, key=lambda item: item[0].stat().st_mtime, reverse=True)


def reader_worker(state: RunState) -> None:
    try:
        assert state.process.stdout is not None
        for line in state.process.stdout:
            state.events.put(("log", line.rstrip("\r\n")))
        state.events.put(("done", state.process.wait()))
    except Exception as exc:
        state.events.put(("error", f"{type(exc).__name__}: {exc}"))


def launch_collector(
    input_path: Path | None,
    single_cas: str,
    sources: list[str],
    tavily_key: str,
    brave_key: str,
    limit: int,
    max_workers: int,
    web_results: int,
    pack_items: int,
    patent_verification: str,
    mce_max_products: int,
    offline: bool,
    include_recent_events: bool,
    resume_dir: Path | None = None,
) -> RunState:
    output_dir = resume_dir if resume_dir is not None else new_output_dir()
    command = [
        sys.executable,
        "-u",
        str(COLLECTOR),
        "--max-workers",
        str(max(1, max_workers)),
    ]
    if resume_dir is not None:
        command.extend(["--resume-dir", str(resume_dir)])
    else:
        command.extend(
            [
                "--output-dir",
                str(output_dir),
                "--sources",
                ",".join(sources),
                "--limit",
                str(max(0, limit)),
                "--web-results",
                str(max(1, web_results)),
                "--pack-items",
                str(max(1, pack_items)),
                "--patent-verification",
                patent_verification
                if patent_verification in {"off", "balanced", "deep"}
                else "balanced",
                "--mce-max-products",
                str(max(0, mce_max_products)),
            ]
        )
        if single_cas:
            command.extend(["--cas", single_cas])
        elif input_path is not None:
            command.append(str(input_path))
        else:
            raise ValueError("未提供询单文件或单个 CAS")
    if offline:
        command.append("--offline")
    if include_recent_events:
        command.append("--include-recent-events")

    environment = os.environ.copy()
    if tavily_key:
        environment["TAVILY_API_KEY"] = tavily_key
    else:
        environment.pop("TAVILY_API_KEY", None)
    if brave_key:
        environment["BRAVE_SEARCH_API_KEY"] = brave_key
    else:
        environment.pop("BRAVE_SEARCH_API_KEY", None)

    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0
    process = subprocess.Popen(
        command,
        cwd=PROJECT_DIR,
        env=environment,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
        creationflags=creationflags,
    )
    state = RunState(process=process, output_dir=output_dir)
    threading.Thread(target=reader_worker, args=(state,), daemon=True).start()
    return state


def launch_summarizer(
    input_dir: Path,
    provider: str,
    model: str,
    api_key: str,
    base_url: str,
    limit: int,
    dry_run: bool,
    reasoning_effort: str,
    max_output_tokens: int,
    thinking_mode: str = "enabled",
    temperature: float = 0.2,
    validation_retries: int = 1,
) -> RunState:
    output_dir = new_summary_output_dir()
    command = [
        sys.executable,
        "-X",
        "utf8",
        "-u",
        str(SUMMARIZER),
        str(input_dir),
        "--output-dir",
        str(output_dir),
        "--provider",
        provider,
        "--limit",
        str(max(0, limit)),
        "--base-url",
        base_url.strip(),
        "--max-output-tokens",
        str(max(1, max_output_tokens)),
        "--validation-retries",
        str(max(0, validation_retries)),
    ]
    if model.strip():
        command.extend(["--model", model.strip()])
    if reasoning_effort:
        command.extend(["--reasoning-effort", reasoning_effort])
    if provider == "deepseek":
        command.extend(["--thinking-mode", thinking_mode])
        command.extend(["--temperature", str(min(2.0, max(0.0, temperature)))])
    if dry_run:
        command.append("--dry-run")

    environment = os.environ.copy()
    environment["PYTHONUTF8"] = "1"
    environment.pop("OPENAI_API_KEY", None)
    environment.pop("DEEPSEEK_API_KEY", None)
    if api_key:
        environment["DEEPSEEK_API_KEY" if provider == "deepseek" else "OPENAI_API_KEY"] = api_key

    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0
    process = subprocess.Popen(
        command,
        cwd=PROJECT_DIR,
        env=environment,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
        creationflags=creationflags,
    )
    state = RunState(process=process, output_dir=output_dir)
    threading.Thread(target=reader_worker, args=(state,), daemon=True).start()
    return state


def drain_events(state: RunState) -> None:
    while True:
        try:
            event, value = state.events.get_nowait()
        except queue.Empty:
            break
        if event == "log":
            state.logs.append(str(value))
            if len(state.logs) > 3000:
                state.logs = state.logs[-3000:]
        elif event == "done":
            state.return_code = int(value)
        elif event == "error":
            state.error = str(value)
            state.return_code = -1


def build_result_zip(output_dir: Path) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(output_dir.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(output_dir))
    return buffer.getvalue()


def render_downloads(output_dir: Path) -> None:
    st.markdown("#### 结果文件")
    columns = st.columns(4)
    files = [
        ("证据包", "llm_evidence_pack.md", "text/markdown"),
        ("证据表", "evidence_candidates.csv", "text/csv"),
        ("Token 估计", "llm_token_estimates.json", "application/json"),
        ("运行元数据", "run_metadata.json", "application/json"),
    ]
    for column, (label, filename, mime) in zip(columns, files):
        path = output_dir / filename
        with column:
            if path.exists():
                st.download_button(
                    label,
                    path.read_bytes(),
                    file_name=filename,
                    mime=mime,
                    use_container_width=True,
                )
            else:
                st.button(label, disabled=True, use_container_width=True)
    st.download_button(
        "下载完整结果 ZIP（含逐产品原始记录）",
        build_result_zip(output_dir),
        file_name=f"{output_dir.name}.zip",
        mime="application/zip",
        use_container_width=True,
    )

    metadata_path = output_dir / "run_metadata.json"
    token_path = output_dir / "llm_token_estimates.json"
    metric_columns = st.columns(7)
    if metadata_path.exists():
        metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
        metric_columns[0].metric("产品数", metadata.get("product_count", 0))
        metric_columns[1].metric("数据源错误", metadata.get("source_error_count", 0))
        metric_columns[2].metric("Tavily 本次 credits", metadata.get("tavily_credits_this_run", 0))
        metric_columns[3].metric("Tavily 缓存命中", metadata.get("tavily_cache_hits", 0))
        metric_columns[4].metric("候选证据", metadata.get("candidate_evidence_items", 0))
        metric_columns[5].metric("证据包入选", metadata.get("selected_pack_items", 0))
        metric_columns[6].metric("MCE credits", metadata.get("mce_tavily_credits_this_run", 0))
        patent_columns = st.columns(4)
        patent_columns[0].metric("专利核验模式", metadata.get("patent_verification_mode", "off"))
        patent_columns[1].metric("核验入选专利", metadata.get("patent_verified_items", 0))
        patent_columns[2].metric("专利本次联网请求", metadata.get("patent_network_requests_this_run", 0))
        patent_columns[3].metric("专利缓存命中", metadata.get("patent_cache_hits", 0))
        mce_skipped = int(metadata.get("mce_products_skipped_by_optional_cap", 0))
        if mce_skipped:
            st.warning(
                f"MCE 可选软上限已生效：{mce_skipped} 个产品未发起 MCE 查询，"
                "详细状态已记入逐产品原始记录。"
            )
    if token_path.exists():
        token_data = json.loads(token_path.read_text(encoding="utf-8"))
        token_count = int(token_data.get("estimated_total_tokens", 0))
        st.info(f"证据包粗略输入 Token：{token_count:,}（字符法估计，不是账户实际扣量）")


def render_summary_downloads(output_dir: Path) -> None:
    st.caption(
        "总结 CSV/JSONL/报告会自动展开 E 编号：显示证据标题、来源、URL，"
        "并单独列出已核验专利号与未经正文核验的候选专利号。"
    )
    files = [
        ("总结报告", "summary_report.md", "text/markdown"),
        ("总结 CSV", "summary_results.csv", "text/csv"),
        ("逐产品 Token", "summary_token_usage.json", "application/json"),
        ("运行元数据", "summary_run_metadata.json", "application/json"),
        ("提示词预算", "prompt_manifest.json", "application/json"),
    ]
    columns = st.columns(len(files))
    for column, (label, filename, mime) in zip(columns, files):
        path = output_dir / filename
        with column:
            if path.exists():
                st.download_button(
                    label,
                    path.read_bytes(),
                    file_name=filename,
                    mime=mime,
                    use_container_width=True,
                )
            else:
                st.button(label, disabled=True, use_container_width=True)
    st.download_button(
        "下载总结结果 ZIP",
        build_result_zip(output_dir),
        file_name=f"{output_dir.name}.zip",
        mime="application/zip",
        use_container_width=True,
    )

    usage_path = output_dir / "summary_token_usage.json"
    manifest_path = output_dir / "prompt_manifest.json"
    if usage_path.exists():
        usage = json.loads(usage_path.read_text(encoding="utf-8"))
        st.info(
            "API 返回的实际 Token："
            f"输入 {int(usage.get('total_input_tokens', 0)):,}，"
            f"输出 {int(usage.get('total_output_tokens', 0)):,}，"
            f"总计 {int(usage.get('total_tokens', 0)):,}"
        )
    elif manifest_path.exists():
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        st.info(
            "离线估算模型输入 Token："
            f"{int(manifest.get('estimated_total_input_tokens', 0)):,}（尚未调用 API）"
        )


@st.fragment(run_every=1.0)
def run_monitor() -> None:
    state: RunState | None = st.session_state.get("run_state")
    if state is None:
        st.caption("尚未开始运行。建议先用“测试 1 个产品”验证密钥和网络。")
        return
    drain_events(state)
    running = state.return_code is None
    status_column, action_column = st.columns([4, 1])
    with status_column:
        if running:
            st.info(f"正在采集……输出目录：`{state.output_dir}`")
        elif state.return_code == 0:
            st.success(f"采集完成。输出目录：`{state.output_dir}`")
        elif state.stopped:
            st.warning(
                f"任务已停止（退出代码 {state.return_code}）。"
                "已完成分子的产品检查点和 HTTP 缓存均会保留，可稍后续跑。"
            )
        else:
            st.error(f"任务结束，退出代码 {state.return_code}。{state.error}")
    with action_column:
        if running:
            if st.button("停止任务", type="secondary", use_container_width=True):
                state.stopped = True
                state.process.terminate()
                st.rerun(scope="fragment")
        else:
            if state.output_dir.exists() and st.button("在资源管理器中打开", use_container_width=True):
                try:
                    os.startfile(str(state.output_dir))  # type: ignore[attr-defined]
                except Exception as exc:
                    st.error(f"无法打开目录：{exc}")
            if st.button("准备下一次运行", use_container_width=True):
                st.session_state.run_state = None
                st.rerun()
    log_text = "\n".join(state.logs) or "等待脚本输出……"
    st.code(log_text, language="text", line_numbers=False)
    resume_path = state.output_dir / "resume_state.json"
    if resume_path.exists():
        try:
            resume_data = json.loads(resume_path.read_text(encoding="utf-8"))
            completed = int(resume_data.get("completed_products") or 0)
            total = int(resume_data.get("total_products") or 0)
            st.progress(
                completed / total if total else 0.0,
                text=f"产品检查点：{completed}/{total}；剩余 {max(0, total - completed)}",
            )
        except (OSError, json.JSONDecodeError, TypeError, ValueError):
            pass
    if not running and state.output_dir.exists():
        render_downloads(state.output_dir)


@st.fragment(run_every=1.0)
def summary_monitor() -> None:
    state: RunState | None = st.session_state.get("summary_state")
    if state is None:
        st.caption("尚未开始总结。建议先使用“离线测试 1 个产品”。")
        return
    drain_events(state)
    running = state.return_code is None
    status_column, action_column = st.columns([4, 1])
    with status_column:
        if running:
            st.info(f"正在运行总结模块……输出目录：`{state.output_dir}`")
        elif state.return_code == 0:
            st.success(f"总结模块完成。输出目录：`{state.output_dir}`")
        elif state.stopped:
            st.warning(f"总结任务已停止（退出代码 {state.return_code}）。")
        else:
            st.error(f"总结任务结束，退出代码 {state.return_code}。{state.error}")
    with action_column:
        if running:
            if st.button("停止总结", type="secondary", use_container_width=True):
                state.stopped = True
                state.process.terminate()
                st.rerun(scope="fragment")
        else:
            if state.output_dir.exists() and st.button(
                "打开总结目录", use_container_width=True
            ):
                try:
                    os.startfile(str(state.output_dir))  # type: ignore[attr-defined]
                except Exception as exc:
                    st.error(f"无法打开目录：{exc}")
            if st.button("准备下一次总结", use_container_width=True):
                st.session_state.summary_state = None
                st.rerun()
    st.code("\n".join(state.logs) or "等待脚本输出……", language="text")
    if not running and state.output_dir.exists():
        render_summary_downloads(state.output_dir)


def main() -> None:
    st.set_page_config(page_title="询单产品证据工作台", page_icon="🧪", layout="wide")
    st.title("询单产品证据工作台")
    st.caption("脚本收集资料 · 大模型逐产品总结 · 两个阶段独立运行和计量")
    if "_saved_gui_settings" not in st.session_state:
        st.session_state._saved_gui_settings = load_gui_settings()
    saved_settings: dict[str, object] = st.session_state._saved_gui_settings

    if not COLLECTOR.exists():
        st.error(f"未找到采集脚本：{COLLECTOR}")
        st.stop()

    with st.sidebar:
        st.header("运行设置")
        max_workers = st.number_input(
            "产品并发数", min_value=1, max_value=16,
            value=int(persisted_value(saved_settings, "max_workers", 4)), step=1,
            key="max_workers",
        )
        web_results = st.number_input(
            "每个网页查询返回结果", min_value=1, max_value=20,
            value=int(persisted_value(saved_settings, "web_results", 8)), step=1,
            key="web_results",
        )
        pack_items = st.number_input(
            "每个产品写入证据包的条目", min_value=1, max_value=50,
            value=int(persisted_value(saved_settings, "pack_items", 14)), step=1,
            key="pack_items",
        )
        patent_mode_options = ["balanced", "deep", "off"]
        saved_patent_mode = str(persisted_value(saved_settings, "patent_verification", "balanced"))
        patent_verification = st.selectbox(
            "专利候选正文二次核验",
            options=patent_mode_options,
            index=patent_mode_options.index(saved_patent_mode) if saved_patent_mode in patent_mode_options else 0,
            format_func=lambda value: {
                "balanced": "平衡（最多读 5 篇，入选 3 篇）",
                "deep": "深度（最多读 10 篇，入选 5 篇）",
                "off": "关闭（只保留专利索引候选）",
            }[value],
            key="patent_verification",
            help="先按专利号、年份和精确检索结果用本地代码筛选，再读取少量正文；同族专利合并，找到 2 条高置信直接证据后提前停止。",
        )
        mce_max_products = st.number_input(
            "MCE 可选软上限（0 = 处理全部）",
            min_value=0,
            max_value=100000,
            value=int(persisted_value(saved_settings, "mce_max_products", 0)),
            step=50,
            key="mce_max_products",
            help="默认 0：不截断，200 或更多产品会自动排队处理。如只想试跑前 50 个，可主动填写 50。",
        )
        offline = st.checkbox(
            "离线模式（只使用缓存）",
            value=bool(persisted_value(saved_settings, "offline", False)),
            key="offline",
        )
        include_recent_events = st.checkbox(
            "额外收集名称相关近期事件",
            value=bool(persisted_value(saved_settings, "include_recent_events", False)),
            help="默认关闭。启用后只作为间接候选归档，不进入默认大模型证据包。",
            key="include_recent_events",
        )
        st.caption("Tavily 固定使用 basic，关闭生成式回答；默认只执行精确 CAS 与 CAS+用途查询。")
        if patent_verification == "balanced":
            st.caption("专利平衡模式：Google Patents 最多 1 次候选检索 + 5 页正文；最终最多 3 篇、每篇 1200 字符进入模型证据包。")
        elif patent_verification == "deep":
            st.caption("专利深度模式：最多 3 次候选检索 + 10 页正文；最终最多 5 篇。适合疑难单品，不建议无差别用于大批量任务。")
        else:
            st.caption("专利正文核验已关闭；PubChem 专利号只作为候选，不会据此认定用途。")
        st.caption(f"参数会自动保存到 `{SETTINGS_FILE}`。")

    input_column, key_column = st.columns([3, 2])
    with input_column:
        st.subheader("1. 选择输入方式")
        input_modes = ["file", "single_cas"]
        saved_input_mode = str(persisted_value(saved_settings, "input_mode", "file"))
        input_mode = st.radio(
            "询单类型",
            options=input_modes,
            index=input_modes.index(saved_input_mode) if saved_input_mode in input_modes else 0,
            format_func=lambda value: "Excel/CSV 批量询单" if value == "file" else "突发单个 CAS",
            horizontal=True,
            key="input_mode",
        )
        input_path: Path | None = None
        single_cas = ""
        if input_mode == "single_cas":
            single_cas = st.text_input(
                "CAS 号",
                value=str(persisted_value(saved_settings, "single_cas", "")),
                placeholder="例如 50-78-2",
                key="single_cas",
            ).strip()
            if single_cas and not cas_is_valid(single_cas):
                st.error("CAS 格式或校验位不正确。")
            elif single_cas:
                st.success(f"将直接查询 CAS {single_cas}，无需创建 Excel。")
        else:
            use_example = st.checkbox(
                "使用工作区中的“询单范例.xlsx”",
                value=bool(persisted_value(saved_settings, "use_example", EXAMPLE_FILE.exists())),
                disabled=not EXAMPLE_FILE.exists(),
                key="use_example",
            )
            uploaded = st.file_uploader(
                "上传 Excel 或 CSV",
                type=["xlsx", "xlsm", "csv"],
                disabled=use_example,
            )
            previous_path_text = str(persisted_value(saved_settings, "last_input_path", ""))
            previous_path = Path(previous_path_text) if previous_path_text else None
            input_path = EXAMPLE_FILE if use_example else None
            if uploaded is not None and not use_example:
                try:
                    input_path = save_upload(uploaded)
                    st.session_state.last_input_path = str(input_path)
                    st.caption(f"已暂存：`{input_path}`")
                except ValueError as exc:
                    st.error(str(exc))
            elif not use_example and previous_path and previous_path.is_file():
                input_path = previous_path
                st.caption(f"已恢复上次输入文件：`{input_path}`")

    with key_column:
        st.subheader("2. API 密钥")
        tavily_key = st.text_input(
            "Tavily API Key",
            value=str(persisted_value(saved_settings, "tavily_key", os.getenv("TAVILY_API_KEY", ""))),
            type="password",
            key="tavily_key",
        )
        brave_key = st.text_input(
            "Brave API Key（备用，可留空）",
            value=str(persisted_value(saved_settings, "brave_key", os.getenv("BRAVE_SEARCH_API_KEY", ""))),
            type="password",
            key="brave_key",
        )
        st.warning("注意：该联网搜索功能可能需要付费，请以 Tavily/Brave 账户套餐和实际 credits 为准。")
        st.caption("密钥会明文保存在本机 GUI 设置文件中；不写入命令行或任务输出。")

    st.subheader("3. 选择数据源")
    saved_sources = persisted_value(saved_settings, "selected_sources", DEFAULT_SOURCES)
    if not isinstance(saved_sources, list):
        saved_sources = DEFAULT_SOURCES
    saved_sources = [name for name in saved_sources if name in SOURCE_OPTIONS]
    selected_sources = st.multiselect(
        "数据源",
        options=list(SOURCE_OPTIONS),
        default=saved_sources or DEFAULT_SOURCES,
        format_func=lambda name: SOURCE_OPTIONS[name],
        label_visibility="collapsed",
        key="selected_sources",
    )

    current_state: RunState | None = st.session_state.get("run_state")
    is_running = current_state is not None and current_state.return_code is None
    input_valid = cas_is_valid(single_cas) if input_mode == "single_cas" else bool(input_path)
    validation_error = not input_valid or not selected_sources
    test_clicked = False
    full_clicked = False
    resume_clicked = False
    selected_resume_dir: Path | None = None
    if input_mode == "single_cas":
        query_column, note_column = st.columns([2, 5])
        with query_column:
            full_clicked = st.button(
                "查询此 CAS",
                type="primary",
                disabled=is_running or validation_error,
                use_container_width=True,
            )
    else:
        test_column, full_column, note_column = st.columns([1, 1, 4])
        with test_column:
            test_clicked = st.button(
                "测试 1 个产品",
                type="secondary",
                disabled=is_running or validation_error,
                use_container_width=True,
            )
        with full_column:
            full_clicked = st.button(
                "处理全部产品",
                type="primary",
                disabled=is_running or validation_error,
                use_container_width=True,
            )
    with note_column:
        if ({"tavily", "mce"} & set(selected_sources)) and not tavily_key and not offline:
            st.warning("已选择 Tavily 或 MCE，但未填写 Tavily Key；相应数据源会被跳过。")
        elif validation_error:
            st.caption("请先提供有效输入，并选择至少一个数据源。")
        else:
            st.caption("测试和正式运行都会生成独立的时间戳输出目录，不覆盖历史结果。")

    resumable_runs = resumable_evidence_output_dirs()
    if resumable_runs:
        with st.expander(f"断点续跑（发现 {len(resumable_runs)} 个未完成任务）", expanded=True):
            resume_options = [str(path) for path, _state in resumable_runs]
            saved_resume = str(persisted_value(saved_settings, "resume_selected_dir", ""))
            resume_index = resume_options.index(saved_resume) if saved_resume in resume_options else 0
            selected_resume_text = st.selectbox(
                "选择原输出目录",
                options=resume_options,
                index=resume_index,
                key="resume_selected_dir",
            )
            selected_resume_dir = Path(selected_resume_text)
            selected_state = next(
                state for path, state in resumable_runs if path == selected_resume_dir
            )
            completed_count = int(selected_state.get("completed_products") or 0)
            total_count = int(selected_state.get("total_products") or 0)
            st.info(
                f"已完成 {completed_count}/{total_count}，剩余 {max(0, total_count - completed_count)}。"
                "续跑会原地使用原产品列表、数据源和采集参数，"
                "直接跳过已完成分子；API Key 使用当前界面中的值。"
            )
            resume_clicked = st.button(
                "从检查点继续",
                type="primary",
                disabled=is_running,
                use_container_width=True,
                key="resume_collection",
            )

    if test_clicked or full_clicked or resume_clicked:
        try:
            persist_current_session(
                last_input_path=str(input_path) if input_path and input_mode == "file" else str(
                    persisted_value(saved_settings, "last_input_path", "")
                )
            )
            collector_state = launch_collector(
                input_path=input_path,
                single_cas=single_cas if input_mode == "single_cas" else "",
                sources=selected_sources,
                tavily_key=tavily_key.strip(),
                brave_key=brave_key.strip(),
                limit=1 if test_clicked else 0,
                max_workers=int(max_workers),
                web_results=int(web_results),
                pack_items=int(pack_items),
                patent_verification=patent_verification,
                mce_max_products=int(mce_max_products),
                offline=offline,
                include_recent_events=include_recent_events,
                resume_dir=selected_resume_dir if resume_clicked else None,
            )
            st.session_state.run_state = collector_state
            st.session_state.summary_input_text = str(collector_state.output_dir)
            persist_current_session()
            st.rerun()
        except Exception as exc:
            st.error(f"无法启动采集脚本：{type(exc).__name__}: {exc}")

    st.divider()
    st.subheader("运行状态")
    run_monitor()

    st.divider()
    st.header("大模型总结（测试版）")
    st.caption(
        "只读取已采集的 llm_evidence_pack.md；不会启用模型网页搜索，"
        "不会重新消耗 Tavily 额度。"
    )
    if not SUMMARIZER.exists():
        st.error(f"未找到总结脚本：{SUMMARIZER}")
        st.stop()

    latest_dir = latest_evidence_output_dir()
    summary_input_text = st.text_input(
        "证据采集输出目录",
        value=str(persisted_value(saved_settings, "summary_input_text", str(latest_dir or ""))),
        help="目录中必须包含 llm_evidence_pack.md。",
        key="summary_input_text",
    )
    summary_input_dir = Path(summary_input_text.strip()) if summary_input_text.strip() else None
    summary_input_valid = bool(
        summary_input_dir
        and summary_input_dir.is_dir()
        and (summary_input_dir / "llm_evidence_pack.md").exists()
    )

    provider_options = ["openai", "deepseek"]
    saved_provider = str(persisted_value(saved_settings, "summary_provider", "openai"))
    summary_provider = st.radio(
        "模型 API 提供商",
        options=provider_options,
        index=provider_options.index(saved_provider) if saved_provider in provider_options else 0,
        format_func=lambda value: "OpenAI" if value == "openai" else "DeepSeek",
        horizontal=True,
        key="summary_provider",
    )
    summary_dry_run = st.checkbox(
        "仅导出提示词（不联网、不调用模型）",
        value=bool(persisted_value(saved_settings, "summary_dry_run", True)),
        key="summary_dry_run",
    )
    st.warning(
        "注意：真实 OpenAI/DeepSeek API 总结功能需要付费，"
        "会按提供商规则消耗 Token/账户余额；“仅导出提示词”不调用模型。"
    )

    summary_settings, summary_credentials = st.columns(2)
    thinking_mode = "enabled"
    temperature = 0.2
    if summary_provider == "openai":
        with summary_settings:
            openai_models = ["gpt-5.6-luna", "gpt-5.6-terra", "gpt-5.6-sol", "custom"]
            saved_openai_model = str(persisted_value(saved_settings, "openai_model_choice", "gpt-5.6-luna"))
            openai_model_choice = st.selectbox(
                "OpenAI 模型",
                options=openai_models,
                index=openai_models.index(saved_openai_model) if saved_openai_model in openai_models else 0,
                format_func=lambda value: {
                    "gpt-5.6-luna": "gpt-5.6-luna（高吞吐/节省）",
                    "gpt-5.6-terra": "gpt-5.6-terra（平衡）",
                    "gpt-5.6-sol": "gpt-5.6-sol（质量优先）",
                    "custom": "自定义模型名",
                }[value],
                key="openai_model_choice",
            )
            summary_model = (
                st.text_input(
                    "自定义 OpenAI 模型名",
                    value=str(persisted_value(saved_settings, "openai_custom_model", "")),
                    key="openai_custom_model",
                )
                if openai_model_choice == "custom"
                else openai_model_choice
            )
            openai_efforts = ["none", "low", "medium", "high", "xhigh", "max"]
            saved_openai_effort = str(persisted_value(saved_settings, "openai_reasoning_effort", "medium"))
            reasoning_effort = st.selectbox(
                "OpenAI reasoning.effort",
                options=openai_efforts,
                index=openai_efforts.index(saved_openai_effort) if saved_openai_effort in openai_efforts else 2,
                format_func=lambda value: {
                    "none": "none（最快）",
                    "low": "low",
                    "medium": "medium（建议起点）",
                    "high": "high",
                    "xhigh": "xhigh",
                    "max": "max（最慢/最高推理量）",
                }[value],
                key="openai_reasoning_effort",
            )
        with summary_credentials:
            provider_api_key = st.text_input(
                "OpenAI API Key",
                value=str(persisted_value(saved_settings, "openai_summary_key", os.getenv("OPENAI_API_KEY", ""))),
                type="password",
                disabled=summary_dry_run,
                key="openai_summary_key",
            )
            api_base_url = st.text_input(
                "OpenAI Responses API 根地址",
                value=str(persisted_value(saved_settings, "openai_base_url", "https://api.openai.com/v1")),
                key="openai_base_url",
            )
            st.caption(
                "调用 POST /responses；使用 strict json_schema，store=false，不配置网页搜索工具。"
            )
    else:
        with summary_settings:
            deepseek_models = ["deepseek-v4-flash", "deepseek-v4-pro", "custom"]
            saved_deepseek_model = str(persisted_value(saved_settings, "deepseek_model_choice", "deepseek-v4-flash"))
            deepseek_model_choice = st.selectbox(
                "DeepSeek 模型",
                options=deepseek_models,
                index=deepseek_models.index(saved_deepseek_model) if saved_deepseek_model in deepseek_models else 0,
                format_func=lambda value: {
                    "deepseek-v4-flash": "deepseek-v4-flash（高吞吐/节省）",
                    "deepseek-v4-pro": "deepseek-v4-pro（质量优先）",
                    "custom": "自定义模型名",
                }[value],
                key="deepseek_model_choice",
            )
            summary_model = (
                st.text_input(
                    "自定义 DeepSeek 模型名",
                    value=str(persisted_value(saved_settings, "deepseek_custom_model", "")),
                    key="deepseek_custom_model",
                )
                if deepseek_model_choice == "custom"
                else deepseek_model_choice
            )
            thinking_options = ["enabled", "disabled"]
            saved_thinking = str(persisted_value(saved_settings, "deepseek_thinking_mode", "enabled"))
            thinking_mode = st.radio(
                "DeepSeek thinking.type",
                options=thinking_options,
                index=thinking_options.index(saved_thinking) if saved_thinking in thinking_options else 0,
                format_func=lambda value: "enabled（启用思考）" if value == "enabled" else "disabled（不思考）",
                horizontal=True,
                key="deepseek_thinking_mode",
            )
            deepseek_efforts = ["high", "max"]
            saved_deepseek_effort = str(persisted_value(saved_settings, "deepseek_reasoning_effort", "high"))
            reasoning_effort = (
                st.selectbox(
                    "DeepSeek reasoning_effort",
                    options=deepseek_efforts,
                    index=deepseek_efforts.index(saved_deepseek_effort) if saved_deepseek_effort in deepseek_efforts else 0,
                    key="deepseek_reasoning_effort",
                )
                if thinking_mode == "enabled"
                else ""
            )
            temperature = st.slider(
                "DeepSeek temperature",
                min_value=0.0,
                max_value=2.0,
                value=float(persisted_value(saved_settings, "deepseek_temperature", 0.2)),
                step=0.1,
                help="询单结构化总结建议使用较低随机性。",
                key="deepseek_temperature",
            )
        with summary_credentials:
            provider_api_key = st.text_input(
                "DeepSeek API Key",
                value=str(persisted_value(saved_settings, "deepseek_summary_key", os.getenv("DEEPSEEK_API_KEY", ""))),
                type="password",
                disabled=summary_dry_run,
                key="deepseek_summary_key",
            )
            api_base_url = st.text_input(
                "DeepSeek API 根地址",
                value=str(persisted_value(saved_settings, "deepseek_base_url", "https://api.deepseek.com")),
                key="deepseek_base_url",
            )
            st.caption(
                "调用 POST /chat/completions；使用 response_format=json_object，再在本地执行业务 Schema 校验。"
            )

    common_settings, security_note = st.columns(2)
    with common_settings:
        max_output_tokens = st.number_input(
            "每产品最大输出 tokens",
            min_value=400,
            max_value=8000,
            value=int(persisted_value(saved_settings, "max_output_tokens", 1400)),
            step=100,
            key="max_output_tokens",
        )
        validation_retries = st.number_input(
            "JSON/业务 Schema 校验失败后重试次数",
            min_value=0,
            max_value=3,
            value=int(persisted_value(saved_settings, "validation_retries", 1)),
            step=1,
            help="重试会发起新的模型请求，所有尝试的 token 都归入该产品。",
            key="validation_retries",
        )
    with security_note:
        st.caption(
            "API Key 会明文保存在本机 GUI 设置文件中，但不写入命令行、提示词、原始响应或任务元数据。"
        )

    summary_state: RunState | None = st.session_state.get("summary_state")
    summary_running = summary_state is not None and summary_state.return_code is None
    live_missing = not summary_dry_run and (
        not provider_api_key.strip() or not summary_model.strip()
    )
    summary_disabled = summary_running or not summary_input_valid or live_missing
    one_column, all_column, summary_note = st.columns([1, 1, 4])
    with one_column:
        summary_one_clicked = st.button(
            "测试 1 个产品",
            type="secondary",
            disabled=summary_disabled,
            use_container_width=True,
            key="summary_test_one",
        )
    with all_column:
        summary_all_clicked = st.button(
            "总结全部产品",
            type="primary",
            disabled=summary_disabled,
            use_container_width=True,
            key="summary_all",
        )
    with summary_note:
        if not summary_input_valid:
            st.warning("请选择包含 llm_evidence_pack.md 的有效目录。")
        elif live_missing:
            st.warning(f"真实 {summary_provider} 调用需要填写 API Key 和模型名称。")
        elif summary_dry_run:
            st.info("当前为安全离线模式，只生成请求文件和 token 估算。")
        else:
            st.warning(
                f"将按产品逐次调用 {summary_provider} API；"
                "总结器不会使用网页搜索工具。"
            )

    if summary_one_clicked or summary_all_clicked:
        try:
            persist_current_session()
            st.session_state.summary_state = launch_summarizer(
                input_dir=summary_input_dir,  # type: ignore[arg-type]
                provider=summary_provider,
                model=summary_model,
                api_key=provider_api_key.strip(),
                base_url=api_base_url,
                limit=1 if summary_one_clicked else 0,
                dry_run=summary_dry_run,
                reasoning_effort=reasoning_effort,
                max_output_tokens=int(max_output_tokens),
                thinking_mode=thinking_mode,
                temperature=float(temperature),
                validation_retries=int(validation_retries),
            )
            st.rerun()
        except Exception as exc:
            st.error(f"无法启动总结脚本：{type(exc).__name__}: {exc}")

    st.subheader("总结运行状态")
    summary_monitor()
    try:
        persist_current_session()
    except OSError as exc:
        st.warning(f"无法保存 GUI 参数：{exc}")


if __name__ == "__main__":
    main()
