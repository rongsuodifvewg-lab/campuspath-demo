# -*- coding: utf-8 -*-
"""
CampusPath｜大学生个人进化决策 OS
Personal Ontology + AI Decision Copilot + Action OS

运行：
python -m streamlit run app.py --server.address 127.0.0.1 --server.port 8501
"""

from __future__ import annotations

import io
import json
import os
import re
import sqlite3
from datetime import datetime
from html import escape
from pathlib import Path
from typing import Any, Dict, List, Tuple

import pandas as pd
import requests
import streamlit as st

try:
    from dotenv import load_dotenv
    load_dotenv(override=True)
except Exception:
    pass

try:
    from bs4 import BeautifulSoup
except Exception:
    BeautifulSoup = None

try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

try:
    from docx import Document
except Exception:
    Document = None


# =========================================================
# 0. 全局配置
# =========================================================

APP_NAME = "CampusPath"
APP_SUBTITLE = "基于 QQ / 微信课程群的大学生个人进化决策 OS"

ROOT = Path(__file__).resolve().parent

# 用新数据库名，避免旧表结构冲突
DB_PATH = ROOT / "campuspath_showcase.db"

MODEL_API_URL = os.getenv("MODEL_API_URL", "https://tokenhub.tencentmaas.com/v1/chat/completions")
MODEL_API_KEY = os.getenv("MODEL_API_KEY", "")
MODEL_NAME = os.getenv("MODEL_NAME", "hy3-preview")

AI_TIMEOUT_SEC = int(os.getenv("AI_TIMEOUT_SEC", "35"))
AI_MAX_TOKENS = int(os.getenv("AI_MAX_TOKENS", "1200"))
AI_REPORT_WORDS = int(os.getenv("AI_REPORT_WORDS", "900"))

# 默认不走系统代理，避免 ProxyError
USE_SYSTEM_PROXY = os.getenv("USE_SYSTEM_PROXY", "0") == "1"

PATHS = ["保研", "考研", "就业", "出国"]
STATUS_OPTIONS = ["todo", "doing", "done", "blocked", "dropped"]

PAGES = [
    "决策驾驶舱",
    "数据采集",
    "个人本体",
    "AI路径决策",
    "Action OS",
    "Outcome反馈",
    "数据资产",
    "提交包",
]

PCG_BOUNDARY = (
    "当前 Demo 不读取真实 QQ / 微信数据库，不自动爬取小红书；"
    "数据来自用户授权粘贴、文件上传、官方网页 URL 导入与本地结构化知识库。"
    "院校政策最终以学校官网、学院官网和研招网最新通知为准。"
)

SKILL_TERMS = [
    "Python", "SQL", "Streamlit", "AIGC", "LLM", "RAG", "Prompt",
    "数据分析", "产品设计", "用户研究", "PRD", "竞品分析",
    "机器学习", "深度学习", "NLP", "爬虫", "前端", "后端",
    "数据库", "算法", "操作系统", "计算机网络", "Linux", "Git",
    "Excel", "Tableau", "PowerBI", "项目管理", "原型", "Figma", "Axure"
]


# =========================================================
# 1. Demo 默认数据
# =========================================================

DEFAULT_RESUME = """张同学｜计算机科学与技术｜大三
学校层次：双非一本
成绩：GPA 3.62 / 4.0
专业排名：前 18%
英语：CET-6 520

项目经历：
1. 校园二手交易平台项目
- 使用 Python / Streamlit 搭建二手商品发布、搜索、收藏、详情页和联系卖家功能
- 负责需求拆解、页面原型、核心流程开发和演示材料整理
- 项目目前完成 Demo，可以展示商品列表、商品详情和发布表单

科研与实习：
- 暂无正式科研论文
- 暂无正式实习

当前困惑：
- 纠结保研、考研、就业三条路径
- 不知道应该先补科研、补实习，还是先确定考研院校
"""

DEFAULT_CHAT = """[辅导员] 本周五 23:59 前完成机器学习第 3 次作业，提交到学习通，逾期不补。
[助教] 数据库课程实验报告 DDL 是周日晚上 22:00，请大家上传到腾讯文档收集表：https://docs.qq.com/demo-campuspath-homework
[班长] 操作系统期中考试安排在下周三 14:00-16:00，范围是进程、线程、调度和内存管理。
[老师] 今天的机器学习 PPT、PDF 课件和录播已经上传 QQ 群文件。
[辅导员] 通知：明天下午课程调到 3 教 205，上课前请完成签到。
[学院教务] 周四晚上 19:00 有保研经验讲座，主题是计算机学院夏令营材料准备和导师联系。
[就业委员] 腾讯 PCG AI 产品实习宣讲会本周五 19:30 线上举行，报名链接：https://docs.qq.com/demo-campuspath-signup
[班委] 微信群同步提醒：想报名腾讯宣讲的同学今晚前填腾讯文档。
[学姐] 考研择校表建议把考试科目、复试比例、城市、往年分数都列出来。
"""

DEFAULT_POLICY = """本校推免细则 Demo：
1. 推免资格原则上要求综合排名前 20%，必修课无不及格记录。
2. 综合成绩由课程成绩、科研竞赛、社会实践、英语能力等组成。
3. 科研竞赛加分需要证书、论文录用证明、专利证明或学院认定材料。
4. 申请材料包括成绩单、专业排名证明、英语成绩、个人陈述、推荐信、获奖证明、项目证明。
5. 学院通常在 9 月前后进行资格审核、公示和综合面试。
6. 最终推免资格、加分项目、名额分配和时间节点以学校和学院官网最新通知为准。
"""

DEFAULT_JD = """腾讯 PCG AI 产品策划实习生 Demo JD：
岗位方向：AI 产品、内容生态、社区产品、工具产品。
工作内容：
1. 参与 AI 产品需求分析、用户场景拆解和原型设计；
2. 分析校园用户在学习、求职、内容创作中的核心痛点；
3. 与算法、研发、设计协作，推进功能从 Demo 到上线；
4. 通过数据分析和用户访谈评估功能效果。
要求：
- 熟悉 AIGC / 大模型 / AI Copilot 产品形态；
- 具备产品思维、结构化表达和基础数据分析能力；
- 有 Python、SQL、原型工具、校园项目或互联网实习经历加分；
- 能输出 PRD、竞品分析、用户故事、Demo 演示材料。
"""

ADMISSIONS = [
    {
        "school": "清华大学",
        "program": "软件学院｜电子信息/软件工程",
        "level": "冲刺",
        "type": "保研/夏令营",
        "match": "高挑战目标，适合作为冲刺，不建议作为唯一选择。",
        "material": "成绩单、排名证明、简历、个人陈述、推荐信、项目证明。",
        "risk": "需要较强科研或项目深度，最终以学院官网通知为准。",
        "url": "https://www.sigs.tsinghua.edu.cn/",
    },
    {
        "school": "南京大学",
        "program": "软件学院｜软件工程",
        "level": "稳妥",
        "type": "预推免",
        "match": "适合有软件项目基础的计算机学生。",
        "material": "成绩单、排名证明、简历、个人陈述、项目材料。",
        "risk": "需要把课程项目包装成可讲清楚的项目证据。",
        "url": "https://software.nju.edu.cn/",
    },
    {
        "school": "北京交通大学",
        "program": "计算机与信息技术学院｜计算机技术",
        "level": "稳妥",
        "type": "预推免",
        "match": "材料完整度和专业基础更关键。",
        "material": "成绩单、排名证明、简历、英语证明。",
        "risk": "以学院当年通知为准，需提前核验时间线。",
        "url": "https://cs.bjtu.edu.cn/",
    },
    {
        "school": "苏州大学",
        "program": "计算机科学与技术学院｜计算机技术",
        "level": "保底",
        "type": "预推免",
        "match": "适合作为保底偏稳妥项目。",
        "material": "成绩单、排名证明、简历、项目说明。",
        "risk": "政策时间可能变化，需关注学院公告。",
        "url": "https://yjs.suda.edu.cn/",
    },
    {
        "school": "电子科技大学",
        "program": "计算机科学与工程学院｜计算机科学与技术",
        "level": "冲刺",
        "type": "考研",
        "match": "适合考研冲刺，但需要尽早确认专业课。",
        "material": "初试成绩、复试材料、简历、项目证明。",
        "risk": "需要核验考试科目、复试比例和近年分数线。",
        "url": "https://yz.chsi.com.cn/",
    },
    {
        "school": "杭州电子科技大学",
        "program": "计算机学院｜计算机技术",
        "level": "保底",
        "type": "考研",
        "match": "区域就业资源较好，适合作为考研稳妥/保底。",
        "material": "报名信息、复试材料、项目材料。",
        "risk": "不要只看城市，需核验分数线和复试形式。",
        "url": "https://yz.chsi.com.cn/",
    },
]

EXPERIENCES = [
    {
        "title": "双非计算机保研材料准备",
        "path": "保研",
        "why": "排名和项目基础尚可，需要先补材料体系。",
        "action": "7天内完成排名证明、简历和项目说明。",
        "pitfall": "不要等通知出来才开始写材料。",
    },
    {
        "title": "夏令营项目讲解",
        "path": "保研",
        "why": "有课程项目但需要讲清个人贡献。",
        "action": "整理 README 和 15 个项目问答。",
        "pitfall": "不要把团队项目都说成自己做的。",
    },
    {
        "title": "科研空白如何补经历",
        "path": "保研",
        "why": "暂无论文，需要补科研潜力证据。",
        "action": "联系 2 位老师或学长询问复现、阅读、助研小任务。",
        "pitfall": "短期任务不要包装成论文成果。",
    },
    {
        "title": "计算机考研择校表",
        "path": "考研",
        "why": "目标纠结，需要用表格降低不确定性。",
        "action": "建立 10 所院校择校表。",
        "pitfall": "不要同时准备多个差异很大的专业课。",
    },
    {
        "title": "第一份互联网实习怎么投",
        "path": "就业",
        "why": "无实习，需要尽快拿真实反馈。",
        "action": "14 天内投递 10 个岗位。",
        "pitfall": "不要只投大厂核心岗位。",
    },
    {
        "title": "AI 产品实习作品集",
        "path": "就业",
        "why": "CampusPath 本身可以成为作品集。",
        "action": "输出一页作品集和 3 分钟讲解词。",
        "pitfall": "不要只写概念，要展示可运行 Demo。",
    },
]

JOBS = [
    {
        "role": "AI产品策划实习生",
        "company": "腾讯PCG",
        "skills": ["产品设计", "用户研究", "AIGC", "数据分析", "Python", "PRD"],
        "summary": "参与 AI 产品需求分析和校园场景验证。",
        "level": "冲刺",
    },
    {
        "role": "数据分析实习生",
        "company": "互联网公司",
        "skills": ["SQL", "Python", "Excel", "Tableau"],
        "summary": "负责业务数据分析、报表和指标拆解。",
        "level": "稳妥",
    },
    {
        "role": "软件开发实习生",
        "company": "互联网公司",
        "skills": ["Python", "数据库", "前端", "后端", "Git"],
        "summary": "参与业务系统开发、测试和上线。",
        "level": "稳妥",
    },
]


# =========================================================
# 2. 通用工具
# =========================================================

def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def h(value: Any) -> str:
    return escape(str(value or ""))


def clamp(v: float, low: int = 0, high: int = 100) -> int:
    return int(max(low, min(high, round(v))))


def compact_text(text: str, n: int = 160) -> str:
    s = re.sub(r"\s+", " ", text or "").strip()
    return s[:n] + ("..." if len(s) > n else "")


def safe_json_loads(value: Any) -> Any:
    if isinstance(value, (dict, list)):
        return value
    try:
        return json.loads(value)
    except Exception:
        return value


def pretty_value(value: Any) -> str:
    value = safe_json_loads(value)
    if isinstance(value, dict):
        return "\n\n".join([f"**{k}**：{pretty_value(v)}" for k, v in value.items()]) or "暂无"
    if isinstance(value, list):
        return "、".join(map(str, value)) if value else "暂无"
    return str(value)


def render_badge(text: str, color: str = "blue") -> str:
    return f"<span class='cp-badge cp-badge-{color}'>{h(text)}</span>"


def grade_to_100(system: str, score: float) -> int:
    if system == "4分制GPA":
        return clamp(score / 4.0 * 100)
    if system == "5分制GPA":
        return clamp(score / 5.0 * 100)
    return clamp(score)


def english_profile(exam_type: str, score: float) -> Dict[str, Any]:
    if exam_type == "CET-6":
        level = "强" if score >= 520 else "中" if score >= 450 else "弱"
    elif exam_type == "CET-4":
        level = "强" if score >= 550 else "中" if score >= 450 else "弱"
    elif exam_type == "IELTS":
        level = "强" if score >= 6.5 else "中" if score >= 6.0 else "弱"
    elif exam_type == "TOEFL":
        level = "强" if score >= 90 else "中" if score >= 80 else "弱"
    else:
        level = "未知"
    return {
        "type": exam_type,
        "score": score,
        "level": level,
        "score_100": {"强": 85, "中": 68, "弱": 45, "未知": 50}[level],
    }


def read_uploaded_file(uploaded_file: Any) -> str:
    if uploaded_file is None:
        return ""
    name = uploaded_file.name.lower()
    raw = uploaded_file.getvalue()

    if name.endswith((".txt", ".md", ".csv")):
        for enc in ["utf-8-sig", "utf-8", "gbk"]:
            try:
                return raw.decode(enc)
            except UnicodeDecodeError:
                pass
        return "文本解析失败：编码无法识别。"

    if name.endswith(".pdf"):
        if PdfReader is None:
            return "PDF 解析失败：请先 pip install pypdf。"
        try:
            reader = PdfReader(io.BytesIO(raw))
            return "\n".join([(p.extract_text() or "") for p in reader.pages]).strip()
        except Exception as exc:
            return f"PDF 解析失败：{exc}"

    if name.endswith(".docx"):
        if Document is None:
            return "DOCX 解析失败：请先 pip install python-docx。"
        try:
            doc = Document(io.BytesIO(raw))
            return "\n".join([p.text for p in doc.paragraphs]).strip()
        except Exception as exc:
            return f"DOCX 解析失败：{exc}"

    return "暂不支持该文件类型。"


# =========================================================
# 3. 数据库
# =========================================================

def db() -> sqlite3.Connection:
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    return c


def execute(sql: str, params: Tuple[Any, ...] = ()) -> None:
    c = db()
    cur = c.cursor()
    cur.execute(sql, params)
    c.commit()
    c.close()


def query(sql: str, params: Tuple[Any, ...] = ()) -> List[Dict[str, Any]]:
    c = db()
    rows = [dict(r) for r in c.execute(sql, params).fetchall()]
    c.close()
    return rows


def insert(table: str, data: Dict[str, Any]) -> int:
    keys = list(data.keys())
    sql = f"INSERT INTO {table} ({','.join(keys)}) VALUES ({','.join(['?'] * len(keys))})"
    c = db()
    cur = c.cursor()
    cur.execute(sql, [data[k] for k in keys])
    c.commit()
    rid = int(cur.lastrowid)
    c.close()
    return rid


def count_table(table: str) -> int:
    try:
        return int(query(f"SELECT COUNT(*) AS n FROM {table}")[0]["n"])
    except Exception:
        return 0


def count_student(table: str, sid: int | None) -> int:
    if not sid:
        return 0
    try:
        return int(query(f"SELECT COUNT(*) AS n FROM {table} WHERE student_id=?", (sid,))[0]["n"])
    except Exception:
        return 0


def init_db() -> None:
    c = db()
    cur = c.cursor()
    cur.executescript("""
    CREATE TABLE IF NOT EXISTS students (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT, grade TEXT, major TEXT, school_level TEXT,
        grade_system TEXT, grade_score REAL, grade_score_100 REAL,
        rank_percent REAL, english_type TEXT, english_score REAL,
        budget TEXT, targets_json TEXT, created_at TEXT, updated_at TEXT
    );

    CREATE TABLE IF NOT EXISTS source_docs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id INTEGER, source_type TEXT, title TEXT, source_url TEXT,
        raw_text TEXT, summary TEXT, created_at TEXT
    );

    CREATE TABLE IF NOT EXISTS evidence_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id INTEGER, source_doc_id INTEGER, claim TEXT,
        excerpt TEXT, confidence REAL, created_at TEXT
    );

    CREATE TABLE IF NOT EXISTS ontology_nodes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id INTEGER, layer TEXT, object_type TEXT, name TEXT,
        props_json TEXT, evidence_id INTEGER, created_at TEXT, updated_at TEXT
    );

    CREATE TABLE IF NOT EXISTS ontology_edges (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id INTEGER, from_node_id INTEGER, relation TEXT,
        to_node_id INTEGER, evidence_id INTEGER, created_at TEXT
    );

    CREATE TABLE IF NOT EXISTS chat_events (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id INTEGER, event_type TEXT, event_label TEXT, content TEXT,
        deadline TEXT, priority TEXT, path_impact TEXT, suggested_action TEXT,
        source_doc_id INTEGER, created_at TEXT
    );

    CREATE TABLE IF NOT EXISTS decision_reports (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id INTEGER, main_path TEXT, backup_path TEXT,
        scores_json TEXT, report_md TEXT, model_used TEXT, created_at TEXT
    );

    CREATE TABLE IF NOT EXISTS action_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id INTEGER, report_id INTEGER, action_type TEXT,
        priority TEXT, task TEXT, rationale TEXT, action_detail TEXT,
        deadline TEXT, deliverable TEXT, risk TEXT, status TEXT,
        created_at TEXT, updated_at TEXT
    );

    CREATE TABLE IF NOT EXISTS outcomes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id INTEGER, action_item_id INTEGER, outcome_type TEXT,
        title TEXT, description TEXT, created_at TEXT
    );
    """)
    c.commit()
    c.close()


# =========================================================
# 4. 数据采集与本体
# =========================================================

def collect_data() -> Dict[str, Any]:
    g100 = grade_to_100(st.session_state.grade_system, float(st.session_state.grade_score))
    return {
        "name": st.session_state.name,
        "grade": st.session_state.grade,
        "major": st.session_state.major,
        "school_level": st.session_state.school_level,
        "grade_system": st.session_state.grade_system,
        "grade_score": float(st.session_state.grade_score),
        "grade_score_100": g100,
        "rank_percent": int(st.session_state.rank_percent),
        "english_type": st.session_state.english_type,
        "english_score": float(st.session_state.english_score),
        "budget": st.session_state.budget,
        "targets": st.session_state.targets,
        "resume_text": st.session_state.resume_text,
        "chat_text": st.session_state.chat_text,
        "policy_text": st.session_state.policy_text,
        "target_jd": st.session_state.target_jd,
    }


def save_profile_to_db() -> int:
    data = collect_data()
    sid = st.session_state.get("student_id")
    payload = {
        "name": data["name"],
        "grade": data["grade"],
        "major": data["major"],
        "school_level": data["school_level"],
        "grade_system": data["grade_system"],
        "grade_score": data["grade_score"],
        "grade_score_100": data["grade_score_100"],
        "rank_percent": data["rank_percent"],
        "english_type": data["english_type"],
        "english_score": data["english_score"],
        "budget": data["budget"],
        "targets_json": json.dumps(data["targets"], ensure_ascii=False),
        "updated_at": now(),
    }
    exists = bool(sid and query("SELECT id FROM students WHERE id=?", (sid,)))
    if exists:
        execute(
            """UPDATE students
            SET name=?,grade=?,major=?,school_level=?,grade_system=?,grade_score=?,grade_score_100=?,
                rank_percent=?,english_type=?,english_score=?,budget=?,targets_json=?,updated_at=?
            WHERE id=?""",
            tuple(payload.values()) + (sid,),
        )
    else:
        sid = insert("students", {**payload, "created_at": now()})
        st.session_state.student_id = sid
    return int(sid)


def source_doc(student_id: int, source_type: str, title: str, raw_text: str, source_url: str = "") -> int:
    return insert("source_docs", {
        "student_id": student_id,
        "source_type": source_type,
        "title": title,
        "source_url": source_url,
        "raw_text": raw_text or "",
        "summary": compact_text(raw_text, 180),
        "created_at": now(),
    })


def evidence(student_id: int, doc_id: int, claim: str, excerpt: str, confidence: float) -> int:
    return insert("evidence_items", {
        "student_id": student_id,
        "source_doc_id": doc_id,
        "claim": claim,
        "excerpt": compact_text(excerpt, 260),
        "confidence": confidence,
        "created_at": now(),
    })


def node(student_id: int, layer: str, object_type: str, name: str, props: Dict[str, Any], evidence_id: int = 0) -> int:
    return insert("ontology_nodes", {
        "student_id": student_id,
        "layer": layer,
        "object_type": object_type,
        "name": name,
        "props_json": json.dumps(props, ensure_ascii=False),
        "evidence_id": evidence_id,
        "created_at": now(),
        "updated_at": now(),
    })


def edge(student_id: int, from_id: int, relation: str, to_id: int, evidence_id: int = 0) -> int:
    return insert("ontology_edges", {
        "student_id": student_id,
        "from_node_id": from_id,
        "relation": relation,
        "to_node_id": to_id,
        "evidence_id": evidence_id,
        "created_at": now(),
    })


def save_uploaded_source(source_type: str, title: str, text: str, filename: str = "") -> int:
    sid = save_profile_to_db()
    return source_doc(sid, source_type, f"{title}｜{filename}" if filename else title, text)


def apply_uploaded_file_to_state(uploaded_file: Any, text_key: str, sig_key: str, label: str) -> None:
    if uploaded_file is None:
        return
    sig = f"{uploaded_file.name}:{getattr(uploaded_file, 'size', 0)}"
    if st.session_state.get(sig_key) == sig:
        return
    text = read_uploaded_file(uploaded_file)
    if not text or "解析失败" in text or "暂不支持" in text:
        st.error(f"{label}导入失败：{text}")
        return
    st.session_state[text_key] = text
    st.session_state[sig_key] = sig
    mapping = {
        "resume_text": ("resume", "简历"),
        "chat_text": ("qq_wechat_group", "QQ / 微信课程群文本"),
        "policy_text": ("policy", "政策文件"),
        "target_jd": ("target_jd", "目标JD"),
    }
    stype, title = mapping.get(text_key, ("uploaded", label))
    doc_id = save_uploaded_source(stype, title, text, uploaded_file.name)
    st.markdown(f"<div class='cp-source-ok'>✅ {h(label)}已读取并保存到数据库：Source Doc ID {doc_id}</div>", unsafe_allow_html=True)


CHAT_KEYWORDS = {
    "ddl": ["作业", "DDL", "提交", "截止", "上传", "实验报告", "腾讯文档"],
    "exam": ["考试", "期中", "期末", "复习"],
    "material": ["PPT", "PDF", "课件", "资料", "群文件", "录播"],
    "notice": ["通知", "提醒", "老师", "助教", "调课", "签到"],
    "opportunity": ["保研", "夏令营", "预推免", "实习", "招聘", "宣讲", "腾讯", "PCG", "讲座", "报名", "微信群", "QQ群"],
}

EVENT_LABEL = {
    "ddl": "作业DDL",
    "exam": "考试测验",
    "material": "课程资料",
    "notice": "老师通知",
    "opportunity": "升学就业机会",
}


def extract_deadline_hint(text: str) -> str:
    hits = []
    for p in [r"本周[一二三四五六日天][^，。；\n]*", r"下周[一二三四五六日天][^，。；\n]*", r"\d{1,2}:\d{2}"]:
        hits.extend(re.findall(p, text))
    return "；".join(hits[:3]) if hits else "待确认"


def analyze_chat(text: str) -> List[Dict[str, Any]]:
    events = []
    for line in [x.strip() for x in (text or "").splitlines() if x.strip()]:
        for typ, kws in CHAT_KEYWORDS.items():
            if any(k.lower() in line.lower() for k in kws):
                if typ in ["ddl", "exam"]:
                    pri, act, impact = "P0", "加入本周待办，记录截止时间和交付物。", "影响课程成绩，间接影响保研资格、GPA和考研基础。"
                elif typ == "opportunity":
                    pri, act, impact = "P1", "记录报名链接和材料要求，转化为路径验证机会。", "影响保研/就业路径，是外部机会信号。"
                else:
                    pri, act, impact = "P2", "归档到资料库，必要时关联到课程或申请材料。", "影响信息完整度和执行效率。"
                events.append({
                    "event_type": typ,
                    "event_label": EVENT_LABEL[typ],
                    "content": line,
                    "deadline": extract_deadline_hint(line),
                    "priority": pri,
                    "path_impact": impact,
                    "suggested_action": act,
                })
    return events


def extract_skills(text: str) -> List[str]:
    lower = (text or "").lower()
    return sorted({s for s in SKILL_TERMS if s.lower() in lower or s in (text or "")})


def resume_signals(resume_text: str, jd_text: str) -> Dict[str, Any]:
    no_intern = any(x in resume_text for x in ["无正式实习", "暂无实习", "无实习", "没有实习"])
    no_research = any(x in resume_text for x in ["无论文", "暂无论文", "没有论文", "暂无正式科研", "无科研", "暂无科研"])
    has_project = any(x in resume_text for x in ["项目", "平台", "系统", "小程序", "Demo", "README", "Python", "Streamlit"])
    has_intern = "实习" in resume_text and not no_intern
    has_research = any(x in resume_text for x in ["科研", "论文", "实验室", "课题", "专利", "发表", "复现"]) and not no_research
    skills = extract_skills(resume_text)
    jd_skills = extract_skills(jd_text)
    gaps = [s for s in jd_skills if s not in skills]
    return {
        "有项目": has_project,
        "有实习": has_intern,
        "有科研": has_research,
        "无实习": no_intern,
        "无科研": no_research,
        "技能关键词": skills,
        "目标JD技能": jd_skills,
        "目标JD技能缺口": gaps,
    }


def calculate_scores(data: Dict[str, Any], signals: Dict[str, Any], events: List[Dict[str, Any]]) -> Dict[str, int]:
    grade = data["grade_score_100"]
    rank_score = 100 - data["rank_percent"]
    eng = english_profile(data["english_type"], data["english_score"])["score_100"]
    project = 10 if signals["有项目"] else 0
    intern = 15 if signals["有实习"] else 0
    research = 16 if signals["有科研"] else 0
    tb = lambda p: 8 if p in data["targets"] else 0
    opp = 4 if any(e["event_type"] == "opportunity" for e in events) else 0
    gap_penalty = min(len(signals["目标JD技能缺口"]) * 2, 10)
    budget = {"低": 25, "中": 60, "高": 85}.get(data["budget"], 60)
    return {
        "保研": clamp(grade * .32 + rank_score * .30 + eng * .12 + project + research + tb("保研") + opp - max(0, data["rank_percent"] - 20) * .35),
        "考研": clamp(grade * .22 + rank_score * .15 + eng * .18 + 28 + project * .5 + tb("考研")),
        "就业": clamp(grade * .15 + eng * .08 + 30 + project * 1.5 + intern + tb("就业") + opp - gap_penalty),
        "出国": clamp(grade * .22 + eng * .24 + budget * .25 + research * .5 + project * .5 + tb("出国")),
    }


def path_level(score: int) -> str:
    if score >= 80:
        return "高可行"
    if score >= 65:
        return "中高可行"
    if score >= 50:
        return "可尝试但需补短板"
    return "当前不建议作为主线"


def recommend_admissions(data: Dict[str, Any], signals: Dict[str, Any]) -> Dict[str, List[Dict[str, Any]]]:
    groups = {"冲刺": [], "稳妥": [], "保底": []}
    for x in ADMISSIONS:
        score = 60
        if x["type"].startswith("保研") and "保研" in data["targets"]:
            score += 15
        if x["type"] == "考研" and "考研" in data["targets"]:
            score += 15
        if x["level"] == "冲刺":
            score -= 8
        if signals["无科研"] and "保研" in x["type"]:
            score -= 7
        item = {**x, "score": clamp(score)}
        groups[x["level"]].append(item)
    return groups


def recommend_experiences(data: Dict[str, Any], signals: Dict[str, Any]) -> List[Dict[str, Any]]:
    recs = []
    for e in EXPERIENCES:
        score = 0
        if e["path"] in data["targets"] or e["path"] == "通用":
            score += 20
        if signals["无科研"] and "科研" in (e["title"] + e["why"]):
            score += 15
        if signals["无实习"] and ("实习" in e["title"] or "投递" in e["action"]):
            score += 15
        if score > 0:
            recs.append({**e, "score": score})
    return sorted(recs, key=lambda x: x["score"], reverse=True)


def recommend_jobs(signals: Dict[str, Any]) -> List[Dict[str, Any]]:
    current = set(signals["技能关键词"])
    out = []
    for j in JOBS:
        hit = [s for s in j["skills"] if s in current]
        missing = [s for s in j["skills"] if s not in current]
        out.append({**j, "score": clamp(50 + len(hit) * 8 - len(missing) * 4), "hit": hit, "missing": missing})
    return sorted(out, key=lambda x: x["score"], reverse=True)


def clear_working_graph(sid: int) -> None:
    for t in ["evidence_items", "ontology_edges", "ontology_nodes", "chat_events"]:
        execute(f"DELETE FROM {t} WHERE student_id=?", (sid,))


def save_sources_only() -> int:
    data = collect_data()
    sid = save_profile_to_db()
    source_doc(sid, "profile", "学生档案", json.dumps({k: v for k, v in data.items() if not k.endswith("_text")}, ensure_ascii=False, indent=2))
    source_doc(sid, "resume", "简历 / 经历", data["resume_text"])
    source_doc(sid, "qq_wechat_group", "QQ / 微信课程群记录", data["chat_text"])
    source_doc(sid, "policy", "学校推免 / 招生政策", data["policy_text"])
    source_doc(sid, "target_jd", "目标岗位或项目要求", data["target_jd"])
    return sid


def build_ontology_only() -> Dict[str, Any]:
    data = collect_data()
    sid = save_profile_to_db()
    clear_working_graph(sid)

    profile_doc = source_doc(sid, "profile", "学生档案", json.dumps({k: v for k, v in data.items() if not k.endswith("_text")}, ensure_ascii=False, indent=2))
    resume_doc = source_doc(sid, "resume", "简历 / 经历", data["resume_text"])
    chat_doc = source_doc(sid, "qq_wechat_group", "QQ / 微信课程群记录", data["chat_text"])
    policy_doc = source_doc(sid, "policy", "学校推免 / 招生政策", data["policy_text"])
    jd_doc = source_doc(sid, "target_jd", "目标岗位或项目要求", data["target_jd"])

    ev = evidence(sid, profile_doc, "学生档案输入", data["name"] + " " + data["major"], .99)
    stu = node(sid, "Identity", "Student", data["name"], {"年级": data["grade"], "专业": data["major"], "学校层次": data["school_level"]}, ev)

    pref = node(sid, "Identity", "Preference", "目标与约束", {"目标方向": data["targets"], "预算": data["budget"]}, ev)
    edge(sid, stu, "has_preference", pref, ev)

    ac = node(sid, "Academic", "AcademicRecord", "成绩与排名", {
        "成绩": f'{data["grade_system"]} {data["grade_score"]}',
        "标准分": data["grade_score_100"],
        "排名百分比": data["rank_percent"],
        "英语": english_profile(data["english_type"], data["english_score"]),
    }, ev)
    edge(sid, stu, "has_academic_record", ac, ev)

    signals = resume_signals(data["resume_text"], data["target_jd"])
    ev_resume = evidence(sid, resume_doc, "简历抽取能力与短板", data["resume_text"], .78)
    cap = node(sid, "Capability", "CapabilityGraph", "能力图谱", signals, ev_resume)
    edge(sid, stu, "has_capability_graph", cap, ev_resume)

    for sk in signals["技能关键词"]:
        sk_node = node(sid, "Capability", "Skill", sk, {"证据来源": "简历关键词"}, ev_resume)
        edge(sid, cap, "contains_skill", sk_node, ev_resume)

    for gap in signals["目标JD技能缺口"]:
        gap_node = node(sid, "GoalDecision", "Gap", gap, {"说明": "目标 JD 提到，但简历证据不足"}, ev_resume)
        edge(sid, stu, "has_gap", gap_node, ev_resume)

    ev_jd = evidence(sid, jd_doc, "目标机会技能要求", data["target_jd"], .72)
    target = node(sid, "Opportunity", "TargetOpportunity", "目标岗位 / 项目要求", {"目标技能": signals["目标JD技能"], "摘要": compact_text(data["target_jd"], 300)}, ev_jd)
    edge(sid, stu, "targets_opportunity", target, ev_jd)

    ev_pol = evidence(sid, policy_doc, "推免政策解析", data["policy_text"], .72)
    pol = node(sid, "Academic", "Policy", "推免 / 招生政策", {"风险提示": "最终以学校官网、学院官网、研招网最新通知为准"}, ev_pol)
    edge(sid, stu, "constrained_by_policy", pol, ev_pol)

    events = analyze_chat(data["chat_text"])
    for e in events:
        eid = insert("chat_events", {
            "student_id": sid,
            "event_type": e["event_type"],
            "event_label": e["event_label"],
            "content": e["content"],
            "deadline": e["deadline"],
            "priority": e["priority"],
            "path_impact": e["path_impact"],
            "suggested_action": e["suggested_action"],
            "source_doc_id": chat_doc,
            "created_at": now(),
        })
        ev_one = evidence(sid, chat_doc, f"QQ / 微信课程群事件：{e['event_label']}", e["content"], .75)
        event_node = node(sid, "Opportunity" if e["event_type"] == "opportunity" else "Academic", "ChatEvent", e["event_label"], {**e, "事件ID": eid}, ev_one)
        edge(sid, stu, "receives_signal", event_node, ev_one)

    for t in data["targets"]:
        goal_node = node(sid, "GoalDecision", "Goal", t, {"状态": "active"}, ev)
        edge(sid, stu, "pursues_goal", goal_node, ev)

    ctx = build_context(sid, data, signals, events)
    st.session_state.context = ctx
    return ctx


def build_context(sid: int, data: Dict[str, Any], signals: Dict[str, Any], events: List[Dict[str, Any]]) -> Dict[str, Any]:
    scores = calculate_scores(data, signals, events)
    ordered = sorted(scores.items(), key=lambda x: x[1], reverse=True)
    admissions = recommend_admissions(data, signals)
    exps = recommend_experiences(data, signals)
    jobs = recommend_jobs(signals)
    actions = build_actions(data, signals, events)
    return {
        "student_id": sid,
        "profile_summary": f"{data['name']}｜{data['grade']}｜{data['major']}｜标准分 {data['grade_score_100']}｜排名前 {data['rank_percent']}%。建议主路径：{ordered[0][0]}；备选路径：{ordered[1][0]}。",
        "signals": signals,
        "events": events,
        "scores": scores,
        "main_path": ordered[0][0],
        "backup_path": ordered[1][0],
        "admissions": admissions,
        "experiences": exps,
        "jobs": jobs,
        "actions": actions,
    }


def build_actions(data: Dict[str, Any], signals: Dict[str, Any], events: List[Dict[str, Any]]) -> List[Dict[str, str]]:
    tasks = [
        {
            "action_type": "VerifyPolicy",
            "priority": "P0",
            "task": "完成推免资格核验",
            "rationale": "排名在关键线附近，保研是否成立取决于资格、挂科、加分和学院名额。",
            "action_detail": "核对排名证明、挂科记录、科研竞赛加分项和9月审核时间。",
            "deadline": "3天内",
            "deliverable": "推免资格核对表 + 缺口清单",
            "risk": "政策最终以学校和学院官网当年通知为准。",
        },
        {
            "action_type": "CreatePortfolio",
            "priority": "P0",
            "task": "把校园二手交易平台升级为可展示作品集",
            "rationale": "项目是当前最强能力证据，可同时服务保研面试、考研复试和实习投递。",
            "action_detail": "补 README、功能截图、数据流图、用户痛点、个人贡献、3个难点和3条量化简历 bullet。",
            "deadline": "7天内",
            "deliverable": "项目 README + 作品集一页纸 + 简历 bullet",
            "risk": "不要把团队贡献全部包装成个人贡献。",
        },
    ]

    if signals["无科研"]:
        tasks.append({
            "action_type": "RequestResearchTask",
            "priority": "P1",
            "task": "启动科研/实验室小任务补强",
            "rationale": "科研空白会影响保研导师匹配。",
            "action_detail": "联系2位老师或学长，询问复现实验、数据整理、论文阅读或助研小任务。",
            "deadline": "10天内",
            "deliverable": "2封联系消息 + 1份科研意向说明",
            "risk": "短期任务不能包装成论文成果。",
        })

    if signals["无实习"]:
        tasks.append({
            "action_type": "TrackApplication",
            "priority": "P1",
            "task": "启动第一批实习投递",
            "rationale": "就业路径最大短板是无正式实习。",
            "action_detail": "投递10个产品/开发/测试/数据分析岗位，记录岗位、简历版本和反馈。",
            "deadline": "14天内",
            "deliverable": "10条投递记录 + 1版岗位定制简历",
            "risk": "不要只投大厂核心岗位。",
        })

    if signals["目标JD技能缺口"]:
        tasks.append({
            "action_type": "AnalyzeSkillGap",
            "priority": "P0",
            "task": "补齐目标JD技能缺口",
            "rationale": "目标 JD 中缺少强证据：" + "、".join(signals["目标JD技能缺口"][:6]),
            "action_detail": "选择一个小项目，把缺口技能绑定到项目成果和简历 bullet。",
            "deadline": "21天内",
            "deliverable": "目标岗位技能差距项目 v1",
            "risk": "只学概念没有作品，会无法支撑面试。",
        })

    for e in events:
        if e["priority"] in ["P0", "P1"]:
            tasks.append({
                "action_type": "CreateTaskFromGroupSignal",
                "priority": e["priority"],
                "task": f"处理课程群信号：{e['event_label']}",
                "rationale": e["path_impact"],
                "action_detail": e["suggested_action"] + " 原文：" + e["content"][:80],
                "deadline": e["deadline"],
                "deliverable": "待办记录 / 报名截图 / 讲座笔记 / 复习清单",
                "risk": "课程群信息如果不进入 Action OS，容易被刷屏淹没。",
            })

    return tasks[:12]


# =========================================================
# 5. 报告
# =========================================================

def make_prompt(data: Dict[str, Any], ctx: Dict[str, Any]) -> str:
    safe_data = {k: v for k, v in data.items() if not k.endswith("_text")}
    slim = {
        "profile_summary": ctx["profile_summary"],
        "scores": ctx["scores"],
        "main_path": ctx["main_path"],
        "backup_path": ctx["backup_path"],
        "signals": ctx["signals"],
        "events": ctx["events"][:5],
        "actions": ctx["actions"][:8],
        "admissions": ctx["admissions"],
        "experiences": ctx["experiences"][:5],
        "jobs": ctx["jobs"][:3],
    }
    payload = json.dumps({
        "student_input": safe_data,
        "resume_excerpt": compact_text(data["resume_text"], 700),
        "qq_wechat_group_excerpt": compact_text(data["chat_text"], 700),
        "policy_excerpt": compact_text(data["policy_text"], 700),
        "target_jd_excerpt": compact_text(data["target_jd"], 700),
        "decision_context": slim,
    }, ensure_ascii=False, indent=2)

    return f"""
你是 CampusPath，一个面向腾讯 PCG 校园 AI 产品大赛的大学生个人进化决策系统，不是完整自主 Agent。
请生成真正给学生看的报告：清晰、短句、可执行、少术语。

硬性要求：
1. 体现 QQ / 微信课程群、腾讯文档、PCG 场景。
2. 不得承诺一定保研、一定上岸、一定录取、一定拿 offer。
3. 院校政策以官网 / 学院官网 / 研招网为准。
4. 经验帖只能作为参考。
5. 每个建议包含优先级、具体动作、截止时间、交付物、风险。
6. 报告控制在 {AI_REPORT_WORDS} 字左右。

请按以下结构输出：
# 你的 CampusPath 路径建议
## 1. 一句话结论
## 2. 你现在最该先做的 3 件事
## 3. 四条路径怎么选
## 4. QQ / 微信课程群里发现的任务和机会
## 5. 院校 / 岗位 / 经验帖建议
## 6. 30 / 60 / 90 天行动计划
## 7. 可复制材料：导师邮件、简历 bullet、60秒自我介绍
## 8. 风险提醒和下一次复盘

输入数据：
{payload}
"""


def call_model(prompt: str) -> Tuple[bool, str, str]:
    if not MODEL_API_KEY:
        return False, "", "未检测到 MODEL_API_KEY。"
    try:
        s = requests.Session()
        s.trust_env = USE_SYSTEM_PROXY
        resp = s.post(
            MODEL_API_URL,
            headers={"Authorization": f"Bearer {MODEL_API_KEY}", "Content-Type": "application/json"},
            json={
                "model": MODEL_NAME,
                "messages": [
                    {"role": "system", "content": "你是严谨、具体、证据驱动的大学生规划助手。"},
                    {"role": "user", "content": prompt},
                ],
                "temperature": 0.35,
                "max_tokens": AI_MAX_TOKENS,
            },
            timeout=AI_TIMEOUT_SEC,
        )
        if resp.status_code != 200:
            return False, "", f"模型接口错误 HTTP {resp.status_code}: {resp.text[:500]}"
        return True, resp.json()["choices"][0]["message"]["content"], "模型调用成功。"
    except requests.exceptions.Timeout:
        return False, "", f"模型调用超时：超过 {AI_TIMEOUT_SEC} 秒，已切换本地兜底。"
    except requests.exceptions.ProxyError as exc:
        return False, "", f"模型调用失败：代理连接失败，已切换本地兜底。原始错误：{exc}"
    except Exception as exc:
        return False, "", f"模型调用失败：{exc}"


def local_report(data: Dict[str, Any], ctx: Dict[str, Any]) -> str:
    scores = ctx["scores"]
    signals = ctx["signals"]
    score_rows = "\n".join([
        f"| {p} | {scores[p]} | {path_level(scores[p])} | {'主路径' if p == ctx['main_path'] else '备选路径' if p == ctx['backup_path'] else '低成本保留'} |"
        for p in PATHS
    ])
    top3_md = "\n".join([
        f"### {i+1}. {a['task']}\n"
        f"- 优先级：{a['priority']}\n"
        f"- 今天怎么做：{a['action_detail']}\n"
        f"- 截止时间：{a['deadline']}\n"
        f"- 交付物：{a['deliverable']}\n"
        f"- 风险：{a['risk']}"
        for i, a in enumerate(ctx["actions"][:3])
    ])
    events_md = "\n".join([
        f"- **{e['event_label']}**：{e['content']} → {e['suggested_action']}"
        for e in ctx["events"][:6]
    ]) or "- 暂无课程群信号。"

    adm_md = "\n".join([
        f"- {level}：{x['school']}｜{x['program']}｜{x['match']}｜材料：{x['material']}"
        for level, rows in ctx["admissions"].items()
        for x in rows[:2]
    ])
    exp_md = "\n".join([
        f"- {x['title']}：{x['action']}｜避坑：{x['pitfall']}"
        for x in ctx["experiences"][:4]
    ])
    job_md = "\n".join([
        f"- {x['company']} {x['role']}：匹配度 {x['score']}，缺口：{'、'.join(x['missing']) if x['missing'] else '暂无明显缺口'}"
        for x in ctx["jobs"][:3]
    ])
    bullets = "\n".join([
        "- 使用 Python / Streamlit 搭建校园二手交易平台 Demo，完成商品发布、搜索、收藏、详情展示和联系卖家等核心流程。",
        "- 负责需求拆解、页面原型、核心流程开发和演示材料整理，将课程项目包装为可展示的完整产品原型。",
        "- 梳理校园二手交易场景与商品信息结构，沉淀用户流程、功能模块和项目复盘材料，可迁移至 AI 产品实习作品集。",
    ])

    return f"""# 你的 CampusPath 路径建议

## 1. 一句话结论
当前建议你把 **{ctx['main_path']}** 作为主路径验证，把 **{ctx['backup_path']}** 作为备选路径。这个建议不是保证结果，而是基于你的成绩、排名、项目、QQ / 微信课程群信息、院校库和岗位 JD 得出的阶段性判断。

## 2. 你现在最该先做的 3 件事
{top3_md}

## 3. 四条路径怎么选
| 路径 | 分数 | 可行性 | 当前定位 |
|---|---:|---|---|
{score_rows}

你的优势是：{ '、'.join(signals['技能关键词']) if signals['技能关键词'] else '已有课程项目和较稳定成绩' }。

你的主要短板是：{ '、'.join(signals['目标JD技能缺口']) if signals['目标JD技能缺口'] else '科研/实习证据还需要继续补强' }。

## 4. QQ / 微信课程群里发现的任务和机会
{events_md}

## 5. 院校 / 岗位 / 经验帖建议
### 院校项目
{adm_md}

### 岗位机会
{job_md}

### 经验帖可借鉴动作
{exp_md}

## 6. 30 / 60 / 90 天行动计划
- **30天内**：完成推免资格核验、一页简历、项目 README、10条投递记录或3个目标院校核验。
- **60天内**：完成目标院校/岗位梯度表，每档至少3个；完成一次项目模拟面试。
- **90天内**：根据投递、导师回复、课程成绩和面试结果，确定主路径和备选路径。

## 7. 可复制材料
### 简历 bullet
{bullets}

### 60秒自我介绍
老师/面试官您好，我是{data['name']}，{data['major']}专业{data['grade']}学生。我的成绩基础相对稳定，专业排名约前 {data['rank_percent']}%。我做过校园二手交易平台项目，主要使用 Python 和 Streamlit 完成商品发布、搜索、收藏和联系卖家等功能。接下来我会把这个项目升级成可展示作品集，同时补充科研或实习证据。

## 8. 风险提醒和下一次复盘
院校政策最终以学校官网、学院官网和研招网最新通知为准。经验帖只作为结构化参考。你每完成一个行动，都应该在 Outcome 里记录结果，让下一次路径判断更准确。
"""


def save_report_and_actions(ctx: Dict[str, Any], report: str, model: str) -> int:
    sid = ctx["student_id"]
    execute("DELETE FROM action_items WHERE student_id=?", (sid,))
    rid = insert("decision_reports", {
        "student_id": sid,
        "main_path": ctx["main_path"],
        "backup_path": ctx["backup_path"],
        "scores_json": json.dumps(ctx["scores"], ensure_ascii=False),
        "report_md": report,
        "model_used": model,
        "created_at": now(),
    })
    for a in ctx["actions"]:
        insert("action_items", {
            "student_id": sid,
            "report_id": rid,
            "action_type": a["action_type"],
            "priority": a["priority"],
            "task": a["task"],
            "rationale": a["rationale"],
            "action_detail": a["action_detail"],
            "deadline": a["deadline"],
            "deliverable": a["deliverable"],
            "risk": a["risk"],
            "status": "todo",
            "created_at": now(),
            "updated_at": now(),
        })
    st.session_state.report_id = rid
    return rid


def generate_decision_report(use_model: bool = True) -> None:
    box = st.empty()
    with box.container():
        st.info("正在生成报告，请不要关闭页面。")
        progress = st.progress(0)
        msg = st.empty()

        msg.write("1/5 读取输入...")
        progress.progress(10)
        data = collect_data()

        msg.write("2/5 构建个人本体...")
        progress.progress(35)
        ctx = build_ontology_only()

        msg.write("3/5 生成报告提示词...")
        progress.progress(55)
        prompt = make_prompt(data, ctx)

        if use_model and MODEL_API_KEY:
            msg.write(f"4/5 调用云端模型，最多等待 {AI_TIMEOUT_SEC} 秒...")
            progress.progress(75)
            ok, out, info = call_model(prompt)
            if ok:
                report, model = out, MODEL_NAME
                st.success(info)
            else:
                report, model = local_report(data, ctx), "本地兜底"
                st.warning(info)
        else:
            msg.write("4/5 使用本地规则生成快速报告...")
            progress.progress(75)
            report, model = local_report(data, ctx), "本地兜底"

        msg.write("5/5 保存报告和行动项...")
        progress.progress(92)
        rid = save_report_and_actions(ctx, report, model)

        st.session_state.report_md = report
        st.session_state.context = ctx
        st.session_state.report_id = rid

        progress.progress(100)
        st.success(f"完成：Student ID {ctx['student_id']}｜Report ID {rid}")


# =========================================================
# 6. 前端
# =========================================================

def inject_css() -> None:
    st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
[data-testid="stToolbar"] {display: none !important;}
[data-testid="stDecoration"] {display: none !important;}
[data-testid="stStatusWidget"] {display: none !important;}
[data-testid="stDeployButton"] {display: none !important;}
.block-container{max-width:1500px;padding-top:1.1rem}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#101828 0%,#1d2939 100%)}
[data-testid="stSidebar"] *{color:#f9fafb!important}
.cp-hero{border-radius:30px;padding:36px;margin-bottom:18px;background:radial-gradient(circle at 10% 8%,rgba(59,130,246,.26),transparent 28%),radial-gradient(circle at 88% 10%,rgba(124,58,237,.22),transparent 26%),linear-gradient(135deg,#fff 0%,#eef4ff 48%,#f5f3ff 100%);border:1px solid rgba(34,81,255,.15);box-shadow:0 18px 50px rgba(16,24,40,.10)}
.cp-eyebrow{display:inline-flex;border:1px solid #c7d7fe;color:#1d4ed8;background:#eff6ff;padding:7px 12px;border-radius:999px;font-size:13px;font-weight:800;margin-bottom:12px}
.cp-title{font-size:44px;line-height:1.05;font-weight:900;letter-spacing:-1.4px;color:#101828;margin:0}
.cp-subtitle{color:#475467;font-size:18px;line-height:1.65;margin-top:14px;max-width:1050px}
.cp-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin:18px 0}
.cp-card,.cp-soft-card{background:#fff;border:1px solid #e4e7ec;border-radius:22px;padding:18px;box-shadow:0 12px 32px rgba(16,24,40,.06);margin:10px 0}
.cp-card h3{margin:0 0 8px 0;font-size:15px;color:#344054}
.cp-card .big{font-size:30px;font-weight:900;color:#101828}
.cp-card p,.cp-card-body{color:#667085;margin:5px 0 0 0;font-size:14px;line-height:1.65}
.cp-section{margin:18px 0 10px 0;font-size:24px;font-weight:850;color:#101828}
.cp-note{padding:13px 15px;background:#f8fafc;border:1px dashed #cbd5e1;border-radius:16px;color:#475467;font-size:13px;line-height:1.6}
.cp-warning{padding:13px 15px;background:#fffaeb;border:1px solid #fedf89;border-radius:16px;color:#93370d;font-size:13px;line-height:1.6}
.cp-card-head{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:8px}
.cp-card-title{font-size:17px;font-weight:800;color:#101828}
.cp-badge{display:inline-block;padding:4px 10px;border-radius:999px;font-size:12px;font-weight:800;margin:2px 4px 2px 0;border:1px solid transparent}
.cp-badge-blue{color:#1d4ed8;background:#eff6ff;border-color:#bfdbfe}
.cp-badge-green{color:#067647;background:#ecfdf3;border-color:#abefc6}
.cp-badge-orange{color:#b54708;background:#fffaeb;border-color:#fedf89}
.cp-badge-purple{color:#6941c6;background:#f4f3ff;border-color:#d9d6fe}
.cp-source-ok{padding:10px 12px;border-radius:14px;background:#ecfdf3;border:1px solid #abefc6;color:#067647;font-size:13px;margin:6px 0}
.cp-mini{color:#667085;font-size:13px}
</style>
""", unsafe_allow_html=True)


def init_session() -> None:
    defaults = {
        "nav": "决策驾驶舱",
        "_nav_widget": "决策驾驶舱",
        "student_id": None,
        "report_id": None,
        "context": None,
        "report_md": "",
        "name": "张同学",
        "grade": "大三",
        "major": "计算机科学与技术",
        "school_level": "双非一本",
        "grade_system": "4分制GPA",
        "grade_score": 3.62,
        "rank_percent": 18,
        "english_type": "CET-6",
        "english_score": 520.0,
        "budget": "中",
        "targets": ["保研", "考研", "就业"],
        "resume_text": DEFAULT_RESUME,
        "chat_text": DEFAULT_CHAT,
        "policy_text": DEFAULT_POLICY,
        "target_jd": DEFAULT_JD,
        "resume_file_sig": "",
        "chat_file_sig": "",
        "policy_file_sig": "",
        "jd_file_sig": "",
    }
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)


def rerun() -> None:
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()


def goto(page: str) -> None:
    st.session_state.nav = page
    rerun()


def on_sidebar_nav_change() -> None:
    st.session_state.nav = st.session_state._nav_widget


def sync_nav_widget_before_render() -> None:
    if st.session_state.get("_nav_widget") != st.session_state.get("nav"):
        st.session_state._nav_widget = st.session_state.nav


def page_header(title: str, subtitle: str = "", back: bool = True) -> None:
    cols = st.columns([1, 6])
    if back:
        with cols[0]:
            if st.button("← 返回", use_container_width=True, key=f"back_{title}"):
                goto("决策驾驶舱")
        with cols[1]:
            st.markdown(f"<div class='cp-section'>{h(title)}</div>", unsafe_allow_html=True)
            if subtitle:
                st.caption(subtitle)
    else:
        st.markdown(f"<div class='cp-section'>{h(title)}</div>", unsafe_allow_html=True)
        if subtitle:
            st.caption(subtitle)


def hero() -> None:
    st.markdown(f"""
<div class='cp-hero'>
  <div class='cp-eyebrow'>Personal Ontology · AI Decision Copilot · Action OS</div>
  <h1 class='cp-title'>{h(APP_NAME)}｜{h(APP_SUBTITLE)}</h1>
  <div class='cp-subtitle'>把学生分散在简历、成绩、QQ / 微信课程群、腾讯文档报名链接、院校招生公告、经验帖摘要和岗位 JD 里的信息，自动沉淀为个人本体，并驱动保研 / 考研 / 就业 / 出国路径决策与行动闭环。</div>
</div>
""", unsafe_allow_html=True)


def dashboard_metrics() -> None:
    sid = st.session_state.get("student_id")
    facts = count_student("ontology_nodes", sid) if sid else count_table("ontology_nodes")
    events = count_student("chat_events", sid) if sid else count_table("chat_events")
    actions = count_student("action_items", sid) if sid else count_table("action_items")
    reports = count_table("decision_reports")
    st.markdown(f"""
<div class='cp-grid'>
  <div class='cp-card'><h3>Ontology 对象</h3><div class='big'>{facts}</div><p>Student / Skill / Goal / Opportunity / Risk</p></div>
  <div class='cp-card'><h3>课程群信号</h3><div class='big'>{events}</div><p>QQ / 微信中的 DDL、考试、资料、通知、机会</p></div>
  <div class='cp-card'><h3>Action OS 任务</h3><div class='big'>{actions}</div><p>建议被写成可执行动作</p></div>
  <div class='cp-card'><h3>决策报告</h3><div class='big'>{reports}</div><p>每次报告可追溯证据链</p></div>
</div>
""", unsafe_allow_html=True)


def score_cards(ctx: Dict[str, Any]) -> None:
    for col, p in zip(st.columns(4), PATHS):
        col.metric(p, f"{ctx['scores'][p]} / 100", path_level(ctx["scores"][p]))


def page_dashboard() -> None:
    hero()
    dashboard_metrics()
    st.markdown(f"<div class='cp-note'>{h(PCG_BOUNDARY)}</div>", unsafe_allow_html=True)
    st.markdown("<div class='cp-section'>核心功能入口</div>", unsafe_allow_html=True)

    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:
        st.markdown("### 🪪 导入档案")
        st.caption("学生信息、简历、QQ / 微信课程群、政策、目标 JD")
        if st.button("进入数据采集", use_container_width=True):
            goto("数据采集")

    with c2:
        st.markdown("### 🧠 构建本体")
        st.caption("生成可展开的个人本体")
        if st.button("构建个人本体", type="primary", use_container_width=True):
            ctx = build_ontology_only()
            st.success(f"本体已构建：Student ID {ctx['student_id']}")

    with c3:
        st.markdown("### 🚀 路径决策")
        st.caption("快速报告几秒返回")
        if st.button("⚡ 快速报告", type="primary", use_container_width=True):
            generate_decision_report(use_model=False)
        if st.button("🧠 深度 AI 报告", use_container_width=True):
            generate_decision_report(use_model=True)

    with c4:
        st.markdown("### ✅ Action OS")
        st.caption("查看和更新行动")
        if st.button("打开 Action OS", use_container_width=True):
            goto("Action OS")

    with c5:
        st.markdown("### 🔁 Outcome")
        st.caption("记录结果反馈")
        if st.button("记录 Outcome", use_container_width=True):
            goto("Outcome反馈")

    st.divider()

    if st.session_state.get("student_id"):
        st.success(f"当前 Student ID：{st.session_state.student_id}")
    else:
        st.warning("当前还没有导入个人档案。请先进入数据采集，或直接点击构建个人本体。")

    if st.session_state.get("context"):
        st.markdown("<div class='cp-section'>当前路径决策</div>", unsafe_allow_html=True)
        score_cards(st.session_state.context)
        st.success(st.session_state.context["profile_summary"])
        if st.session_state.get("report_md"):
            with st.expander("查看完整报告", expanded=True):
                st.markdown(st.session_state.report_md)


def page_data() -> None:
    page_header("数据采集：Observe 层", "上传文件后会自动保存到本地数据库。支持 QQ / 微信课程群文本。")
    st.markdown("<div class='cp-warning'>Demo 不读取真实 QQ/微信数据库，不自动爬取小红书。你可以上传或粘贴导出的文本。</div>", unsafe_allow_html=True)

    if st.session_state.get("student_id"):
        st.success(f"当前 Student ID：{st.session_state.student_id}")
    else:
        st.info("当前还没有 Student ID。填写档案后点击「导入个人档案到数据库」。")

    with st.form("profile_form"):
        a, b, c = st.columns(3)

        with a:
            st.text_input("姓名 / 昵称", key="name")
            st.selectbox("年级", ["大一", "大二", "大三", "大四", "研一", "研二"], key="grade")
            st.text_input("专业", key="major")

        with b:
            st.selectbox("学校层次", ["C9/985", "985", "211", "双非一本", "普通本科", "专科"], key="school_level")
            st.selectbox("成绩制度", ["4分制GPA", "5分制GPA", "百分制"], key="grade_system")
            max_s = 4.0 if st.session_state.grade_system == "4分制GPA" else 5.0 if st.session_state.grade_system == "5分制GPA" else 100.0
            if st.session_state.grade_score > max_s:
                st.session_state.grade_score = max_s
            st.number_input("成绩 / GPA", min_value=0.0, max_value=max_s, step=0.01, key="grade_score")

        with c:
            st.number_input("专业排名百分比", min_value=1, max_value=100, key="rank_percent")
            st.selectbox("英语考试类型", ["CET-4", "CET-6", "IELTS", "TOEFL", "GRE", "暂无"], key="english_type")
            st.number_input("英语成绩", min_value=0.0, step=0.5, key="english_score")

        d, e = st.columns(2)

        with d:
            st.selectbox("预算", ["低", "中", "高"], key="budget")

        with e:
            st.multiselect("目标方向", PATHS, key="targets")

        if st.form_submit_button("🪪 导入个人档案到数据库", use_container_width=True):
            st.success(f"个人档案已写入 SQLite。Student ID：{save_profile_to_db()}")

    st.divider()
    t1, t2, t3, t4 = st.tabs(["📄 简历", "💬 QQ / 微信课程群", "🏫 推免/政策/招生简章", "💼 目标岗位JD"])

    with t1:
        f = st.file_uploader("上传简历 txt / md / pdf / docx", type=["txt", "md", "pdf", "docx"], key="resume_file_uploader")
        apply_uploaded_file_to_state(f, "resume_text", "resume_file_sig", "简历")
        st.text_area("简历 / 经历文本", height=260, key="resume_text")

    with t2:
        f = st.file_uploader("上传 QQ / 微信课程群聊天文本 txt / md", type=["txt", "md"], key="chat_file_uploader")
        apply_uploaded_file_to_state(f, "chat_text", "chat_file_sig", "QQ / 微信课程群文本")
        st.text_area("QQ / 微信课程群文本 Demo", height=260, key="chat_text")

    with t3:
        f = st.file_uploader("上传推免细则 / 招生简章 txt / md / pdf / docx", type=["txt", "md", "pdf", "docx"], key="policy_file_uploader")
        apply_uploaded_file_to_state(f, "policy_text", "policy_file_sig", "政策文件")
        st.text_area("政策 / 推免细则文本", height=230, key="policy_text")

    with t4:
        f = st.file_uploader("上传目标岗位 JD txt / md / pdf / docx", type=["txt", "md", "pdf", "docx"], key="jd_file_uploader")
        apply_uploaded_file_to_state(f, "target_jd", "jd_file_sig", "目标 JD")
        st.text_area("目标岗位 JD / 目标项目要求", height=260, key="target_jd")

    st.divider()
    st.markdown("<div class='cp-section'>下一步操作</div>", unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        if st.button("🪪 保存档案", use_container_width=True):
            st.success(f"已保存个人档案。Student ID：{save_profile_to_db()}")

    with c2:
        if st.button("📥 保存当前全部文本", use_container_width=True):
            st.success(f"当前全部文本已保存到 source_docs。Student ID：{save_sources_only()}")

    with c3:
        if st.button("🧠 构建个人本体", type="primary", use_container_width=True):
            st.success(f"个人本体已构建。Student ID：{build_ontology_only()['student_id']}")

    with c4:
        if st.button("⚡ 快速路径报告", type="primary", use_container_width=True):
            generate_decision_report(use_model=False)
        if st.button("🧠 深度 AI 报告", use_container_width=True):
            generate_decision_report(use_model=True)


def render_node_props(props_json: Any) -> None:
    props = safe_json_loads(props_json)
    if isinstance(props, dict):
        for k, v in props.items():
            st.markdown(f"- **{k}**：{pretty_value(v)}")
    else:
        st.write(pretty_value(props))


def render_node_relations(sid: int, node_id: int) -> None:
    relations = query(
        """SELECT e.relation, fn.name AS from_name, tn.name AS to_name
        FROM ontology_edges e
        LEFT JOIN ontology_nodes fn ON e.from_node_id=fn.id
        LEFT JOIN ontology_nodes tn ON e.to_node_id=tn.id
        WHERE e.student_id=? AND (e.from_node_id=? OR e.to_node_id=?)
        ORDER BY e.id DESC LIMIT 8""",
        (sid, node_id, node_id),
    )
    if not relations:
        st.caption("暂无关系记录")
        return
    for r in relations:
        st.markdown(f"- **{r.get('from_name','')}** → `{r.get('relation','')}` → **{r.get('to_name','')}**")


def page_ontology() -> None:
    page_header("个人本体：Student Life Ontology", "可以点开每个对象，查看属性、关系和证据来源。")
    sid = st.session_state.get("student_id")
    if not sid:
        st.info("请先在「数据采集」页导入档案并构建本体。")
        return

    nodes = query(
        """SELECT n.*, e.claim, e.excerpt, e.confidence
        FROM ontology_nodes n
        LEFT JOIN evidence_items e ON n.evidence_id=e.id
        WHERE n.student_id=?
        ORDER BY n.id DESC""",
        (sid,),
    )
    if not nodes:
        st.warning("当前还没有本体对象。请先点击「构建个人本体」。")
        return

    a, b, c, d = st.columns(4)
    a.metric("本体对象", count_student("ontology_nodes", sid))
    b.metric("关系边", count_student("ontology_edges", sid))
    c.metric("证据", count_student("evidence_items", sid))
    d.metric("数据源", count_student("source_docs", sid))

    layer_names = {
        "Identity": "🧑 身份与偏好",
        "Academic": "🎓 学业与政策",
        "Capability": "🧩 能力与项目",
        "GoalDecision": "🎯 目标与差距",
        "Opportunity": "🚪 机会信号",
        "Outcome": "🔁 结果反馈",
    }

    layers = [x for x in layer_names if any(n["layer"] == x for n in nodes)] + sorted({n["layer"] for n in nodes if n["layer"] not in layer_names})
    tabs = st.tabs([layer_names.get(l, l) for l in layers])

    for tab, layer in zip(tabs, layers):
        with tab:
            for n in [x for x in nodes if x["layer"] == layer]:
                with st.expander(f"{n['object_type']}｜{n['name']}", expanded=False):
                    st.markdown(render_badge(n["object_type"], "blue") + render_badge(f"Evidence {n.get('evidence_id') or '-'}", "purple"), unsafe_allow_html=True)
                    st.markdown("#### 对象属性")
                    render_node_props(n.get("props_json"))
                    st.markdown("#### 关系")
                    render_node_relations(sid, n["id"])
                    st.markdown("#### 证据来源")
                    st.markdown(f"**判断依据：** {n.get('claim') or '暂无证据描述'}")
                    st.markdown(f"> {n.get('excerpt') or '暂无证据摘录'}")
                    try:
                        st.progress(float(n.get("confidence") or 0))
                        st.caption(f"置信度：{round(float(n.get('confidence') or 0)*100)}%")
                    except Exception:
                        st.caption("置信度：暂无")


def page_decision() -> None:
    page_header("AI 路径决策：Decide 层", "报告改成给学生看的清晰执行建议。")
    ctx = st.session_state.get("context")
    if not ctx:
        st.info("请先生成报告或构建本体。")
        return

    score_cards(ctx)
    st.success(ctx["profile_summary"])
    tabs = st.tabs(["课程群信号", "院校梯度", "经验库", "岗位匹配", "完整报告"])

    with tabs[0]:
        if not ctx["events"]:
            st.info("暂无课程群信号。")
        for e in ctx["events"]:
            st.markdown(f"- **{e['event_label']}**｜{e['priority']}｜{e['content']} → {e['suggested_action']}")

    with tabs[1]:
        for level, rows in ctx["admissions"].items():
            st.subheader(level)
            for x in rows:
                st.markdown(f"- **{x['school']}**｜{x['program']}｜匹配：{x['match']}｜材料：{x['material']}｜风险：{x['risk']}")

    with tabs[2]:
        for x in ctx["experiences"]:
            st.markdown(f"- **{x['title']}**：{x['action']}｜避坑：{x['pitfall']}")

    with tabs[3]:
        for x in ctx["jobs"]:
            st.markdown(f"- **{x['company']} {x['role']}**｜匹配度 {x['score']}｜缺口：{'、'.join(x['missing']) if x['missing'] else '暂无'}")

    with tabs[4]:
        if st.session_state.get("report_md"):
            st.markdown(st.session_state.report_md)
        else:
            st.info("已构建本体，但还没有生成完整报告。")


def page_action_os() -> None:
    page_header("Action OS：Act 层", "把建议拆成可执行行动项。")
    sid = st.session_state.get("student_id")
    if not sid:
        st.info("请先生成报告。")
        return

    rows = query("SELECT * FROM action_items WHERE student_id=? ORDER BY CASE priority WHEN 'P0' THEN 1 WHEN 'P1' THEN 2 ELSE 3 END, id DESC", (sid,))
    if not rows:
        st.info("暂无行动项。请先生成报告。")
        return

    filt = st.multiselect("状态筛选", STATUS_OPTIONS, default=STATUS_OPTIONS)

    for r in [x for x in rows if x["status"] in filt]:
        with st.expander(f"#{r['id']}｜{r['priority']}｜{r['task']}｜{r['status']}", expanded=r["priority"] == "P0"):
            st.markdown(render_badge(r["priority"], "orange" if r["priority"] == "P0" else "blue") + render_badge(r["status"], "purple"), unsafe_allow_html=True)
            st.write("**依据：**", r["rationale"])
            st.write("**具体动作：**", r["action_detail"])
            st.write("**截止时间：**", r["deadline"])
            st.write("**交付物：**", r["deliverable"])
            st.write("**风险：**", r["risk"])

            new_status = st.selectbox("状态", STATUS_OPTIONS, index=STATUS_OPTIONS.index(r["status"]) if r["status"] in STATUS_OPTIONS else 0, key=f"status_{r['id']}")
            if st.button("更新状态", key=f"upd_{r['id']}"):
                execute("UPDATE action_items SET status=?, updated_at=? WHERE id=?", (new_status, now(), r["id"]))
                st.success("已更新。")


def page_outcome() -> None:
    page_header("Outcome 反馈：Reflect 层", "记录行动结果，让下一轮决策更准确。")
    sid = st.session_state.get("student_id")
    if not sid:
        st.info("请先生成报告，再记录反馈。")
        return

    actions = query("SELECT id, priority, task, status, deliverable FROM action_items WHERE student_id=? ORDER BY id DESC", (sid,))
    if not actions:
        st.info("暂无可关联的行动项。请先生成路径报告。")
        return

    c1, c2 = st.columns([1, 1])

    with c1:
        st.markdown("### 🔁 记录一次结果反馈")
        opts = {f"#{a['id']}｜{a['priority']}｜{a['task']}｜{a['status']}": a["id"] for a in actions}

        with st.form("outcome_form_pretty"):
            selected = st.selectbox("关联行动项", list(opts.keys()))
            typ = st.selectbox("反馈类型", ["任务完成", "已投递", "收到面试", "拿到 Offer", "收到拒信", "导师回复", "考试成绩", "阶段复盘"])
            title = st.text_input("反馈标题", "完成项目 README v1")
            desc = st.text_area("结果描述 / 复盘", "完成了功能截图、技术栈说明、个人贡献和下一步优化计划。")
            submitted = st.form_submit_button("写入反馈并更新个人本体", use_container_width=True)

        if submitted:
            type_map = {
                "任务完成": "task_done",
                "已投递": "application_sent",
                "收到面试": "interview",
                "拿到 Offer": "offer",
                "收到拒信": "rejection",
                "导师回复": "mentor_reply",
                "考试成绩": "exam_score",
                "阶段复盘": "reflection",
            }
            aid = opts[selected]
            out_type = type_map.get(typ, "reflection")

            oid = insert("outcomes", {
                "student_id": sid,
                "action_item_id": aid,
                "outcome_type": out_type,
                "title": title,
                "description": desc,
                "created_at": now(),
            })
            doc = source_doc(sid, "outcome", title, desc)
            ev = evidence(sid, doc, f"Outcome：{typ}", desc, .92)
            node(sid, "Outcome", "Outcome", title, {"类型": typ, "描述": desc, "outcome_id": oid}, ev)

            execute("UPDATE action_items SET status=?, updated_at=? WHERE id=?", ("done", now(), aid))
            st.success("反馈已写入，本体已更新，关联行动项已标记为 done。")

    with c2:
        st.markdown("### ✅ 待反馈行动项")
        for a in actions[:8]:
            st.markdown(
                f"<div class='cp-soft-card'><div class='cp-card-head'><div class='cp-card-title'>#{a['id']}｜{h(a['task'])}</div><div>{render_badge(a['priority'], 'orange' if a['priority']=='P0' else 'blue')}{render_badge(a['status'], 'purple')}</div></div><div class='cp-card-body'>交付物：{h(a.get('deliverable','暂无'))}</div></div>",
                unsafe_allow_html=True,
            )

    st.divider()
    st.markdown("### 最近反馈记录")
    outs = query("SELECT o.*, a.task FROM outcomes o LEFT JOIN action_items a ON o.action_item_id=a.id WHERE o.student_id=? ORDER BY o.id DESC LIMIT 10", (sid,))
    if not outs:
        st.info("暂无反馈记录。")
        return
    for o in outs:
        st.markdown(
            f"<div class='cp-soft-card'><div class='cp-card-head'><div class='cp-card-title'>🔁 {h(o.get('title',''))}</div><div>{render_badge(o.get('outcome_type','feedback'), 'green')}</div></div><div class='cp-card-body'><b>关联行动：</b>{h(o.get('task') or '未关联')}<br><b>反馈内容：</b>{h(o.get('description') or '暂无描述')}<div class='cp-mini'>记录时间：{h(o.get('created_at',''))}</div></div></div>",
            unsafe_allow_html=True,
        )


def page_database() -> None:
    page_header("数据资产与知识库", "默认只展示产品化摘要，开发者表格隐藏在调试区。")
    a, b, c, d = st.columns(4)
    a.metric("学生档案", count_table("students"))
    b.metric("数据源文档", count_table("source_docs"))
    c.metric("决策报告", count_table("decision_reports"))
    d.metric("反馈记录", count_table("outcomes"))

    st.markdown("### 最近导入的数据源")
    docs = query("SELECT id, source_type, title, summary, created_at FROM source_docs ORDER BY id DESC LIMIT 12")

    if not docs:
        st.info("暂无数据源。请先在「数据采集」页上传或保存文件。")

    for drow in docs:
        color = {
            "resume": "blue",
            "qq_wechat_group": "green",
            "policy": "orange",
            "target_jd": "purple",
            "outcome": "green",
        }.get(drow.get("source_type", ""), "blue")
        st.markdown(
            f"<div class='cp-soft-card'><div class='cp-card-head'><div class='cp-card-title'>📄 {h(drow.get('title',''))}</div><div>{render_badge(drow.get('source_type',''), color)}</div></div><div class='cp-card-body'>{h(drow.get('summary',''))}<div class='cp-mini'>Source Doc ID：{h(drow.get('id'))}｜导入时间：{h(drow.get('created_at'))}</div></div></div>",
            unsafe_allow_html=True,
        )

    with st.expander("开发者调试区：路演时不要打开", expanded=False):
        table = st.selectbox("选择表", ["students", "source_docs", "evidence_items", "ontology_nodes", "ontology_edges", "chat_events", "decision_reports", "action_items", "outcomes"], key="debug_table_select")
        st.dataframe(pd.DataFrame(query(f"SELECT * FROM {table} ORDER BY id DESC LIMIT 300")), use_container_width=True, height=420)

        if st.button("危险操作：重置本地数据库", key="reset_db_button"):
            if DB_PATH.exists():
                DB_PATH.unlink()
            init_db()
            st.session_state.student_id = None
            st.session_state.report_id = None
            st.session_state.context = None
            st.session_state.report_md = ""
            st.success("已重置数据库，请刷新页面。")


def page_submission() -> None:
    page_header("腾讯比赛提交包辅助页", "录屏和 PDF 可以按这个结构准备。")
    st.markdown("""
### 3分钟视频结构
- 0:00-0:25 痛点：学生信息分散在成绩、简历、QQ / 微信课程群、招生公告和经验帖中。
- 0:25-0:55 数据采集：展示档案、简历、QQ / 微信课程群、政策、目标 JD。
- 0:55-1:25 个人本体：展示可展开的 Nodes / Edges / Evidence。
- 1:25-1:55 路径决策：展示保研 / 考研 / 就业 / 出国评分。
- 1:55-2:20 外部机会：展示院校梯度、经验帖、岗位匹配。
- 2:20-2:45 Action OS：展示 AI 建议如何变成任务。
- 2:45-3:00 Outcome：展示结果写回本体，形成个人进化闭环。
""")


def init_session() -> None:
    defaults = {
        "nav": "决策驾驶舱",
        "_nav_widget": "决策驾驶舱",
        "student_id": None,
        "report_id": None,
        "context": None,
        "report_md": "",
        "name": "张同学",
        "grade": "大三",
        "major": "计算机科学与技术",
        "school_level": "双非一本",
        "grade_system": "4分制GPA",
        "grade_score": 3.62,
        "rank_percent": 18,
        "english_type": "CET-6",
        "english_score": 520.0,
        "budget": "中",
        "targets": ["保研", "考研", "就业"],
        "resume_text": DEFAULT_RESUME,
        "chat_text": DEFAULT_CHAT,
        "policy_text": DEFAULT_POLICY,
        "target_jd": DEFAULT_JD,
        "resume_file_sig": "",
        "chat_file_sig": "",
        "policy_file_sig": "",
        "jd_file_sig": "",
    }
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)


def rerun() -> None:
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()


def goto(page: str) -> None:
    st.session_state.nav = page
    rerun()


def on_sidebar_nav_change() -> None:
    st.session_state.nav = st.session_state._nav_widget


def sync_nav_widget_before_render() -> None:
    if st.session_state.get("_nav_widget") != st.session_state.get("nav"):
        st.session_state._nav_widget = st.session_state.nav


def main() -> None:
    st.set_page_config(page_title="CampusPath Decision OS", page_icon="🧭", layout="wide")
    inject_css()
    init_db()
    init_session()
    sync_nav_widget_before_render()

    with st.sidebar:
        st.title("🧭 CampusPath")
        st.caption("Personal Ontology + AI Decision Copilot + Action OS")
        st.radio("导航", PAGES, key="_nav_widget", on_change=on_sidebar_nav_change)

        st.divider()
        st.success(f"Student ID：{st.session_state.student_id}") if st.session_state.get("student_id") else st.warning("未导入学生档案")
        st.success(f"Report ID：{st.session_state.report_id}") if st.session_state.get("report_id") else st.info("暂无报告")

        st.divider()
        st.write("模型状态")
        st.code(MODEL_NAME)
        st.success("已检测到 MODEL_API_KEY") if MODEL_API_KEY else st.warning("未配置 Key，将使用本地兜底")
        st.caption(f"USE_SYSTEM_PROXY={USE_SYSTEM_PROXY}｜AI_TIMEOUT_SEC={AI_TIMEOUT_SEC}")

    page = st.session_state.nav

    if page == "决策驾驶舱":
        page_dashboard()
    elif page == "数据采集":
        page_data()
    elif page == "个人本体":
        page_ontology()
    elif page == "AI路径决策":
        page_decision()
    elif page == "Action OS":
        page_action_os()
    elif page == "Outcome反馈":
        page_outcome()
    elif page == "数据资产":
        page_database()
    elif page == "提交包":
        page_submission()


if __name__ == "__main__":
    main()