import sys
import os
import json
import base64
import time
import re
import traceback
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

from openai import OpenAI

# ================= 1. Word 与 UI 库导入 =================
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QListWidget, QListWidgetItem, QSplitter,
    QScrollArea, QFrame, QFileDialog, QProgressBar, QMessageBox,
    QDialog, QFormLayout, QLineEdit, QComboBox, QToolBar,
    QSizePolicy, QTabWidget, QTextEdit, QGroupBox, QGridLayout,
    QSpinBox, QPlainTextEdit, QDialogButtonBox,
    QToolButton, QMenu
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QPointF, QRectF
from PyQt6.QtGui import (
    QPixmap, QIcon, QColor, QAction, QPainter, QPen, QBrush, QFont, QImage,
    QPainterPath
)

# === Graphics 组件 ===
from PyQt6.QtWidgets import (
    QGraphicsView, QGraphicsScene, QGraphicsPixmapItem
)

CONFIG_FILE = "app_config_lec.json"
TEMPLATE_NAME = "模板.docx"
MAX_IMAGES = 20

EXPORT_IMG_DIR = "_export_marked"  # 导出用的带标注图片目录

# ================= 2. 核心默认数据配置 =================

DEFAULT_BUSINESS_DATA = {
    "company_project_map": {
        "勐海县泽兴供水有限公司": ["城乡供水一体化项目", "勐海农村供水保障项目"],
        "勐海县润博水利投资有限公司": ["勐阿水库建设项目"],
        "江城县润成水利投资有限公司": ["热水河水库建设项目"],
        "澜沧县润成水利投资有限公司": ["三道箐水库建设项目"]
    },
    "company_unit_map": {
        "勐海县泽兴供水有限公司": "云南建投第二水利水电建设有限公司",
        "勐海县润博水利投资有限公司": "云南建投第二水利水电建设有限公司",
        "江城县润成水利投资有限公司": "云南建投第二水利水电建设有限公司",
        "澜沧县润成水利投资有限公司": "云南省水利水电工程有限公司"
    },
    "check_content_options": [
        "安全文明施工专项检查",
        "工程质量专项检查",
        "项目综合检查",
        "节前安全生产检查",
        "复工复产专项检查"
    ],
    "project_overview_map": {
        "勐海农村供水保障项目": "本工程位于西双版纳州勐海县，主要建设内容包括新建取水坝、输水管网及配套水厂设施，旨在解决周边5个乡镇的农村饮水安全问题，设计供水规模为2.5万吨/日。",
        "城乡供水一体化项目": "勐海县城乡供水一体化建设项目涉及勐海县城、勐遮镇、勐混镇、勐阿镇、打洛镇、勐满镇、格朗和乡、勐宋乡8个片区，覆盖现状人口28.53万人。主要建设内容为：新建3座水厂，总建设规模32000m³/d，其中县城三水厂20000m³/d，格朗和乡4000m³/d，勐混镇 8000m³/d。扩建水厂1座，勐遮镇扩容建设5000m³/d 工艺设施，扩容后总处理规模15000m³/d。利用存量水厂7座，现状总供水规模61500m³/d，其中县城一水厂10000m³/d，县城二水厂 20000m³/d，勐遮水厂10000m³/d，打洛镇曼彦水厂7500m³/d，勐阿水厂6000m³/d，勐满水厂4000m³/d，勐宋水厂4000m³/d。建设DN100-DN900输配水管网376.87km，配套建设信息化设施、阀门井、排泥阀、闸阀、入户管及其他附属设施。",
        "勐阿水库建设项目": "勐海县勐阿水库项目总投资7.645亿元。勐海县勐阿水库规模为中型，由枢纽工程、输(供)水工程和水厂工程组成，枢纽工程由大坝、溢洪道和输水(兼导流)隧洞组成。合同工期48个月。",
        "热水河水库建设项目": "江城县热水河水库工程主要由枢纽工程和输水工程两部分组成。枢纽工程主要包括拦河坝、溢洪道、导流输水隧洞等，建成后将有效缓解江城县城供水压力。江城热水河水库项目总投资5.61亿元。江城县热水河水库工程由枢纽工程和输水工程组成，合同工期为48个月。",
        "三道箐水库建设项目": "澜沧县三道箐水库位于澜沧县中北部的东河乡拉巴河上游的三道箐河上，水库工程由枢纽工程及灌区工程组成。枢纽工程主要由大坝、1～2#副坝、溢洪道、输水导流兼放空隧洞及主坝～1#副坝库岸防渗组成。水库为小（1）型水库，总库容406万m3，澜沧县三道箐水库项目总投资2.32808亿元，合同工期24个月。"
    }
}

DEFAULT_PROMPTS = {
    "🏗️ 施工全能扫描 (安质+文明施工)": """你是一位拥有30年经验的“工程质量安全总监”。请对施工现场照片进行“地毯式”深度排查，覆盖【安全隐患】、【实体质量】、【文明施工】三个维度。

### 一、 核心任务目标
**尽可能全面地罗列出所有肉眼可见的问题**。宁可错杀，不可漏过。

### 二、 评判标准体系 (定性分级)

**1. 🔴 严重/红线问题 (对应红色)**
   - **安全**: 致命风险。例：临边无防护、高处作业未系安全带、特种设备关键缺失、深基坑边堆载严重、灭火器失效、私拉乱接电线。
   - **质量**: 结构性缺陷。例：混凝土严重狗洞/露筋、受力钢筋截断/间距严重错位、承重墙裂缝、防水层严重破损。

**2. 🟠 一般/较大问题 (对应橙色)**
   - **安全**: 违规行为。例：未佩戴安全帽（或未系下颌带）、梯子不稳、气瓶无防震圈、脚手架踏板未铺满。
   - **质量**: 规范不符。例：钢筋轻微锈蚀、砖墙灰缝不饱满、模板拼缝不严、保护层垫块缺失。

**3. 🔵 文明施工与改进 (对应蓝色)**
   - **现场脏乱**: 地面积水、垃圾未清理、材料堆放杂乱无章、裸土未覆盖（扬尘）。
   - **标识缺失**: 缺少警示牌、缺少操作规程牌。
   - **外观瑕疵**: 墙面轻微污染、线条不直。

### 三、 重点排查清单 (请逐一扫描)

1. **人的不安全行为**: 安全帽(带子)、反光衣、安全带(高挂低用)、吸烟、穿拖鞋。
2. **物的不安全状态**: 
   - **临电**: 必须“一机一闸一漏”，电缆不得泡水/拖地凌乱。
   - **架体**: 剪刀撑是否连续、立杆悬空、扣件缺失。
   - **机械**: 吊钩防脱、钢丝绳断丝、违规载人。
3. **实体质量**: 蜂窝麻面、裂缝、烂根、钢筋间距、搭接长度、直螺纹套筒。
4. **文明施工 (5S)**: "工完场清"是否落实？材料是否分类码放？道路是否硬化？是否存在扬尘隐患？

### 四、 输出格式 (JSON)
必须严格返回 JSON 数组，不要 Markdown 标记。
`risk_level` 必须包含“严重”、“一般”或“文明施工”字样以触发颜色警告。

### 五、 视觉定位（Bounding Box）
虽然我需要你识别问题，但请尽量给出该问题在图片中的矩形框坐标 bbox。
- bbox 形式：[x1, y1, x2, y2]
- 坐标单位：像素
- 坐标基于：原图尺寸（不是缩放后的预览图）
- (0,0) 为图片左上角，x 向右，y 向下
- 若无法可靠定位：bbox 返回 null

[
  {
    "risk_level": "严重安全隐患",
    "issue": "……",
    "regulation": "……",
    "correction": "……",
    "bbox": [0,0,0,0],
    "confidence": 0.0
  }
]""",
    "🏠 纯日常生活 (整理/健康/居家)": """你是一位资深的生活管家。请以提升生活品质为目标，分析照片中的场景。

### 输出格式 (JSON)
[
    {
        "risk_level": "卫生警示",
        "issue": "……",
        "regulation": "食品卫生常识",
        "correction": "……",
        "bbox": null,
        "confidence": 0.0
    }
]"""
}

DEFAULT_PROVIDER_PRESETS = {
    "阿里百炼 (Qwen-VL)": {
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "model": "qwen-vl-max"
    },
    "硅基流动 (SiliconFlow)": {
        "base_url": "https://api.siliconflow.cn/v1",
        "model": "Qwen/Qwen2-VL-72B-Instruct"
    },
    "DeepSeek (官方)": {
        "base_url": "https://api.deepseek.com/v1",
        "model": "deepseek-chat"
    },
    "OpenAI (GPT-4o)": {
        "base_url": "https://api.openai.com/v1",
        "model": "gpt-4o"
    },
    "自定义 (Custom)": {
        "base_url": "",
        "model": ""
    }
}


# ================= 3. 配置管理 =================

class ConfigManager:
    @staticmethod
    def get_default_config():
        return {
            "current_provider": "阿里百炼 (Qwen-VL)",
            "api_key": "",
            "last_prompt": list(DEFAULT_PROMPTS.keys())[0],
            "custom_provider_settings": {"base_url": "", "model": ""},
            "business_data": DEFAULT_BUSINESS_DATA,
            "prompts": DEFAULT_PROMPTS,
            "provider_presets": DEFAULT_PROVIDER_PRESETS,

            "max_concurrency": 3,
            "max_retries": 2,
            "request_timeout_sec": 60,
            "temperature": 0.1,

            "last_check_person": "",
            "recent_check_areas": []
        }

    @staticmethod
    def load():
        default = ConfigManager.get_default_config()
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    saved = json.load(f)

                # 深度补全
                for k, v in default.items():
                    if k not in saved:
                        saved[k] = v

                if "business_data" not in saved:
                    saved["business_data"] = default["business_data"]
                else:
                    for key in default["business_data"]:
                        if key not in saved["business_data"]:
                            saved["business_data"][key] = default["business_data"][key]

                if "prompts" not in saved:
                    saved["prompts"] = default["prompts"]

                if "provider_presets" not in saved:
                    saved["provider_presets"] = default["provider_presets"]

                return saved
            except Exception as e:
                print(f"配置文件加载失败，使用默认值: {e}")
        else:
            ConfigManager.save(default)

        return default

    @staticmethod
    def save(config):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"保存配置文件失败: {e}")


# ================= 4. JSON 解析与清洗 =================

def _strip_code_fences(text: str) -> str:
    t = (text or "").strip()
    t = t.replace("```json", "").replace("```JSON", "").replace("```", "")
    return t.strip()


def _extract_json_array_candidate(text: str) -> Optional[str]:
    if not text:
        return None
    t = _strip_code_fences(text)

    if t.startswith("[") and t.endswith("]"):
        return t

    s = t.find("[")
    e = t.rfind("]")
    if s != -1 and e != -1 and e > s:
        return t[s:e + 1]

    m = re.search(r"\[[\s\S]*\]", t)
    if m:
        return m.group(0)

    return None


def _repair_common_json_issues(s: str) -> str:
    if not s:
        return s
    s = s.replace("\ufeff", "").strip()
    s = s.replace("“", "\"").replace("”", "\"").replace("‘", "'").replace("’", "'")
    s = re.sub(r",\s*([}\]])", r"\1", s)
    return s


def _normalize_bbox(b: Any) -> Optional[List[int]]:
    if b is None:
        return None
    if not isinstance(b, (list, tuple)) or len(b) != 4:
        return None
    try:
        x1, y1, x2, y2 = [int(float(v)) for v in b]
    except Exception:
        return None
    x1, x2 = sorted([x1, x2])
    y1, y2 = sorted([y1, y2])
    if x2 - x1 <= 1 or y2 - y1 <= 1:
        return None
    return [x1, y1, x2, y2]


def parse_issues_from_model_output(raw: str) -> Tuple[List[Dict[str, Any]], Optional[str]]:
    if raw is None:
        return [], "空响应"

    candidate = _extract_json_array_candidate(raw)
    if not candidate:
        return [], "未找到 JSON 数组"

    candidate = _repair_common_json_issues(candidate)

    try:
        data = json.loads(candidate)
        if not isinstance(data, list):
            return [], "JSON 顶层不是数组"

        norm: List[Dict[str, Any]] = []
        for item in data:
            if not isinstance(item, dict):
                continue
            bbox = _normalize_bbox(item.get("bbox", None))
            conf = item.get("confidence", None)
            try:
                conf_f = float(conf) if conf is not None else None
            except Exception:
                conf_f = None

            norm.append({
                "risk_level": str(item.get("risk_level", "")).strip(),
                "issue": str(item.get("issue", "")).strip(),
                "regulation": str(item.get("regulation", "")).strip(),
                "correction": str(item.get("correction", "")).strip(),
                "bbox": bbox,
                "confidence": conf_f
            })
        return norm, None
    except Exception as e:
        return [], f"JSON 解析失败: {e}"


# ================= 5. 画框/叠加标注：导出图片工具 =================

def ensure_export_dir() -> str:
    if not os.path.exists(EXPORT_IMG_DIR):
        os.makedirs(EXPORT_IMG_DIR, exist_ok=True)
    return EXPORT_IMG_DIR


def _risk_pen(level: str) -> QPen:
    lv = level or ""
    if any(x in lv for x in ["重大", "严重", "红线"]):
        color = QColor("#FF0000")
    elif any(x in lv for x in ["一般", "较大", "质量"]):
        color = QColor("#FF8800")
    else:
        color = QColor("#2196F3")
    pen = QPen(color, 6)
    pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
    return pen


def draw_user_annotations(img: QImage, annotations: List[Dict[str, Any]]) -> QImage:
    """
    把用户涂鸦烧录到图像上。annotations 坐标为原图像素坐标。
    """
    if img.isNull():
        return img
    if not annotations:
        return img
    out = img.copy()
    p = QPainter(out)
    p.setRenderHint(QPainter.RenderHint.Antialiasing, True)

    for a in annotations:
        t = a.get("type")
        color = QColor(a.get("color", "#FF0000"))
        w = int(a.get("width", 6))
        pen = QPen(color, w)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        p.setPen(pen)
        p.setBrush(Qt.BrushStyle.NoBrush)

        if t == "rect":
            x1, y1, x2, y2 = a.get("bbox", [0, 0, 0, 0])
            p.drawRect(QRectF(x1, y1, x2 - x1, y2 - y1))
        elif t == "ellipse":
            x1, y1, x2, y2 = a.get("bbox", [0, 0, 0, 0])
            p.drawEllipse(QRectF(x1, y1, x2 - x1, y2 - y1))
        elif t == "arrow":
            x1, y1 = a.get("p1", [0, 0])
            x2, y2 = a.get("p2", [0, 0])
            p.drawLine(QPointF(x1, y1), QPointF(x2, y2))
            # 箭头
            import math
            angle = math.atan2(y2 - y1, x2 - x1)
            head_len = 28
            head_ang = math.radians(28)
            p1 = QPointF(x2 - head_len * math.cos(angle - head_ang), y2 - head_len * math.sin(angle - head_ang))
            p2 = QPointF(x2 - head_len * math.cos(angle + head_ang), y2 - head_len * math.sin(angle + head_ang))
            p.drawLine(QPointF(x2, y2), p1)
            p.drawLine(QPointF(x2, y2), p2)
        elif t == "text":
            x, y = a.get("pos", [0, 0])
            txt = a.get("text", "")
            font = QFont()
            font.setPointSize(int(a.get("font_size", 28)))
            font.setBold(True)
            p.setFont(font)
            p.setPen(QPen(color, max(2, w // 2)))
            # 白底描边增强可读性
            outline = QPainterPath()
            outline.addText(QPointF(x, y), font, txt)
            p.setPen(QPen(QColor(255, 255, 255, 220), 10))
            p.drawPath(outline)
            p.setPen(QPen(color, 4))
            p.drawText(QPointF(x, y), txt)

    p.end()
    return out


def build_export_marked_image(original_path: str,
                              issues: List[Dict[str, Any]],
                              user_annotations: List[Dict[str, Any]],
                              out_path: str) -> bool:
    img = QImage(original_path)
    if img.isNull():
        return False

    # 【修改】：不再调用 draw_ai_bboxes_on_image，直接使用原图作为底图进行用户标注绘制
    # img2 = draw_ai_bboxes_on_image(img, issues)
    img2 = img.copy()

    # 叠加用户的手动标注（含手动引用的问题描述）
    img3 = draw_user_annotations(img2, user_annotations)

    # 输出 PNG
    ok = img3.save(out_path, "PNG")
    return bool(ok)


# ================= 6. Word 报告生成器 =================

class WordReportGenerator:
    @staticmethod
    def set_font(run, font_name='宋体', size=None, bold=False, color=None):
        run.font.name = font_name
        r = run._element
        r.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        if size:
            run.font.size = Pt(size)
        run.font.bold = bold
        if color:
            run.font.color.rgb = color

    @staticmethod
    def _replace_text_in_paragraph(paragraph, replacements):
        if not paragraph.text:
            return
        for key, value in replacements.items():
            if key in paragraph.text:
                val_str = str(value) if value else ""
                paragraph.text = paragraph.text.replace(key, val_str)
                for run in paragraph.runs:
                    WordReportGenerator.set_font(run, size=12, bold=run.font.bold)

    @staticmethod
    def replace_placeholders(doc, info):
        replacements = {
            "{{项目公司名称}}": info.get("project_company", ""),
            "{{项目名称}}": info.get("project_name", ""),
            "{{检查部位}}": info.get("check_area", ""),
            "{{检查人员}}": info.get("check_person", ""),
            "{{被检查单位}}": info.get("inspected_unit", ""),
            "{{检查内容}}": info.get("check_content", ""),
            "{{项目概况}}": info.get("project_overview", ""),
            "{{检查日期}}": info.get("check_date", ""),
            "{{整改期限}}": info.get("rectification_deadline", "")
        }
        for para in doc.paragraphs:
            WordReportGenerator._replace_text_in_paragraph(para, replacements)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        WordReportGenerator._replace_text_in_paragraph(para, replacements)

    @staticmethod
    def _dedupe_keep_order(items: List[str]) -> List[str]:
        seen = set()
        out = []
        for x in items:
            x2 = (x or "").strip()
            if not x2 or x2 in ["无", "暂无", "N/A", "无明显隐患"]:
                continue
            if x2 not in seen:
                seen.add(x2)
                out.append(x2)
        return out

    @staticmethod
    def generate(tasks: List[Dict[str, Any]], save_path: str, project_info: Dict[str, str],
                 template_path=TEMPLATE_NAME):
        if os.path.exists(template_path):
            doc = Document(template_path)
        else:
            doc = Document()
            section = doc.sections[0]
            section.top_margin = Cm(2.0)
            section.bottom_margin = Cm(2.0)
            section.left_margin = Cm(2.0)
            section.right_margin = Cm(2.0)
            doc.add_paragraph(f"【注意】未找到模板文件 {template_path}，使用默认空白格式。")

        WordReportGenerator.replace_placeholders(doc, project_info)
        doc.add_paragraph()

        valid_tasks = []
        for t in tasks:
            has_issues = t.get("status") == "done"
            has_annotations = bool(t.get("annotations"))
            if has_issues or has_annotations:
                valid_tasks.append(t)

        if not valid_tasks:
            doc.add_paragraph("【提示】当前没有已完成分析或已标注的图片任务。")
            doc.save(save_path)
            return

        for idx, task in enumerate(valid_tasks, 1):
            table = doc.add_table(rows=1, cols=1)
            table.style = 'Table Grid'
            table.autofit = False
            cell = table.cell(0, 0)
            cell.width = Cm(17.0)

            p_title = cell.paragraphs[0]
            p_title.paragraph_format.space_before = Pt(4)
            p_title.paragraph_format.space_after = Pt(4)
            p_title.paragraph_format.left_indent = Cm(0.2)

            title = f"问题 {idx}"
            group = (task.get("meta") or {}).get("group")
            if group:
                title += f"（点位：{group}）"

            if task.get("status") != "done":
                title += " (人工标注项)"

            run_title = p_title.add_run(title)
            WordReportGenerator.set_font(run_title, size=12, bold=True)

            issues = task.get("edited_issues") if task.get("edited_issues") is not None else task.get("issues")
            if not issues:
                issues = []

            safety_texts, quality_texts, civil_texts = [], [], []
            corrections = []

            for item in issues:
                r_level = (item.get("risk_level") or "").strip()
                issue = (item.get("issue") or "").strip()
                reg = (item.get("regulation") or "").strip()
                corr = (item.get("correction") or "").strip()

                if not issue:
                    continue

                full_desc = issue
                if reg and reg not in ["无", "常识", "食品卫生常识"]:
                    full_desc += f"（违反 {reg}）"

                if "质量" in r_level:
                    quality_texts.append(full_desc)
                elif "文明" in r_level:
                    civil_texts.append(full_desc)
                else:
                    safety_texts.append(full_desc)

                if corr:
                    corrections.append(corr)

            def add_section(label: str, texts: List[str], color: Optional[RGBColor] = None):
                if not texts:
                    return
                p = cell.add_paragraph()
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.left_indent = Cm(0.2)
                p.paragraph_format.right_indent = Cm(0.2)
                p.paragraph_format.line_spacing = 1.2
                run_label = p.add_run(label)
                WordReportGenerator.set_font(run_label, bold=True, size=11)
                merged_txt = "；".join(texts) + "。"
                run_text = p.add_run(merged_txt)
                WordReportGenerator.set_font(run_text, size=11, color=color)

            add_section("安全问题：", safety_texts)
            add_section("质量问题：", quality_texts)
            add_section("文明施工问题：", civil_texts)

            if not (safety_texts or quality_texts or civil_texts) and task.get("annotations"):
                p_note = cell.add_paragraph()
                p_note.paragraph_format.left_indent = Cm(0.2)
                run_note = p_note.add_run("详情见图片标注（人工补充）。")
                WordReportGenerator.set_font(run_note, size=11, color=RGBColor(0, 0, 0))

            p_corr = cell.add_paragraph()
            p_corr.paragraph_format.space_before = Pt(2)
            p_corr.paragraph_format.space_after = Pt(2)
            p_corr.paragraph_format.left_indent = Cm(0.2)
            p_corr.paragraph_format.right_indent = Cm(0.2)
            p_corr.paragraph_format.line_spacing = 1.2
            run_label = p_corr.add_run("整改要求：")
            WordReportGenerator.set_font(run_label, bold=True, size=11)

            dedup = WordReportGenerator._dedupe_keep_order(corrections)
            if not dedup:
                run_text = p_corr.add_run("无。")
                WordReportGenerator.set_font(run_text, size=11, color=RGBColor(0, 100, 0))
            else:
                for i, c in enumerate(dedup, 1):
                    p = cell.add_paragraph()
                    p.paragraph_format.space_before = Pt(1)
                    p.paragraph_format.space_after = Pt(1)
                    p.paragraph_format.left_indent = Cm(0.8)
                    p.paragraph_format.right_indent = Cm(0.2)
                    p.paragraph_format.line_spacing = 1.2
                    run = p.add_run(f"{i}. {c}")
                    WordReportGenerator.set_font(run, size=11, color=RGBColor(0, 100, 0))

            img_path = task.get("export_image_path")
            if not img_path or not os.path.exists(img_path):
                img_path = task.get("path", "")

            if img_path and os.path.exists(img_path):
                p_img = cell.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(4)
                p_img.paragraph_format.space_after = Pt(4)
                try:
                    p_img.add_run().add_picture(img_path, width=Cm(13.5))
                except Exception as e:
                    p_img.add_run(f"[图片加载失败: {str(e)}]")
            else:
                p_img = cell.add_paragraph()
                p_img.add_run("[图片文件缺失]")

            if idx < len(valid_tasks):
                spacer = doc.add_paragraph()
                spacer.paragraph_format.space_after = Pt(10)

        doc.save(save_path)


# ================= 7. 后台分析线程 =================

def build_strict_json_guard() -> str:
    return """
你必须严格输出 JSON 数组（以 [ 开始，以 ] 结束），不要输出任何解释文字、不要输出 Markdown。
规则：
1) 每个元素必须包含 risk_level、issue、regulation、correction 四个字段。
2) 若画面信息不足，请在 issue/correction 中明确写“无法确认/疑似/建议现场复核”，不要编造具体参数。
3) risk_level 必须包含以下之一：严重、一般、文明施工。
4) 若能定位，请额外输出 bbox 字段：[x1,y1,x2,y2]（像素坐标，基于原图尺寸）。无法定位则 bbox 为 null。
""".strip()


class AnalysisWorker(QThread):
    finished = pyqtSignal(str, dict)

    def __init__(self, task: dict, config: dict, prompt_text: str):
        super().__init__()
        self.task = task
        self.config = config
        self.prompt_text = prompt_text

    def _get_provider_conf(self) -> Tuple[str, str, str, Optional[str]]:
        p_name = self.config.get("current_provider")
        api_key = self.config.get("api_key")

        presets = self.config.get("provider_presets", DEFAULT_PROVIDER_PRESETS)
        p_conf = presets.get(p_name, {})
        base_url = p_conf.get("base_url")
        model = p_conf.get("model")

        if p_name == "自定义 (Custom)" and (not base_url or not model):
            custom_sets = self.config.get("custom_provider_settings", {})
            base_url = custom_sets.get("base_url")
            model = custom_sets.get("model")

        return p_name, api_key, base_url, model

    def _should_retry(self, err: Exception) -> bool:
        msg = str(err).lower()
        retry_tokens = ["timeout", "timed out", "429", "rate", "limit", "overloaded", "503", "connection",
                        "temporarily"]
        return any(t in msg for t in retry_tokens)

    def run(self):
        started = time.time()
        try:
            p_name, api_key, base_url, model = self._get_provider_conf()

            if not api_key:
                self.finished.emit(self.task['id'], {"ok": False, "error": "未配置 API Key"})
                return
            if not base_url or not model:
                self.finished.emit(self.task['id'], {"ok": False, "error": "未配置模型 Base URL 或 名称"})
                return

            client = OpenAI(api_key=api_key, base_url=base_url)

            with open(self.task['path'], "rb") as f:
                b64 = base64.b64encode(f.read()).decode()

            max_retries = int(self.config.get("max_retries", 2))
            temperature = float(self.config.get("temperature", 0.1))

            system_prompt = self.prompt_text.strip() + "\n\n" + build_strict_json_guard()

            last_err = None
            raw_content = ""
            for attempt in range(max_retries + 1):
                try:
                    resp = client.chat.completions.create(
                        model=model,
                        messages=[
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": [
                                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}},
                                {"type": "text", "text": "请按要求输出 JSON 数组。"}
                            ]}
                        ],
                        temperature=temperature
                    )
                    raw_content = resp.choices[0].message.content or ""
                    issues, parse_err = parse_issues_from_model_output(raw_content)

                    elapsed = time.time() - started
                    if parse_err:
                        self.finished.emit(self.task['id'], {
                            "ok": False,
                            "error": parse_err,
                            "raw_output": raw_content,
                            "issues": [],
                            "elapsed_sec": round(elapsed, 2),
                            "provider": p_name,
                            "model": model
                        })
                        return

                    self.finished.emit(self.task['id'], {
                        "ok": True,
                        "error": None,
                        "raw_output": raw_content,
                        "issues": issues,
                        "elapsed_sec": round(elapsed, 2),
                        "provider": p_name,
                        "model": model
                    })
                    return

                except Exception as e:
                    last_err = e
                    if attempt < max_retries and self._should_retry(e):
                        backoff = min(8, 2 ** attempt)
                        time.sleep(backoff)
                        continue
                    break

            elapsed = time.time() - started
            self.finished.emit(self.task['id'], {
                "ok": False,
                "error": str(last_err) if last_err else "未知错误",
                "raw_output": raw_content,
                "issues": [],
                "elapsed_sec": round(elapsed, 2),
                "provider": p_name,
                "model": model
            })

        except Exception as e:
            elapsed = time.time() - started
            self.finished.emit(self.task['id'], {
                "ok": False,
                "error": f"{e}\n{traceback.format_exc()}",
                "raw_output": "",
                "issues": [],
                "elapsed_sec": round(elapsed, 2)
            })


# ================= 8. 图片标注组件 (修改版：支持拖动) =================

from PyQt6.QtWidgets import (
    QGraphicsView, QGraphicsScene, QGraphicsPixmapItem,
    QGraphicsRectItem, QGraphicsEllipseItem, QGraphicsPathItem,
    QGraphicsTextItem, QGraphicsItem
)
from PyQt6.QtGui import QPainterPath


class AnnotatableImageView(QGraphicsView):
    """
    - 显示图片
    - 支持用户绘制：rect/ellipse/arrow/text/issue_tag
    - 【核心修改】：创建真正的 QGraphicsItem 以支持鼠标拖动调整位置
    """
    annotation_changed = pyqtSignal()

    TOOL_NONE = "none"
    TOOL_RECT = "rect"
    TOOL_ELLIPSE = "ellipse"
    TOOL_ARROW = "arrow"
    TOOL_TEXT = "text"
    TOOL_ISSUE_TAG = "issue_tag"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setScene(QGraphicsScene(self))
        self._pix_item = QGraphicsPixmapItem()
        # 必须设为不可移动，否则拖动标注时可能会误拖动底图
        self._pix_item.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, False)
        self.scene().addItem(self._pix_item)

        self._img_path: Optional[str] = None
        self._base_pix: Optional[QPixmap] = None
        self._base_img_size = (1, 1)

        self._ai_issues: List[Dict[str, Any]] = []
        self._current_issues_data: List[Dict[str, Any]] = []

        self._tool = self.TOOL_NONE
        self._draw_color = "#FF0000"
        self._draw_width = 6

        self._dragging = False
        self._start_img_pt: Optional[QPointF] = None
        self._temp_end_img_pt: Optional[QPointF] = None

        self.setRenderHints(
            QPainter.RenderHint.Antialiasing |
            QPainter.RenderHint.SmoothPixmapTransform
        )
        self.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)

        # 允许框选拖拽
        self.setDragMode(QGraphicsView.DragMode.NoDrag)

    def set_tool(self, tool: str):
        self._tool = tool
        # 如果是浏览模式，允许手型拖动视图；绘图模式则禁用
        if tool == self.TOOL_NONE:
            self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
        else:
            self.setDragMode(QGraphicsView.DragMode.NoDrag)

    def set_image(self, path: str):
        self._img_path = path
        pix = QPixmap(path)
        self._base_pix = pix
        self._pix_item.setPixmap(pix)
        self._base_img_size = (max(1, pix.width()), max(1, pix.height()))
        self.scene().setSceneRect(QRectF(0, 0, pix.width(), pix.height()))
        self.fitInView(self.sceneRect(), Qt.AspectRatioMode.KeepAspectRatio)
        self.viewport().update()

    def set_ai_issues(self, issues: List[Dict[str, Any]]):
        self._ai_issues = issues or []

    def set_current_issues_data(self, issues: List[Dict[str, Any]]):
        self._current_issues_data = issues

    def set_user_annotations(self, ann: List[Dict[str, Any]]):
        """加载数据时，清空当前场景中的标注Item，重新生成可交互Item"""
        # 1. 清除旧的标注 Item (保留底图 _pix_item)
        for item in self.scene().items():
            if item != self._pix_item:
                self.scene().removeItem(item)

        # 2. 重新创建
        if not ann:
            return

        for a in ann:
            self._create_graphics_item_from_data(a)

        self.viewport().update()

    def get_user_annotations(self) -> List[Dict[str, Any]]:
        """
        【核心修改】：导出时，遍历 Scene 中的 Item，获取其当前的真实坐标。
        这样用户拖动后，导出的数据就是拖动后的位置。
        """
        annotations = []
        # 遍历场景中所有 Item
        # 注意：scene.items() 包含所有 item，需要过滤掉底图
        # 为了保持顺序，最好按照 ZValue 排序，或者简单的倒序
        items = self.scene().items(Qt.SortOrder.AscendingOrder)

        for item in items:
            if item == self._pix_item:
                continue

            # 提取数据
            data = item.data(Qt.ItemDataRole.UserRole)
            if not data or not isinstance(data, dict):
                continue

            atype = data.get("type")
            # 获取当前的位置偏移 (用户可能拖动了)
            pos_offset = item.pos()

            # 根据类型重新计算坐标
            if atype in ["rect", "ellipse"]:
                # 原始矩形 + 偏移量
                orig_rect = item.rect()
                # 映射回 Scene 坐标（即图片坐标）
                scene_poly = item.mapToScene(orig_rect)
                scene_rect = scene_poly.boundingRect()
                data["bbox"] = [
                    int(scene_rect.left()), int(scene_rect.top()),
                    int(scene_rect.right()), int(scene_rect.bottom())
                ]

            elif atype == "arrow":
                # 箭头作为一个整体 PathItem，位置就是 pos
                # 简便做法：我们存储箭头创建时的相对路径，导出时加上 pos
                # 但为了兼容 draw_user_annotations，我们需要更新 p1, p2
                # 这是一个简化的处理：只更新整体偏移，不处理变形
                orig_p1 = data.get("orig_p1", [0, 0])
                orig_p2 = data.get("orig_p2", [0, 0])
                data["p1"] = [int(orig_p1[0] + pos_offset.x()), int(orig_p1[1] + pos_offset.y())]
                data["p2"] = [int(orig_p2[0] + pos_offset.x()), int(orig_p2[1] + pos_offset.y())]
                # 清理临时数据
                if "orig_p1" in data: del data["orig_p1"]
                if "orig_p2" in data: del data["orig_p2"]

            elif atype == "text":
                # TextItem 的位置就是 pos
                scene_pos = item.scenePos()
                data["pos"] = [int(scene_pos.x()), int(scene_pos.y())]

            annotations.append(data)

        return annotations

    def clear_annotations(self):
        for item in self.scene().items():
            if item != self._pix_item:
                self.scene().removeItem(item)
        self.annotation_changed.emit()

    def undo(self):
        # 简单的撤销：删除最后添加的一个 Item
        items = [i for i in self.scene().items(Qt.SortOrder.AscendingOrder) if i != self._pix_item]
        if items:
            self.scene().removeItem(items[-1])
            self.annotation_changed.emit()

    def _to_img_point(self, view_pos) -> QPointF:
        sp = self.mapToScene(view_pos)
        # 限制在图片范围内
        x = min(max(sp.x(), 0.0), float(self._base_img_size[0]))
        y = min(max(sp.y(), 0.0), float(self._base_img_size[1]))
        return QPointF(x, y)

    def mousePressEvent(self, event):
        # 如果点击的是已有的可移动 Item，优先让 Qt 处理拖动
        item = self.itemAt(event.position().toPoint())
        if item and item != self._pix_item and self._tool != self.TOOL_NONE:
            # 如果当前在绘图模式，但点到了一个已存在的对象，
            # 此时看需求：是优先选中移动，还是强制画新图？
            # 通常逻辑：按住 Shift 强制画图，否则优先选中。
            # 这里简化：只要选中了Item且Item可移动，就交给父类处理（移动）
            # 除非当前是“绘图”操作开始
            pass

        if event.button() == Qt.MouseButton.LeftButton and self._tool != self.TOOL_NONE:
            # 如果点击处没有可移动图元，或者我们想强制画图
            if not item or item == self._pix_item:
                self._dragging = True
                self._start_img_pt = self._to_img_point(event.position().toPoint())
                self._temp_end_img_pt = self._start_img_pt
                return  # 拦截，不传递给父类（防止 ScrollHandDrag 生效）

        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._dragging:
            self._temp_end_img_pt = self._to_img_point(event.position().toPoint())
            self.viewport().update()  # 触发 drawForeground 画临时框
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._dragging and event.button() == Qt.MouseButton.LeftButton:
            self._dragging = False
            end_pt = self._to_img_point(event.position().toPoint())
            start_pt = self._start_img_pt or end_pt

            # 创建数据结构
            new_data = None

            if self._tool in [self.TOOL_RECT, self.TOOL_ELLIPSE]:
                x1, y1 = start_pt.x(), start_pt.y()
                x2, y2 = end_pt.x(), end_pt.y()
                x1, x2 = sorted([x1, x2])
                y1, y2 = sorted([y1, y2])
                if (x2 - x1) >= 3 and (y2 - y1) >= 3:
                    new_data = {
                        "type": self._tool,
                        "bbox": [int(x1), int(y1), int(x2), int(y2)],
                        "color": self._draw_color,
                        "width": self._draw_width
                    }

            elif self._tool == self.TOOL_ARROW:
                if (abs(end_pt.x() - start_pt.x()) + abs(end_pt.y() - start_pt.y())) >= 3:
                    new_data = {
                        "type": "arrow",
                        "p1": [int(start_pt.x()), int(start_pt.y())],
                        "p2": [int(end_pt.x()), int(end_pt.y())],
                        "color": self._draw_color,
                        "width": self._draw_width
                    }

            elif self._tool == self.TOOL_TEXT:
                text, ok = self._prompt_text()
                if ok and text.strip():
                    new_data = {
                        "type": "text",
                        "pos": [int(end_pt.x()), int(end_pt.y())],
                        "text": text.strip(),
                        "color": self._draw_color,
                        "width": max(2, self._draw_width // 2),
                        "font_size": 28
                    }

            elif self._tool == self.TOOL_ISSUE_TAG:
                if not self._current_issues_data:
                    QMessageBox.warning(self, "提示", "当前图片没有AI识别出的问题，无法引用。")
                else:
                    dlg = IssueSelectionDialog(self, self._current_issues_data)
                    if dlg.exec() == QDialog.DialogCode.Accepted:
                        new_data = {
                            "type": "text",
                            "pos": [int(end_pt.x()), int(end_pt.y())],
                            "text": dlg.selected_text,
                            "color": dlg.selected_color,
                            "width": 4,
                            "font_size": 36
                        }

            # 如果生成了数据，立即转换为 Scene Item
            if new_data:
                self._create_graphics_item_from_data(new_data)
                self.annotation_changed.emit()

            self._start_img_pt = None
            self._temp_end_img_pt = None
            self.viewport().update()
            return

        super().mouseReleaseEvent(event)

        def mouseReleaseEvent(self, event):
            if self._dragging and event.button() == Qt.MouseButton.LeftButton:
                self._dragging = False
                end_pt = self._to_img_point(event.position().toPoint())
                start_pt = self._start_img_pt or end_pt

                # 创建数据结构
                new_data = None

                if self._tool in [self.TOOL_RECT, self.TOOL_ELLIPSE]:
                    x1, y1 = start_pt.x(), start_pt.y()
                    x2, y2 = end_pt.x(), end_pt.y()
                    x1, x2 = sorted([x1, x2])
                    y1, y2 = sorted([y1, y2])
                    if (x2 - x1) >= 3 and (y2 - y1) >= 3:
                        new_data = {
                            "type": self._tool,
                            "bbox": [int(x1), int(y1), int(x2), int(y2)],
                            "color": self._draw_color,
                            "width": self._draw_width
                        }

                elif self._tool == self.TOOL_ARROW:
                    if (abs(end_pt.x() - start_pt.x()) + abs(end_pt.y() - start_pt.y())) >= 3:
                        new_data = {
                            "type": "arrow",
                            "p1": [int(start_pt.x()), int(start_pt.y())],
                            "p2": [int(end_pt.x()), int(end_pt.y())],
                            "color": self._draw_color,
                            "width": self._draw_width
                        }

                elif self._tool == self.TOOL_TEXT:
                    text, ok = self._prompt_text()
                    if ok and text.strip():
                        new_data = {
                            "type": "text",
                            "pos": [int(end_pt.x()), int(end_pt.y())],
                            "text": text.strip(),
                            "color": self._draw_color,
                            "width": max(2, self._draw_width // 2),
                            "font_size": 28
                        }

                elif self._tool == self.TOOL_ISSUE_TAG:
                    if not self._current_issues_data:
                        QMessageBox.warning(self, "提示", "当前图片没有AI识别出的问题，无法引用。")
                    else:
                        dlg = IssueSelectionDialog(self, self._current_issues_data)
                        if dlg.exec() == QDialog.DialogCode.Accepted:
                            new_data = {
                                "type": "text",  # 注意这里：引用标签也是 text 类型
                                "pos": [int(end_pt.x()), int(end_pt.y())],
                                "text": dlg.selected_text,
                                "color": dlg.selected_color,
                                "width": 4,
                                "font_size": 36
                            }

                # 如果生成了数据，立即转换为 Scene Item
                if new_data:
                    self._create_graphics_item_from_data(new_data)
                    self.annotation_changed.emit()

                self._start_img_pt = None
                self._temp_end_img_pt = None
                self.viewport().update()
                return

            super().mouseReleaseEvent(event)

        # =============== 重点修改位置 ===============
        # 必须确保这个函数靠左对齐，与上面的 def mouseReleaseEvent 平级
        # 绝不能缩进在上面的函数里面
        # ==========================================
    def mouseDoubleClickEvent(self, event):
        """
        双击事件：同时支持修改 [手动文字] 和 [引用标签]
        """
         # 1. 获取点击位置
        click_pos = event.position().toPoint()
        sp = self.mapToScene(click_pos)

         # 2. 扩大搜索范围，防止点不准
        search_rect = QRectF(sp.x() - 10, sp.y() - 10, 20, 20)
        items = self.scene().items(search_rect)

        for item in items:
            # 3. 寻找文字图元
             if isinstance(item, QGraphicsTextItem):
                data = item.data(Qt.ItemDataRole.UserRole)

                # 只要 type 是 text，无论是手动输入的还是标签引用的，都进入编辑模式
                if data and isinstance(data, dict) and data.get("type") == "text":

                # 获取旧文本
                    old_text = item.toPlainText()

                # 弹出输入框
                    new_text, ok = self._prompt_text(old_text)

                if ok and new_text.strip():
                     # 更新显示内容（保留微透明背景以维持点击区域）
                    item.setHtml(
                    f"<div style='background-color:rgba(255,255,255,0.01);'>{new_text.strip()}</div>")

                    # 更新底层数据
                    data["text"] = new_text.strip()
                    item.setData(Qt.ItemDataRole.UserRole, data)

                    self.annotation_changed.emit()
                    self.viewport().update()
                    return  # 只要处理了一个文字，就停止处理，防止重叠时触发多次

                super().mouseDoubleClickEvent(event)


    def _create_graphics_item_from_data(self, data: Dict[str, Any]):
        """根据数据字典创建可移动的 QGraphicsItem"""
        t = data.get("type")
        color = QColor(data.get("color", "#FF0000"))
        w = int(data.get("width", 6))

        pen = QPen(color, w)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)

        item = None

        if t == self.TOOL_RECT:
            bbox = data.get("bbox")
            rect = QRectF(bbox[0], bbox[1], bbox[2] - bbox[0], bbox[3] - bbox[1])
            item = QGraphicsRectItem(rect)
            item.setPen(pen)

        elif t == self.TOOL_ELLIPSE:
            bbox = data.get("bbox")
            rect = QRectF(bbox[0], bbox[1], bbox[2] - bbox[0], bbox[3] - bbox[1])
            item = QGraphicsEllipseItem(rect)
            item.setPen(pen)

        elif t == "arrow":
            p1 = data.get("p1")
            p2 = data.get("p2")
            path = QPainterPath()
            start = QPointF(p1[0], p1[1])
            end = QPointF(p2[0], p2[1])
            path.moveTo(start)
            path.lineTo(end)

            # 画箭头头部
            import math
            angle = math.atan2(end.y() - start.y(), end.x() - start.x())
            head_len = w * 4
            head_ang = math.radians(25)

            arrow_p1 = QPointF(end.x() - head_len * math.cos(angle - head_ang),
                               end.y() - head_len * math.sin(angle - head_ang))
            arrow_p2 = QPointF(end.x() - head_len * math.cos(angle + head_ang),
                               end.y() - head_len * math.sin(angle + head_ang))

            # 简单的箭头路径
            path.moveTo(end)
            path.lineTo(arrow_p1)
            path.moveTo(end)
            path.lineTo(arrow_p2)

            item = QGraphicsPathItem(path)
            item.setPen(pen)

            # 存储原始坐标，以便计算偏移
            data["orig_p1"] = p1
            data["orig_p2"] = p2


        elif t == "text":

            text = data.get("text", "")

            pos = data.get("pos")

            item = QGraphicsTextItem(text)

            # 字体设置

            f = QFont()

            f.setPointSize(int(data.get("font_size", 28)))

            f.setBold(True)

            item.setFont(f)

            item.setDefaultTextColor(color)

            item.setPos(pos[0], pos[1])

            # --- 关键修改：增加这三行，让文字块变得“容易被点中” ---

            # 禁用文字内部的编辑模式，防止拦截双击事件

            item.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)

            # 设置一个极其微弱的背景色（透明度为1），肉眼看不见，但会让整个矩形区域可点击

            item.setHtml(f"<div style='background-color:rgba(255,255,255,0.01);'>{text}</div>")

        if item:
            # 【关键】：设置标志，允许鼠标拖动和选中
            item.setFlags(QGraphicsItem.GraphicsItemFlag.ItemIsMovable |
                          QGraphicsItem.GraphicsItemFlag.ItemIsSelectable)
            # 将原始数据绑定到 item，以便导出时知道它是啥
            item.setData(Qt.ItemDataRole.UserRole, data)
            self.scene().addItem(item)

    def _prompt_text(self, default_text="") -> Tuple[str, bool]:
        dlg = QDialog(self)
        dlg.setWindowTitle("输入/修改标注")
        dlg.resize(420, 160)
        layout = QVBoxLayout(dlg)
        edit = QLineEdit()
        edit.setPlaceholderText("例如：钢筋外露 / 临边无防护")

        # 关键：如果有旧文本，先填进去
        if default_text:
            edit.setText(default_text)

        layout.addWidget(edit)
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Cancel | QDialogButtonBox.StandardButton.Ok)
        layout.addWidget(btns)
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)

        # 自动聚焦并全选，方便直接打字覆盖
        edit.setFocus()
        edit.selectAll()

        ok = dlg.exec() == QDialog.DialogCode.Accepted
        return edit.text(), ok

    def drawForeground(self, painter: QPainter, rect: QRectF):
        # 移除了绘制已保存标注的循环，因为现在它们是 Scene 里的 Item 了
        super().drawForeground(painter, rect)

        # 只绘制正在拖拽时的临时预览虚线
        if self._dragging and self._start_img_pt and self._temp_end_img_pt:
            painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
            painter.setPen(QPen(QColor("#00E5FF"), 4, Qt.PenStyle.DashLine))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            s = self._start_img_pt
            e = self._temp_end_img_pt
            x1, x2 = sorted([s.x(), e.x()])
            y1, y2 = sorted([s.y(), e.y()])
            r = QRectF(x1, y1, x2 - x1, y2 - y1)

            if self._tool == self.TOOL_RECT:
                painter.drawRect(r)
            elif self._tool == self.TOOL_ELLIPSE:
                painter.drawEllipse(r)
            elif self._tool == self.TOOL_ARROW:
                painter.drawLine(s, e)


# ================= 新增类：问题快捷选择对话框 =================
class IssueSelectionDialog(QDialog):
    def __init__(self, parent, issues: List[Dict[str, Any]]):
        super().__init__(parent)
        self.setWindowTitle("选择要引用的问题")
        self.resize(500, 300)
        self.selected_text = ""
        self.selected_color = "#FF0000"

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("请选择此标注关联的问题（将自动填入问题描述）："))

        self.list_widget = QListWidget()
        for idx, item in enumerate(issues, 1):
            level = item.get("risk_level", "一般")
            desc = item.get("issue", "未知问题")
            # 构建显示文本
            display_text = f"{idx}. [{level}] {desc}"

            list_item = QListWidgetItem(display_text)
            list_item.setData(Qt.ItemDataRole.UserRole, desc)  # 【修改】存储具体描述
            list_item.setData(Qt.ItemDataRole.UserRole + 1, level)
            self.list_widget.addItem(list_item)

        layout.addWidget(self.list_widget)

        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def accept(self):
        item = self.list_widget.currentItem()
        if item:
            desc = item.data(Qt.ItemDataRole.UserRole)
            level = item.data(Qt.ItemDataRole.UserRole + 1)

            # 【修改】直接使用问题的描述文字，代替序号
            # 简单截断一下过长的描述，防止图片上全是字
            short_desc = desc
            if len(short_desc) > 15:
                short_desc = short_desc[:15]

            self.selected_text = short_desc

            # 根据风险等级决定颜色
            if any(x in str(level) for x in ["严重", "红线"]):
                self.selected_color = "#FF0000"  # 红
            elif any(x in str(level) for x in ["文明"]):
                self.selected_color = "#2196F3"  # 蓝
            else:
                self.selected_color = "#FF8800"  # 橙

        super().accept()


# ================= 9. UI 组件 =================

class IssueEditDialog(QDialog):
    def __init__(self, parent, item: Dict[str, Any]):
        super().__init__(parent)
        self.setWindowTitle("编辑问题")
        self.resize(560, 460)
        self.item = dict(item)

        layout = QVBoxLayout(self)

        form = QFormLayout()
        self.cbo_level = QComboBox()
        self.cbo_level.addItems([
            "严重安全隐患", "一般安全隐患", "严重质量缺陷", "一般质量缺陷", "文明施工问题"
        ])
        if self.item.get("risk_level"):
            idx = self.cbo_level.findText(self.item["risk_level"])
            if idx >= 0:
                self.cbo_level.setCurrentIndex(idx)

        self.txt_issue = QPlainTextEdit()
        self.txt_issue.setPlainText(self.item.get("issue", ""))

        self.txt_reg = QLineEdit()
        self.txt_reg.setText(self.item.get("regulation", ""))

        self.txt_corr = QPlainTextEdit()
        self.txt_corr.setPlainText(self.item.get("correction", ""))

        self.txt_bbox = QLineEdit()
        bbox = self.item.get("bbox")
        self.txt_bbox.setPlaceholderText("例如：100,200,300,380 或留空")
        if bbox:
            self.txt_bbox.setText(",".join([str(x) for x in bbox]))

        form.addRow("风险等级:", self.cbo_level)
        form.addRow("问题描述:", self.txt_issue)
        form.addRow("依据:", self.txt_reg)
        form.addRow("整改建议:", self.txt_corr)
        form.addRow("bbox(可选):", self.txt_bbox)

        layout.addLayout(form)

        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Cancel | QDialogButtonBox.StandardButton.Ok)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def get_value(self) -> Dict[str, Any]:
        bbox_txt = self.txt_bbox.text().strip()
        bbox = None
        if bbox_txt:
            try:
                parts = [int(float(x.strip())) for x in bbox_txt.split(",")]
                if len(parts) == 4:
                    bbox = _normalize_bbox(parts)
            except Exception:
                bbox = None
        return {
            "risk_level": self.cbo_level.currentText().strip(),
            "issue": self.txt_issue.toPlainText().strip(),
            "regulation": self.txt_reg.text().strip(),
            "correction": self.txt_corr.toPlainText().strip(),
            "bbox": bbox,
            "confidence": self.item.get("confidence")
        }


class RiskCard(QFrame):
    edit_requested = pyqtSignal(dict)
    delete_requested = pyqtSignal(dict)

    def __init__(self, item: Dict[str, Any]):
        super().__init__()
        self.item = item
        self.setFrameShape(QFrame.Shape.StyledPanel)

        level = item.get("risk_level", "一般")

        colors = {"红": "#FFE5E5", "橙": "#FFF4E5", "蓝": "#E3F2FD"}
        borders = {"红": "#FF0000", "橙": "#FF8800", "蓝": "#2196F3"}

        if any(x in level for x in ["重大", "严重", "High", "警示", "红线"]):
            bg, bd = colors["红"], borders["红"]
        elif any(x in level for x in ["较大", "一般", "质量", "需整理", "Medium"]):
            bg, bd = colors["橙"], borders["橙"]
        else:
            bg, bd = colors["蓝"], borders["蓝"]

        self.setStyleSheet(
            f"RiskCard {{ background-color: {bg}; border-left: 5px solid {bd}; border-radius: 4px; margin-bottom: 6px; padding: 6px; }}"
        )

        layout = QVBoxLayout(self)
        header = QHBoxLayout()

        header.addWidget(QLabel(f"<b>[{level}]</b>"))

        lbl_issue = QLabel(item.get("issue", ""))
        lbl_issue.setWordWrap(True)
        header.addWidget(lbl_issue, 1)

        btn_edit = QPushButton("编辑")
        btn_edit.setFixedWidth(70)
        btn_edit.clicked.connect(lambda: self.edit_requested.emit(self.item))
        header.addWidget(btn_edit)

        btn_del = QPushButton("删除")
        btn_del.setFixedWidth(70)
        btn_del.clicked.connect(lambda: self.delete_requested.emit(self.item))
        header.addWidget(btn_del)

        layout.addLayout(header)

        bbox = item.get("bbox")
        bbox_text = f"{bbox}" if bbox else "无/未定位"
        layout.addWidget(QLabel(f"依据: {item.get('regulation', '')}"))
        layout.addWidget(QLabel(f"定位 bbox: {bbox_text}"))
        lbl_fix = QLabel(f"建议: {item.get('correction', '')}")
        lbl_fix.setStyleSheet("color: #2E7D32; font-weight: bold;")
        lbl_fix.setWordWrap(True)
        layout.addWidget(lbl_fix)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = ConfigManager.load()
        self.refresh_business_data()

        self.tasks: List[Dict[str, Any]] = []
        self.current_task_id: Optional[str] = None

        self.running_workers: Dict[str, AnalysisWorker] = {}
        self.pending_queue: List[str] = []

        self.total_task = 0
        self.done_task = 0

        self.init_ui()

        self._resize_timer = QTimer(self)
        self._resize_timer.setInterval(200)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.timeout.connect(self._refresh_current_image)

    def refresh_business_data(self):
        self.business_data = self.config.get("business_data", DEFAULT_BUSINESS_DATA)

    def init_ui(self):
        self.setWindowTitle("普洱版纳区域检查报告助手（手动标注版）")
        self.resize(1320, 980)

        toolbar = QToolBar("Main")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        toolbar.addWidget(QLabel("  场景模式: "))
        self.cbo_prompt = QComboBox()
        prompts = self.config.get("prompts", DEFAULT_PROMPTS)
        self.cbo_prompt.addItems(prompts.keys())
        self.cbo_prompt.setCurrentText(self.config.get("last_prompt", list(prompts.keys())[0]))
        self.cbo_prompt.setMinimumWidth(280)
        self.cbo_prompt.currentTextChanged.connect(self.save_prompt_selection)
        toolbar.addWidget(self.cbo_prompt)

        toolbar.addSeparator()

        btn_add = QAction(QIcon(), "➕ 添加图片", self)
        btn_add.triggered.connect(self.add_files)
        toolbar.addAction(btn_add)

        btn_run = QAction(QIcon(), "▶ 开始分析", self)
        btn_run.triggered.connect(self.start_analysis)
        toolbar.addAction(btn_run)

        btn_pause = QAction("⏸ 暂停", self)
        btn_pause.triggered.connect(self.pause_analysis)
        toolbar.addAction(btn_pause)

        btn_clear = QAction("🗑️ 清空队列", self)
        btn_clear.triggered.connect(self.clear_queue)
        toolbar.addAction(btn_clear)

        btn_export_tool = QToolButton()
        btn_export_tool.setText("📄 导出报告 ▼")
        btn_export_tool.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        btn_export_tool.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)

        export_menu = QMenu(self)

        # 1. 检查模板 (对应 检查模板.docx)
        act_report_check = QAction("通用检查报告 (使用 检查模板.docx)", self)
        act_report_check.triggered.connect(lambda: self.export_word("检查模板.docx"))
        export_menu.addAction(act_report_check)

        # 2. 通知单模板 (对应 通知单模板.docx)
        act_report_notice = QAction("整改通知单 (使用 通知单模板.docx)", self)
        act_report_notice.triggered.connect(lambda: self.export_word("通知单模板.docx"))
        export_menu.addAction(act_report_notice)

        # 3. 简报模板 (对应 简报模板.docx)
        act_report_simple = QAction("简报模式 (使用 简报模板.docx)", self)
        act_report_simple.triggered.connect(lambda: self.export_word("简报模板.docx"))
        export_menu.addAction(act_report_simple)

        btn_export_tool.setMenu(export_menu)
        toolbar.addWidget(btn_export_tool) # 添加到工具栏

        empty = QWidget()
        empty.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        toolbar.addWidget(empty)

        btn_setting = QAction("⚙ 设置", self)
        btn_setting.triggered.connect(self.open_settings)
        toolbar.addAction(btn_setting)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 顶部：基础信息
        info_group = QGroupBox("📄 报告基础信息 (数据源可配置)")
        info_group.setFixedHeight(210)
        info_layout = QGridLayout(info_group)
        info_layout.setContentsMargins(10, 10, 10, 10)

        self.input_company = QComboBox()
        self.update_company_combo()
        self.input_company.setEditable(False)

        self.input_project = QComboBox()
        self.input_project.setEditable(False)

        self.input_inspected_unit = QLineEdit()
        self.input_inspected_unit.setPlaceholderText("自动生成，也可手动修改")

        self.input_check_content = QComboBox()
        self.update_check_content_combo()
        self.input_check_content.setEditable(True)

        self.input_area = QLineEdit()
        self.input_area.setPlaceholderText("例如：乡镇或者枢纽、隧洞等（将记忆最近使用）")

        self.input_person = QLineEdit()
        self.input_person.setPlaceholderText("请输入检查人姓名（将记忆）")
        self.input_person.setText(self.config.get("last_check_person", ""))

        self.input_date = QLineEdit()
        self.input_date.setText(datetime.now().strftime("%Y-%m-%d"))

        self.input_deadline = QLineEdit()
        self.input_deadline.setPlaceholderText("例如：2025-12-30")

        quick_box = QHBoxLayout()
        btn_3 = QPushButton("+3天")
        btn_7 = QPushButton("+7天")
        btn_15 = QPushButton("+15天")
        for b in (btn_3, btn_7, btn_15):
            b.setFixedWidth(70)
        btn_3.clicked.connect(lambda: self._set_deadline_days(3))
        btn_7.clicked.connect(lambda: self._set_deadline_days(7))
        btn_15.clicked.connect(lambda: self._set_deadline_days(15))
        quick_box.addWidget(btn_3)
        quick_box.addWidget(btn_7)
        quick_box.addWidget(btn_15)
        quick_box.addStretch(1)
        quick_deadline_widget = QWidget()
        quick_deadline_widget.setLayout(quick_box)

        self.input_group = QLineEdit()
        self.input_group.setPlaceholderText("点位/部位分组（可选，如：隧洞进口段）")

        self.input_company.currentTextChanged.connect(self.on_company_changed)
        if self.input_company.count() > 0:
            self.on_company_changed(self.input_company.currentText())

        info_layout.addWidget(QLabel("项目公司名称:"), 0, 0)
        info_layout.addWidget(self.input_company, 0, 1)
        info_layout.addWidget(QLabel("检查项目名称:"), 0, 2)
        info_layout.addWidget(self.input_project, 0, 3)

        info_layout.addWidget(QLabel("被检查单位:"), 1, 0)
        info_layout.addWidget(self.input_inspected_unit, 1, 1)
        info_layout.addWidget(QLabel("检查内容:"), 1, 2)
        info_layout.addWidget(self.input_check_content, 1, 3)

        info_layout.addWidget(QLabel("检查部位:"), 2, 0)
        info_layout.addWidget(self.input_area, 2, 1)
        info_layout.addWidget(QLabel("检查人员:"), 2, 2)
        info_layout.addWidget(self.input_person, 2, 3)

        info_layout.addWidget(QLabel("检查日期:"), 3, 0)
        info_layout.addWidget(self.input_date, 3, 1)
        info_layout.addWidget(QLabel("整改期限:"), 3, 2)
        info_layout.addWidget(self.input_deadline, 3, 3)

        info_layout.addWidget(QLabel("期限快捷:"), 4, 2)
        info_layout.addWidget(quick_deadline_widget, 4, 3)
        info_layout.addWidget(QLabel("点位分组(可选):"), 4, 0)
        info_layout.addWidget(self.input_group, 4, 1)

        main_layout.addWidget(info_group)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        # 左侧
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)

        self.lbl_count = QLabel(f"待审队列 (0/{MAX_IMAGES})")
        left_layout.addWidget(self.lbl_count)

        self.list_widget = QListWidget()
        self.list_widget.itemClicked.connect(self.on_item_clicked)
        left_layout.addWidget(self.list_widget)

        batch_box = QHBoxLayout()
        btn_apply_group = QPushButton("批量设点位")
        btn_apply_group.clicked.connect(self.apply_group_to_all_tasks)
        btn_retry_error = QPushButton("重试失败")
        btn_retry_error.clicked.connect(self.retry_errors)
        batch_box.addWidget(btn_apply_group)
        batch_box.addWidget(btn_retry_error)
        left_layout.addLayout(batch_box)

        # 右侧
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)

        # === 优化后的标注工具栏 (拆分为两行) ===
        self.btn_tool_none = QPushButton("浏览")
        self.btn_tool_rect = QPushButton("框")
        self.btn_tool_ellipse = QPushButton("圈")
        self.btn_tool_arrow = QPushButton("箭头")
        self.btn_tool_text = QPushButton("文字")
        self.btn_tool_tag = QPushButton("🏷️引用问题")
        self.btn_tool_tag.setStyleSheet("color: blue; font-weight: bold;")

        self.btn_undo = QPushButton("撤销")
        self.btn_clear_anno = QPushButton("清空")
        self.btn_save_marked = QPushButton("保存截图")

        all_btns = [
            self.btn_tool_none, self.btn_tool_rect, self.btn_tool_ellipse,
            self.btn_tool_arrow, self.btn_tool_text, self.btn_tool_tag,
            self.btn_undo, self.btn_clear_anno, self.btn_save_marked
        ]
        for b in all_btns:
            b.setMinimumHeight(28)
            b.setFixedWidth(65)

        self.btn_tool_tag.setFixedWidth(80)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("绘图:"))
        row1.addWidget(self.btn_tool_none)
        row1.addWidget(self.btn_tool_rect)
        row1.addWidget(self.btn_tool_ellipse)
        row1.addWidget(self.btn_tool_arrow)
        row1.addWidget(self.btn_tool_text)
        row1.addWidget(self.btn_tool_tag)
        row1.addStretch()

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("操作:"))
        row2.addWidget(self.btn_undo)
        row2.addWidget(self.btn_clear_anno)
        row2.addWidget(self.btn_save_marked)
        row2.addStretch()

        tool_container = QWidget()
        tool_layout = QVBoxLayout(tool_container)
        tool_layout.setContentsMargins(0, 5, 0, 5)
        tool_layout.setSpacing(2)
        tool_layout.addLayout(row1)
        tool_layout.addLayout(row2)

        right_layout.addWidget(tool_container)

        self.image_view = AnnotatableImageView()
        self.image_view.setMinimumHeight(420)
        self.image_view.annotation_changed.connect(self._on_annotation_changed)
        right_layout.addWidget(self.image_view, 2)

        self.txt_raw = QPlainTextEdit()
        self.txt_raw.setReadOnly(True)
        self.txt_raw.setPlaceholderText("模型原始输出（解析失败/复核时查看）")
        self.txt_raw.setMaximumHeight(160)

        self.result_container = QWidget()
        self.result_layout = QVBoxLayout(self.result_container)
        self.result_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self.result_container)
        right_layout.addWidget(scroll, 3)

        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([380, 940])
        main_layout.addWidget(splitter)

        self.status_bar = self.statusBar()
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedWidth(240)
        self.status_bar.addPermanentWidget(self.progress_bar)

        self.btn_tool_none.clicked.connect(lambda: self._set_tool(AnnotatableImageView.TOOL_NONE))
        self.btn_tool_rect.clicked.connect(lambda: self._set_tool(AnnotatableImageView.TOOL_RECT))
        self.btn_tool_ellipse.clicked.connect(lambda: self._set_tool(AnnotatableImageView.TOOL_ELLIPSE))
        self.btn_tool_arrow.clicked.connect(lambda: self._set_tool(AnnotatableImageView.TOOL_ARROW))
        self.btn_tool_text.clicked.connect(lambda: self._set_tool(AnnotatableImageView.TOOL_TEXT))
        self.btn_tool_tag.clicked.connect(lambda: self._set_tool(AnnotatableImageView.TOOL_ISSUE_TAG))
        self.btn_undo.clicked.connect(self._undo_annotation)
        self.btn_clear_anno.clicked.connect(self._clear_annotation)
        self.btn_save_marked.clicked.connect(self._save_marked_for_current_task)

    def _set_tool(self, tool: str):
        self.image_view.set_tool(tool)
        self.status_bar.showMessage(f"当前标注工具：{tool}")

    def _undo_annotation(self):
        self.image_view.undo()

    def _clear_annotation(self):
        self.image_view.clear_annotations()

    def _on_annotation_changed(self):
        task = self._current_task()
        if not task:
            return
        task["annotations"] = self.image_view.get_user_annotations()

    def _current_task(self) -> Optional[Dict[str, Any]]:
        if not self.current_task_id:
            return None
        return next((t for t in self.tasks if t['id'] == self.current_task_id), None)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._resize_timer.start()

    def _refresh_current_image(self):
        pass

    def update_company_combo(self):
        current_text = self.input_company.currentText() if hasattr(self, "input_company") else ""
        if hasattr(self, "input_company"):
            self.input_company.blockSignals(True)
            self.input_company.clear()
            company_map = self.business_data.get("company_project_map", {})
            self.input_company.addItems(company_map.keys())
            index = self.input_company.findText(current_text)
            if index >= 0:
                self.input_company.setCurrentIndex(index)
            elif self.input_company.count() > 0:
                self.input_company.setCurrentIndex(0)
            self.input_company.blockSignals(False)

    def update_check_content_combo(self):
        current_text = self.input_check_content.currentText() if hasattr(self, "input_check_content") else ""
        if hasattr(self, "input_check_content"):
            self.input_check_content.clear()
            check_options = self.business_data.get("check_content_options", [])
            self.input_check_content.addItems(check_options)
            self.input_check_content.setEditText(current_text)

    def on_company_changed(self, company_name):
        self.input_project.clear()
        comp_proj_map = self.business_data.get("company_project_map", {})
        projects = comp_proj_map.get(company_name, [])
        self.input_project.addItems(projects)
        if projects:
            self.input_project.setCurrentIndex(0)

        comp_unit_map = self.business_data.get("company_unit_map", {})
        unit_name = comp_unit_map.get(company_name, "")
        self.input_inspected_unit.setText(unit_name)

    def save_prompt_selection(self, text):
        if not text:
            return
        self.config["last_prompt"] = text
        ConfigManager.save(self.config)

    def _set_deadline_days(self, days: int):
        try:
            base = datetime.strptime(self.input_date.text().strip(), "%Y-%m-%d")
        except Exception:
            base = datetime.now()
        target = base + timedelta(days=days)
        self.input_deadline.setText(target.strftime("%Y-%m-%d"))

    def add_files(self):
        current_count = len(self.tasks)
        if current_count >= MAX_IMAGES:
            QMessageBox.warning(self, "数量限制",
                                f"为保证运行稳定，单次排查请控制在 {MAX_IMAGES} 张图片以内。\n建议先清空队列。")
            return

        remaining = MAX_IMAGES - current_count
        paths, _ = QFileDialog.getOpenFileNames(self, f"选择图片 (还能选 {remaining} 张)", "",
                                                "Images (*.jpg *.png *.jpeg)")
        if not paths:
            return

        if len(paths) > remaining:
            QMessageBox.warning(self, "超限提示", f"你选择了 {len(paths)} 张，自动截取前 {remaining} 张。")
            paths = paths[:remaining]

        default_group = self.input_group.text().strip() or None

        for path in paths:
            if any(t['path'] == path for t in self.tasks):
                continue
            task_id = str(time.time()) + os.path.basename(path)

            task = {
                "id": task_id,
                "path": path,
                "name": os.path.basename(path),
                "status": "waiting",
                "issues": [],
                "edited_issues": None,
                "raw_output": "",
                "error": None,
                "elapsed_sec": None,
                "meta": {"group": default_group},
                "annotations": [],
                "export_image_path": None
            }
            self.tasks.append(task)

            item = QListWidgetItem(os.path.basename(path))
            item.setData(Qt.ItemDataRole.UserRole, task_id)
            self.list_widget.addItem(item)

        self.lbl_count.setText(f"待审队列 ({len(self.tasks)}/{MAX_IMAGES})")

    def clear_queue(self):
        if any(t['status'] == 'analyzing' for t in self.tasks) or self.running_workers:
            QMessageBox.warning(self, "警告", "任务正在分析中，请暂停/等待完成后再清空！")
            return
        reply = QMessageBox.question(
            self, '确认', '确定要清空所有待审任务吗？',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.tasks.clear()
            self.pending_queue.clear()
            self.running_workers.clear()
            self.list_widget.clear()
            self.lbl_count.setText(f"待审队列 (0/{MAX_IMAGES})")
            self.current_task_id = None
            self.txt_raw.clear()
            self.image_view.scene().clear()
            self.image_view = AnnotatableImageView()
            while self.result_layout.count():
                child = self.result_layout.takeAt(0)
                if child.widget():
                    child.widget().deleteLater()
            self.status_bar.showMessage("队列已清空")

    def pause_analysis(self):
        self.pending_queue.clear()
        for t in self.tasks:
            if t["status"] == "queued":
                t["status"] = "waiting"
                self.update_list_color(t["id"], "#000000")
        self.status_bar.showMessage("已暂停：未开始的任务已取消排队（进行中的仍会完成）")

    def apply_group_to_all_tasks(self):
        group = self.input_group.text().strip()
        if not group:
            QMessageBox.information(self, "提示", "请先填写“点位分组(可选)”再执行批量设置。")
            return
        for t in self.tasks:
            if "meta" not in t:
                t["meta"] = {}
            t["meta"]["group"] = group
        self.status_bar.showMessage(f"已批量设置点位：{group}")
        if self.current_task_id:
            task = self._current_task()
            if task:
                self.render_result(task)

    def retry_errors(self):
        error_tasks = [t for t in self.tasks if t["status"] == "error"]
        if not error_tasks:
            self.status_bar.showMessage("没有失败任务可重试")
            return
        for t in error_tasks:
            t["status"] = "waiting"
            t["error"] = None
            self.update_list_color(t["id"], "#000000")
        self.status_bar.showMessage(f"已重置 {len(error_tasks)} 个失败任务为待分析")
        self.start_analysis()

    def _remember_fields(self):
        person = self.input_person.text().strip()
        if person:
            self.config["last_check_person"] = person

        area = self.input_area.text().strip()
        if area:
            recent = self.config.get("recent_check_areas", []) or []
            if area in recent:
                recent.remove(area)
            recent.insert(0, area)
            self.config["recent_check_areas"] = recent[:20]

        ConfigManager.save(self.config)

    def start_analysis(self):
        if not self.config.get("api_key"):
            QMessageBox.warning(self, "缺 Key", "请在右上角设置中填写 API Key")
            return

        self._remember_fields()

        waiting = [t for t in self.tasks if t['status'] in ['waiting', 'error']]
        if not waiting:
            self.status_bar.showMessage("没有待处理的任务")
            return

        for t in waiting:
            if t["id"] not in self.pending_queue and t["id"] not in self.running_workers:
                self.pending_queue.append(t["id"])
                t["status"] = "queued"
                self.update_list_color(t["id"], "#444444")

        self.progress_bar.setVisible(True)
        self.total_task = len([t for t in self.tasks if t["status"] in ["queued", "analyzing"]]) + len(
            self.running_workers)
        self.done_task = len([t for t in self.tasks if t["status"] == "done"])

        self._kick_scheduler()

    def _kick_scheduler(self):
        max_conc = int(self.config.get("max_concurrency", 3))
        while len(self.running_workers) < max_conc and self.pending_queue:
            task_id = self.pending_queue.pop(0)
            task = next((t for t in self.tasks if t['id'] == task_id), None)
            if not task:
                continue

            selected_template_name = self.cbo_prompt.currentText()
            prompts_dict = self.config.get("prompts", DEFAULT_PROMPTS)
            prompt_content = prompts_dict.get(selected_template_name, list(DEFAULT_PROMPTS.values())[0])

            task["status"] = "analyzing"
            task["error"] = None
            task["raw_output"] = ""
            task["issues"] = []
            task["edited_issues"] = None
            task["export_image_path"] = None

            self.update_list_color(task_id, "#0000FF")
            worker = AnalysisWorker(task, self.config, prompt_content)
            worker.finished.connect(self.on_worker_done)
            self.running_workers[task_id] = worker
            worker.start()

        total = max(1, self.total_task)
        done = len([t for t in self.tasks if t["status"] == "done"])
        self.progress_bar.setValue(int(done / total * 100))

        if not self.running_workers and not self.pending_queue:
            self.status_bar.showMessage("✅ 队列分析完成")
            self.progress_bar.setValue(100)

    def on_worker_done(self, task_id: str, result: dict):
        task = next((t for t in self.tasks if t['id'] == task_id), None)
        if task:
            task["raw_output"] = result.get("raw_output", "") or ""
            task["elapsed_sec"] = result.get("elapsed_sec")
            if result.get("ok"):
                task['status'] = 'done'
                task['issues'] = result.get("issues", []) or []
                task["error"] = None
                self.update_list_color(task_id, "#008000")
            else:
                task['status'] = 'error'
                task['issues'] = []
                task["error"] = result.get("error") or "未知错误"
                self.update_list_color(task_id, "#FF0000")

            if self.current_task_id == task_id:
                self.render_result(task)

        if task_id in self.running_workers:
            self.running_workers.pop(task_id, None)

        self._kick_scheduler()

    def render_result(self, task: dict):
        while self.result_layout.count():
            child = self.result_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

        if os.path.exists(task.get("path", "")):
            self.image_view.set_image(task["path"])

        # 即使 AI 有 issues，也不再自动显示，但数据需要传进去给“引用问题”功能用
        issues = task.get("edited_issues") if task.get("edited_issues") is not None else task.get("issues", [])
        self.image_view.set_ai_issues(issues)
        self.image_view.set_current_issues_data(issues)
        self.image_view.set_user_annotations(task.get("annotations", []) or [])

        if task['status'] == 'analyzing':
            self.result_layout.addWidget(QLabel("正在智能分析中（准确性优先，可能稍慢）..."))
            return
        if task['status'] == 'queued':
            self.result_layout.addWidget(QLabel("已加入队列，等待分析..."))
            return
        if task['status'] == 'error':
            msg = task.get("error") or "未知错误"
            lbl = QLabel(f"❌ 分析/解析失败：{msg}\n\n你可以点击“重试失败”。")
            lbl.setWordWrap(True)
            self.result_layout.addWidget(lbl)
            return

        if task['status'] == 'done':
            if not issues:
                self.result_layout.addWidget(QLabel("✅ 未发现明显隐患或改进项（或模型输出为空）"))
                return

            for item in issues:
                card = RiskCard(item)
                card.edit_requested.connect(self.edit_issue)
                card.delete_requested.connect(self.delete_issue)
                self.result_layout.addWidget(card)

            tip = QLabel("提示：已关闭自动画框。请使用“绘图”工具手动圈出重点，使用“引用问题”按钮快速添加文字描述。")
            tip.setWordWrap(True)
            self.result_layout.addWidget(tip)

            if os.path.exists(task.get("path", "")):
                self.image_view.set_image(task["path"])

    def edit_issue(self, item: Dict[str, Any]):
        task = self._current_task()
        if not task or task.get("status") != "done":
            return

        issues = task.get("edited_issues") if task.get("edited_issues") is not None else (task.get("issues") or [])
        dlg = IssueEditDialog(self, item)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            new_item = dlg.get_value()
            if task.get("edited_issues") is None:
                task["edited_issues"] = [dict(x) for x in issues]

            replaced = False
            for i, x in enumerate(task["edited_issues"]):
                if x is item or x == item:
                    task["edited_issues"][i] = new_item
                    replaced = True
                    break
            if not replaced:
                task["edited_issues"].append(new_item)

            task["export_image_path"] = None
            self.render_result(task)

    def delete_issue(self, item: Dict[str, Any]):
        task = self._current_task()
        if not task or task.get("status") != "done":
            return

        issues = task.get("edited_issues") if task.get("edited_issues") is not None else (task.get("issues") or [])
        if task.get("edited_issues") is None:
            task["edited_issues"] = [dict(x) for x in issues]

        task["edited_issues"] = [x for x in task["edited_issues"] if x != item]
        task["export_image_path"] = None
        self.render_result(task)

    def on_item_clicked(self, item):
        task_id = item.data(Qt.ItemDataRole.UserRole)
        self.current_task_id = task_id
        task = next((t for t in self.tasks if t['id'] == task_id), None)
        if not task:
            return
        self.render_result(task)

    def update_list_color(self, task_id, color):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == task_id:
                item.setForeground(QColor(color))

    def _save_marked_for_current_task(self):
        task = self._current_task()
        if not task:
            return
        if not os.path.exists(task.get("path", "")):
            QMessageBox.warning(self, "失败", "当前图片不存在")
            return

        issues = task.get("edited_issues") if task.get("edited_issues") is not None else task.get("issues", [])
        anns = task.get("annotations", []) or []

        ensure_export_dir()
        base_name = os.path.splitext(os.path.basename(task["path"]))[0]
        out_path = os.path.join(EXPORT_IMG_DIR, f"{base_name}_{task['id']}.png")

        ok = build_export_marked_image(task["path"], issues, anns, out_path)
        if not ok:
            QMessageBox.warning(self, "失败", "生成带标注图片失败（图片格式或路径异常）")
            return

        task["export_image_path"] = out_path
        QMessageBox.information(self, "成功", f"已生成带标注图片：\n{out_path}")

    def export_word(self, template_name):
        if not self.tasks:
            QMessageBox.warning(self, "提示", "队列为空，无法导出。")
            return

        if not os.path.exists(template_name):
            reply = QMessageBox.warning(
                self,
                "模板缺失警告",
                f"未在程序目录下找到文件：【{template_name}】\n\n"
                f"1. 请确保该 Word 模板文件已放入程序运行目录。\n"
                f"2. 点击【Yes】将强制使用“空白格式”生成报告。\n"
                f"3. 点击【No】取消导出以检查文件。",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return
            # 如果选 Yes，后续 WordReportGenerator 会检测不到文件从而创建空白文档
            # ================= 修改结束 =================

        abs_export_dir = os.path.abspath(ensure_export_dir())

        count_processed = 0
        for t in self.tasks:
            # ... (这部分代码保持不变，处理图片导出的逻辑) ...
            has_issues = (t.get("edited_issues") is not None) or (bool(t.get("issues")))
            has_anns = bool(t.get("annotations"))
            if not has_issues and not has_anns:
                continue
            if not os.path.exists(t.get("path", "")):
                continue
            issues = t.get("edited_issues") if t.get("edited_issues") is not None else t.get("issues", [])
            anns = t.get("annotations", []) or []
            base_name = os.path.splitext(os.path.basename(t["path"]))[0]
            safe_base_name = "".join([c for c in base_name if c.isalnum() or c in (' ', '_', '-')]).strip()
            safe_id = str(t['id'])[-6:]
            out_filename = f"{safe_base_name}_{safe_id}.png"
            out_path = os.path.join(abs_export_dir, out_filename)
            ok = build_export_marked_image(t["path"], issues, anns, out_path)
            if ok:
                t["export_image_path"] = out_path
                count_processed += 1
            else:
                t["export_image_path"] = None

        current_project_name = self.input_project.currentText()
        overview_map = self.business_data.get("project_overview_map", {})
        overview_text = overview_map.get(current_project_name, "暂无该项目的详细概况信息。")

        project_info = {
            "project_company": self.input_company.currentText(),
            "project_name": current_project_name,
            "project_overview": overview_text,
            "inspected_unit": self.input_inspected_unit.text().strip(),
            "check_content": self.input_check_content.currentText().strip(),
            "check_area": self.input_area.text().strip(),
            "rectification_deadline": self.input_deadline.text().strip(),
            "check_person": self.input_person.text().strip(),
            "check_date": self.input_date.text().strip()
        }

        # 生成默认文件名
        current_time_str = datetime.now().strftime('%Y%m%d_%H%M%S')
        prefix = project_info['project_name'] if project_info['project_name'] else "检查报告"

        # 根据模板名称生成更有意义的文件名后缀
        file_suffix = "报告"
        if "通知单" in template_name:
            file_suffix = "通知单"
        elif "简报" in template_name:
            file_suffix = "简报"
        elif "检查" in template_name:
            file_suffix = "检查报告"

        default_name = f"{prefix}_{file_suffix}_{current_time_str}.docx"

        path, _ = QFileDialog.getSaveFileName(self, "保存报告", default_name, "Word Files (*.docx)")
        if not path:
            return

        try:
            WordReportGenerator.generate(self.tasks, path, project_info, template_path=template_name)
            QMessageBox.information(self, "成功",
                                    f"报告已生成！\n模板：{template_name}\n路径：{path}\n\n已包含 {count_processed} 张标注插图。")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"生成报告时发生错误：\n{str(e)}\n{traceback.format_exc()}")

    def open_settings(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("系统高级设置")
        dlg.resize(760, 650)

        tabs = QTabWidget()

        tab_conn = QWidget()
        layout_conn = QFormLayout(tab_conn)
        provider_presets = self.config.get("provider_presets", DEFAULT_PROVIDER_PRESETS)

        cbo_provider = QComboBox()
        cbo_provider.addItems(provider_presets.keys())
        curr_prov = self.config.get("current_provider")
        if curr_prov not in provider_presets:
            curr_prov = list(provider_presets.keys())[0]
        cbo_provider.setCurrentText(curr_prov)

        txt_base_url = QLineEdit()
        txt_model = QLineEdit()
        txt_key = QLineEdit(self.config.get("api_key"))
        txt_key.setEchoMode(QLineEdit.EchoMode.Password)

        def on_provider_change(text):
            preset = provider_presets.get(text, {})
            if text == "自定义 (Custom)":
                custom_saved = self.config.get("custom_provider_settings", {})
                txt_base_url.setText(custom_saved.get("base_url", ""))
                txt_model.setText(custom_saved.get("model", ""))
                txt_base_url.setReadOnly(False)
                txt_model.setReadOnly(False)
            else:
                txt_base_url.setText(preset.get("base_url", ""))
                txt_model.setText(preset.get("model", ""))
                txt_base_url.setReadOnly(False)
                txt_model.setReadOnly(False)

        cbo_provider.currentTextChanged.connect(on_provider_change)
        on_provider_change(cbo_provider.currentText())

        layout_conn.addRow("模型厂商:", cbo_provider)
        layout_conn.addRow("Base URL:", txt_base_url)
        layout_conn.addRow("模型名称:", txt_model)
        layout_conn.addRow("API Key:", txt_key)

        sp_conc = QSpinBox()
        sp_conc.setRange(1, 10)
        sp_conc.setValue(int(self.config.get("max_concurrency", 3)))

        sp_retry = QSpinBox()
        sp_retry.setRange(0, 5)
        sp_retry.setValue(int(self.config.get("max_retries", 2)))

        sp_temp = QLineEdit(str(self.config.get("temperature", 0.1)))

        layout_conn.addRow("最大并发(建议2~3):", sp_conc)
        layout_conn.addRow("自动重试次数:", sp_retry)
        layout_conn.addRow("temperature(越低越稳):", sp_temp)

        tabs.addTab(tab_conn, "🔌 连接设置")

        tab_prompt = QWidget()
        layout_prompt = QVBoxLayout(tab_prompt)
        local_prompts = self.config.get("prompts", DEFAULT_PROMPTS).copy()
        cbo_template = QComboBox()
        cbo_template.addItems(local_prompts.keys())
        txt_prompt_edit = QTextEdit()
        self._temp_last_selected_prompt = cbo_template.currentText()

        def load_prompt(name):
            txt_prompt_edit.setText(local_prompts.get(name, ""))
            self._temp_last_selected_prompt = name

        def save_prompt_to_mem():
            if self._temp_last_selected_prompt in local_prompts:
                local_prompts[self._temp_last_selected_prompt] = txt_prompt_edit.toPlainText()

        cbo_template.currentTextChanged.connect(lambda n: (save_prompt_to_mem(), load_prompt(n)))
        if self._temp_last_selected_prompt:
            load_prompt(self._temp_last_selected_prompt)

        layout_prompt.addWidget(QLabel("选择模板进行编辑:"))
        layout_prompt.addWidget(cbo_template)
        layout_prompt.addWidget(txt_prompt_edit)
        tabs.addTab(tab_prompt, "📝 提示词编辑")

        tab_data = QWidget()
        layout_data = QVBoxLayout(tab_data)

        lbl_info = QLabel(
            "此处配置公司名称、项目名称、被检单位及项目概况。\n请保持 JSON 格式正确（注意双引号和逗号）。修改后点击保存即可生效。"
        )
        lbl_info.setWordWrap(True)
        txt_data_edit = QTextEdit()

        current_biz_data = self.config.get("business_data", DEFAULT_BUSINESS_DATA)
        txt_data_edit.setText(json.dumps(current_biz_data, indent=4, ensure_ascii=False))

        layout_data.addWidget(lbl_info)
        layout_data.addWidget(txt_data_edit)
        tabs.addTab(tab_data, "📊 业务数据配置")

        tab_diag = QWidget()
        layout_diag = QFormLayout(tab_diag)
        lbl_person = QLabel(self.config.get("last_check_person", ""))
        lbl_areas = QPlainTextEdit()
        lbl_areas.setReadOnly(True)
        lbl_areas.setPlainText("\n".join(self.config.get("recent_check_areas", []) or []))
        layout_diag.addRow("最近检查人员:", lbl_person)
        layout_diag.addRow("最近检查部位(Top20):", lbl_areas)
        tabs.addTab(tab_diag, "🧰 诊断")

        btn_box = QHBoxLayout()
        btn_save = QPushButton("保存所有配置")
        btn_save.setMinimumHeight(40)
        btn_save.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; border-radius: 4px;")
        btn_cancel = QPushButton("取消")

        def save_all():
            try:
                save_prompt_to_mem()

                raw_json = txt_data_edit.toPlainText()
                new_biz_data = json.loads(raw_json)

                self.config["current_provider"] = cbo_provider.currentText()
                self.config["api_key"] = txt_key.text().strip()
                self.config["prompts"] = local_prompts
                self.config["business_data"] = new_biz_data

                self.config["max_concurrency"] = int(sp_conc.value())
                self.config["max_retries"] = int(sp_retry.value())
                try:
                    self.config["temperature"] = float(sp_temp.text().strip())
                except Exception:
                    self.config["temperature"] = 0.1

                if cbo_provider.currentText() == "自定义 (Custom)":
                    self.config["custom_provider_settings"] = {
                        "base_url": txt_base_url.text().strip(),
                        "model": txt_model.text().strip()
                    }

                ConfigManager.save(self.config)

                self.refresh_business_data()
                self.update_company_combo()
                self.update_check_content_combo()
                self.on_company_changed(self.input_company.currentText())

                self.cbo_prompt.blockSignals(True)
                curr = self.cbo_prompt.currentText()
                self.cbo_prompt.clear()
                self.cbo_prompt.addItems(self.config["prompts"].keys())
                if curr in self.config["prompts"]:
                    self.cbo_prompt.setCurrentText(curr)
                self.cbo_prompt.blockSignals(False)

                dlg.accept()
                self.status_bar.showMessage("✅ 配置已保存")

            except json.JSONDecodeError as e:
                QMessageBox.critical(dlg, "格式错误", f"业务数据 JSON 格式有误，请检查:\n{e}")
            except Exception as e:
                QMessageBox.critical(dlg, "保存失败", f"错误信息: {str(e)}")

        btn_save.clicked.connect(save_all)
        btn_cancel.clicked.connect(dlg.reject)

        btn_box.addStretch()
        btn_box.addWidget(btn_cancel)
        btn_box.addWidget(btn_save)

        layout = QVBoxLayout(dlg)
        layout.addWidget(tabs)
        layout.addLayout(btn_box)
        dlg.exec()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    w = MainWindow()
    w.show()
    sys.exit(app.exec())
