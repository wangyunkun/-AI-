import sys
import ssl  # 必须放在最前面！
import os
import json
import base64
import time
import re
import traceback
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

# 设置环境变量防止冲突
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"
os.environ["QT_API"] = "pyqt6"

# 必须在 PyQt6 之前导入 OpenAI
import httpx
from openai import OpenAI

# Word 库
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# ================= PyQt6 完整导入区 (确保无误) =================

# 1. 核心常量与工具 (Qt, QBuffer 等在这里)
from PyQt6.QtCore import (
    Qt, QThread, pyqtSignal, QTimer, QPointF, QRectF,
    QBuffer, QByteArray, QIODevice
)

# 2. GUI 绘图组件 (QImage, QPixmap, QColor 等在这里)
from PyQt6.QtGui import (
    QPixmap, QIcon, QColor, QAction, QPainter, QPen, QBrush, QFont,
    QImage, QPainterPath
)

# 3. 窗口控件 (这里不能有 Qt !)
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QListWidget, QListWidgetItem, QSplitter,
    QScrollArea, QFrame, QFileDialog, QProgressBar, QMessageBox,
    QDialog, QFormLayout, QLineEdit, QComboBox, QToolBar,
    QSizePolicy, QTabWidget, QTextEdit, QGroupBox, QGridLayout,
    QSpinBox, QPlainTextEdit, QDialogButtonBox,
    QToolButton, QMenu, QInputDialog
)

# 4. 图形视图组件
from PyQt6.QtWidgets import (
    QGraphicsView, QGraphicsScene, QGraphicsPixmapItem,
    QGraphicsRectItem, QGraphicsEllipseItem, QGraphicsPathItem,
    QGraphicsTextItem, QGraphicsItem
)

# ================= 5. 全局配置常量 =================

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
    """
    增强版 JSON 修复工具：自动补全丢失的逗号、引号，修复 Python 风格的 None/True 等。
    """
    if not s:
        return "[]"

    # 1. 预处理：移除 Markdown 标记和首尾空白
    s = re.sub(r"^```json", "", s, flags=re.MULTILINE | re.IGNORECASE)
    s = re.sub(r"^```", "", s, flags=re.MULTILINE)
    s = re.sub(r"```$", "", s, flags=re.MULTILINE)
    s = s.strip()

    # 2. 提取数组部分 (提取最外层的 [])
    start = s.find("[")
    end = s.rfind("]")
    if start != -1 and end != -1:
        s = s[start:end + 1]

    # 3. 基础字符清洗 (将 Python 格式转为 JSON 标准格式)
    s = s.replace("'", '"')  # 单引号转双引号
    s = s.replace("None", "null")  # Python None -> null
    s = s.replace("True", "true")  # Python True -> true
    s = s.replace("False", "false")
    s = s.replace("\ufeff", "")
    s = s.replace("“", "\"").replace("”", "\"")  # 中文引号修正

    # ================= 4. 强力逗号补全 (通用逻辑) =================

    # 场景 A: 对象/数组之间缺逗号 (例如 } { -> }, { )
    s = re.sub(r"}\s*{", "}, {", s)
    s = re.sub(r"]\s*\[", "], [", s)

    # 场景 B: 字段之间缺逗号 (通用匹配)
    # 逻辑：如果一个值结束了，后面跟着一个引号(新Key的开始)，且中间没有逗号，则强制补逗号。
    # [0-9}\]\"el] 匹配值的结尾字符：数字, }, ], ", e(true/false), l(null)
    # \s+ 匹配中间的空白
    # (?=") 预测后面跟着一个引号
    s = re.sub(r'([0-9}\]\"el])\s+(?=")', r'\1, ', s)

    # 场景 C: 数组内部数字缺逗号 (针对 bbox: [10 20 30])
    def fix_array_spaces(match):
        txt = match.group(1)
        # 将两个数字之间的空格替换为逗号
        return "[" + re.sub(r"(\d)\s+(\d)", r"\1, \2", txt) + "]"

    # 仅修复看起来像数值数组的内容
    s = re.sub(r"\[([\d\s\.-]+)\]", fix_array_spaces, s)

    # ============================================================

    # 5. 清理多余逗号 (例如 ", }" -> "}")
    s = re.sub(r",\s*([}\]])", r"\1", s)

    # 6. 移除不可见控制字符 (防止解析器报错)
    s = re.sub(r'[\x00-\x1f\x7f]', ' ', s)

    return s


def _normalize_bbox(b: Any) -> Optional[List[int]]:
    if b is None:
        return None
    if not isinstance(b, (list, tuple)) or len(b) != 4:
        return None
    try:
        # 1. 强制数值转换，防止字符串混入
        coords = [float(v) for v in b]

        # 2. 【核心修复】安全限制坐标范围
        # Qt 的绘图坐标如果超过 32767 (short) 或 INT_MAX 都有可能导致底层崩溃
        # 这里限制在 -10000 到 100000 之间，足够容纳绝大多数图片，同时防止溢出
        SAFE_MIN = -10000
        SAFE_MAX = 100000

        cleaned = []
        for val in coords:
            if val < SAFE_MIN: val = SAFE_MIN
            if val > SAFE_MAX: val = SAFE_MAX
            cleaned.append(int(val))

        x1, y1, x2, y2 = cleaned
    except Exception:
        return None

    # 3. 排序与有效性检查
    x1, x2 = sorted([x1, x2])
    y1, y2 = sorted([y1, y2])

    # 防止空框
    if x2 - x1 <= 1 or y2 - y1 <= 1:
        return None

    return [x1, y1, x2, y2]


def parse_issues_from_model_output(raw: str) -> Tuple[List[Dict[str, Any]], Optional[str]]:
    if raw is None:
        return [], "空响应"

    # 1. 提取 JSON 候选片段
    candidate = _extract_json_array_candidate(raw)
    if not candidate:
        return [], "未找到 JSON 数组"

    # 2. 先进行正则清洗 (增加对未加引号 Key 的预处理)
    text = _repair_common_json_issues(candidate)

    # 额外预处理：尝试给常见字段名强制加引号（防止正则漏网）
    # 针对 key: value 的情况
    known_keys = ["risk_level", "issue", "regulation", "correction", "bbox", "confidence"]
    for key in known_keys:
        # 如果出现 逗号/大括号 + 空格 + key + 冒号，说明 key 没加引号
        # (?<=[,{]\s) 匹配前面是逗号或大括号
        # (?=\s*:) 匹配后面是冒号
        text = re.sub(r'(?<=[,{]\s)' + key + r'(?=\s*:)', f'"{key}"', text)
        # 处理行首的情况
        text = re.sub(r'^\s*' + key + r'(?=\s*:)', f'"{key}"', text, flags=re.MULTILINE)

    data = None
    last_error = None

    # 3. 【核弹级修复】迭代式 JSON 解析
    # 增加重试次数到 10 次
    for attempt in range(10):
        try:
            data = json.loads(text)
            break  # 解析成功
        except json.JSONDecodeError as e:
            last_error = e
            msg = str(e)
            # print(f"DEBUG: JSON修复第{attempt+1}次: {msg} at pos {e.pos}") # 调试用

            # --- 策略 A: 缺少逗号 (Expecting ',' delimiter) ---
            if "Expecting ',' delimiter" in msg:
                try:
                    text = text[:e.pos] + "," + text[e.pos:]
                    continue
                except:
                    pass

            # --- 策略 B: 属性名问题/多余逗号 (Expecting property name...) ---
            elif "Expecting property name" in msg:
                try:
                    # 1. 检查是不是多余的逗号 ({ "a":1, })
                    prev_chunk = text[:e.pos].rstrip()
                    if prev_chunk.endswith(","):
                        comma_idx = text.rfind(",", 0, e.pos)
                        if comma_idx != -1:
                            text = text[:comma_idx] + text[e.pos:]
                            continue

                    # 2. 检查是不是单引号 Key ({'a': 1})
                    if e.pos < len(text) and text[e.pos] == "'":
                        text = text[:e.pos] + '"' + text[e.pos + 1:]
                        continue

                    # 3. 【新增】检查是不是未加引号的 Key ({ a: 1 })
                    # 如果报错位置是一个字母，尝试向后找到冒号，把这中间的单词包上引号
                    curr_char = text[e.pos]
                    if curr_char.isalpha():
                        # 寻找单词结束位置
                        match = re.match(r'\w+', text[e.pos:])
                        if match:
                            word = match.group(0)
                            # 替换为带引号的形式
                            text = text[:e.pos] + f'"{word}"' + text[e.pos + len(word):]
                            continue
                except:
                    pass

            # --- 策略 C: 字符串未闭合 (Unterminated string) ---
            elif "Unterminated string" in msg:
                try:
                    text += '"}]'
                    continue
                except:
                    pass

            # --- 策略 D: 期待值 (Expecting value) ---
            elif "Expecting value" in msg:
                try:
                    prev_chunk = text[:e.pos].rstrip()
                    if prev_chunk.endswith(","):
                        comma_idx = text.rfind(",", 0, e.pos)
                        if comma_idx != -1:
                            text = text[:comma_idx] + text[e.pos:]
                            continue
                except:
                    pass

            # 如果没有 continue，说明无法处理当前错误，只能尝试下一个策略或者退出
            # 这里不 break，而是让它进入下一次循环（也许上面的预处理有点用？）
            # 但为了防止死循环，如果文本没变，最好还是 break。这里简单处理：
            pass

    if data is None:
        # --- 最后的兜底：正则暴力提取 ---
        fallback_data = []
        try:
            # 匹配所有完整的 {...} 对象，尽可能抢救数据
            raw_objects = re.findall(r'\{[^{}]+\}', text)
            for obj_str in raw_objects:
                try:
                    # 清理一下可能的 Python 风格数据
                    obj_str = obj_str.replace("'", '"').replace("None", "null").replace("True", "true")
                    # 针对单个对象再试一次 known_keys 修复
                    for key in known_keys:
                        obj_str = re.sub(r'(?<=[,{]\s)' + key + r'(?=\s*:)', f'"{key}"', obj_str)

                    item = json.loads(obj_str)
                    fallback_data.append(item)
                except:
                    continue
        except:
            pass

        if fallback_data:
            data = fallback_data
            # print(f"⚠️ 抢救回 {len(data)} 条数据")
        else:
            return [], f"JSON 解析最终失败: {last_error}"

    try:
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
        return [], f"数据标准化失败: {e}"


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
# 确保这个辅助函数存在于 AnalysisWorker 类上方
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
    result_ready = pyqtSignal(str, dict)

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

    def _compress_image(self, path: str) -> str:
        """
        核心防崩逻辑：使用 QImageReader 限制读取大小，防止 OOM 闪退。
        """
        try:
            from PyQt6.QtGui import QImageReader

            # 1. 预检查图片信息，不直接加载数据
            reader = QImageReader(path)
            if not reader.canRead():
                print(f"❌ 无法读取图片: {path}")
                return ""

            # 2. 【核心修复】限制内存分配 (例如限制为 256MB)
            # 防止加载损坏的或分辨率异常巨大的图片
            reader.setAllocationLimit(256)

            # 3. 如果图片过大，先设置缩放读取（这一步非常关键，大幅降低内存）
            original_size = reader.size()
            max_dim = 1536
            if original_size.width() > max_dim or original_size.height() > max_dim:
                # 计算缩放比例
                reader.setScaledSize(original_size.scaled(max_dim, max_dim, Qt.AspectRatioMode.KeepAspectRatio))

            # 4. 执行读取
            img = reader.read()
            if img.isNull():
                print(f"❌ 图片数据为空: {reader.errorString()}")
                return ""

            # 5. 压缩转 Base64 (JPEG 质量 80)
            ba = QByteArray()
            buf = QBuffer(ba)
            buf.open(QIODevice.OpenModeFlag.WriteOnly)
            img.save(buf, "JPEG", 80)
            b64_str = ba.toBase64().data().decode("utf-8")

            # 显式清理
            del img
            del reader

            return b64_str

        except Exception as e:
            print(f"❌ 压缩过程异常: {e}\n{traceback.format_exc()}")
            return ""

    def run(self):
        started = time.time()
        p_name = "未知"
        model = "未知"

        try:
            p_name, api_key, base_url, model = self._get_provider_conf()

            if not api_key or not base_url or not model:
                self.result_ready.emit(self.task['id'], {
                    "ok": False, "error": "配置缺失(Key/URL/Model)", "elapsed_sec": 0
                })
                return

            # 1. 执行压缩
            img_b64 = self._compress_image(self.task["path"])

            # 2. 如果压缩失败，直接报错，不再继续（防止原图撑爆内存）
            if not img_b64:
                self.result_ready.emit(self.task['id'], {
                    "ok": False, "error": "图片加载或压缩失败(可能是路径含中文或缺少组件)", "elapsed_sec": 0
                })
                return

            # 3. 发送请求
            with httpx.Client(
                    http2=False,
                    verify=False,
                    trust_env=False,
                    timeout=float(self.config.get("request_timeout_sec", 60))
            ) as http_client:

                client = OpenAI(api_key=api_key, base_url=base_url, http_client=http_client)

                system_prompt = (self.prompt_text.strip() + "\n\n" + build_strict_json_guard())
                max_retries = int(self.config.get("max_retries", 2))
                last_error = None

                for attempt in range(max_retries + 1):
                    try:
                        resp = client.chat.completions.create(
                            model=model,
                            messages=[
                                {"role": "system", "content": system_prompt},
                                {
                                    "role": "user",
                                    "content": [
                                        {"type": "image_url",
                                         "image_url": {"url": f"data:image/jpeg;base64,{img_b64}"}},
                                        {"type": "text", "text": "请严格按要求输出 JSON 数组。"}
                                    ]
                                }
                            ],
                            temperature=float(self.config.get("temperature", 0.1))
                        )

                        raw = resp.choices[0].message.content or ""
                        issues, err = parse_issues_from_model_output(raw)
                        elapsed = round(time.time() - started, 2)

                        if err:
                            self.result_ready.emit(self.task["id"], {
                                "ok": False, "error": f"解析失败: {err}", "issues": [],
                                "elapsed_sec": elapsed, "provider": p_name, "model": model
                            })
                            return

                        self.result_ready.emit(self.task["id"], {
                            "ok": True, "error": None, "issues": issues,
                            "elapsed_sec": elapsed, "provider": p_name, "model": model
                        })
                        return

                    except Exception as e:
                        last_error = e
                        print(f"请求重试 ({attempt + 1}): {e}")
                        if attempt < max_retries:
                            time.sleep(2)
                        else:
                            break

            elapsed = round(time.time() - started, 2)
            self.result_ready.emit(self.task["id"], {
                "ok": False, "error": str(last_error), "issues": [], "elapsed_sec": elapsed
            })

        except BaseException as e:
            elapsed = round(time.time() - started, 2)
            print("系统级异常:", traceback.format_exc())
            self.result_ready.emit(self.task["id"], {
                "ok": False, "error": f"异常: {e}", "issues": [], "elapsed_sec": elapsed
            })


# ================= 新增：自定义可编辑文字项 (彻底解决交互冲突) =================
class EditableTextItem(QGraphicsTextItem):
    """
    自定义文字项：解决 View 拖拽模式下的事件冲突
    """

    def __init__(self, text, parent=None, callback=None):
        super().__init__(text, parent)
        self.callback = callback

        # 核心 Flag
        self.setFlags(
            QGraphicsItem.GraphicsItemFlag.ItemIsMovable |
            QGraphicsItem.GraphicsItemFlag.ItemIsSelectable |
            QGraphicsItem.GraphicsItemFlag.ItemIsFocusable
        )
        self.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)

        # 样式
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setDefaultTextColor(QColor("#FF0000"))

    def mouseDoubleClickEvent(self, event):
        """双击进入编辑模式"""
        if event.button() == Qt.MouseButton.LeftButton:
            # 1. 切换为编辑模式
            self.setTextInteractionFlags(Qt.TextInteractionFlag.TextEditorInteraction)

            # 2. 【关键】编辑期间禁止移动，否则选文字时框会跑
            self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, False)

            # 3. 强制获取焦点
            self.setFocus()
            self.setCursor(Qt.CursorShape.IBeamCursor)

            # 4. 交给父类处理光标定位
            super().mouseDoubleClickEvent(event)

            # 5. 通知 View 暂时禁用画布拖拽
            if self.scene() and self.scene().views():
                self.scene().views()[0].setDragMode(QGraphicsView.DragMode.NoDrag)
        else:
            super().mouseDoubleClickEvent(event)

    def focusOutEvent(self, event):
        """失去焦点时保存"""
        self.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        # 清除选中效果
        cursor = self.textCursor()
        cursor.clearSelection()
        self.setTextCursor(cursor)

        if self.callback:
            self.callback(self)

        # 恢复 View 的手型
        if self.scene() and self.scene().views():
            view = self.scene().views()[0]
            if hasattr(view, "_tool") and view._tool == "none":
                view.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)

        super().focusOutEvent(event)


    # ================= 修复方案：新增 EditableTextItem 类 =================


class EditableTextItem(QGraphicsTextItem):
    """
    自定义文字项：
    1. 自身管理“移动”与“编辑”状态的切换。
    2. 解决与 View 拖拽手势的冲突。
    """

    def __init__(self, text, parent=None, callback=None):
        super().__init__(text, parent)
        self.callback = callback  # 编辑完成后的回调（用于保存历史/撤销）

        # 初始状态：允许移动、选中、聚焦，但不可编辑文字
        self.setFlags(
            QGraphicsItem.GraphicsItemFlag.ItemIsMovable |
            QGraphicsItem.GraphicsItemFlag.ItemIsSelectable |
            QGraphicsItem.GraphicsItemFlag.ItemIsFocusable
        )
        self.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)

        # 样式设置
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setDefaultTextColor(QColor("#FF0000"))

    def mouseDoubleClickEvent(self, event):
        """双击进入编辑模式"""
        if event.button() == Qt.MouseButton.LeftButton:
            # 1. 开启文字编辑
            self.setTextInteractionFlags(Qt.TextInteractionFlag.TextEditorInteraction)

            # 2. 【关键】进入编辑时必须禁止移动，否则鼠标选字会变成拖动框体
            self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, False)

            # 3. 强制获取焦点并弹出光标
            self.setFocus()
            self.setCursor(Qt.CursorShape.IBeamCursor)

            # 4. 通知 View 暂时彻底禁用画布拖拽（双重保险）
            if self.scene() and self.scene().views():
                self.scene().views()[0].setDragMode(QGraphicsView.DragMode.NoDrag)

            # 5. 调用父类处理光标定位
            super().mouseDoubleClickEvent(event)
        else:
            super().mouseDoubleClickEvent(event)

    def focusOutEvent(self, event):
        """失去焦点（点击别处）时，保存并退出编辑"""
        # 1. 关闭编辑，恢复只读
        self.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)

        # 2. 恢复可移动状态
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        # 3. 清除文字的选中背景色（美观）
        cursor = self.textCursor()
        cursor.clearSelection()
        self.setTextCursor(cursor)

        # 4. 触发回调通知 View 保存数据
        if self.callback:
            self.callback(self)

        # 5. 尝试恢复 View 的手型拖拽（如果当前不是在绘图工具模式下）
        if self.scene() and self.scene().views():
            view = self.scene().views()[0]
            if hasattr(view, "_tool") and view._tool == "none":
                view.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)

        super().focusOutEvent(event)


# ================= 修复版：图片标注画布 =================
class AnnotatableImageView(QGraphicsView):
    annotation_changed = pyqtSignal()
    tool_reset = pyqtSignal()

    TOOL_NONE = "none"
    TOOL_RECT = "rect"
    TOOL_ELLIPSE = "ellipse"
    TOOL_ARROW = "arrow"
    TOOL_TEXT = "text"
    TOOL_ISSUE_TAG = "issue_tag"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setScene(QGraphicsScene(self))

        # === 底图图元 ===
        self._pix_item = QGraphicsPixmapItem()
        self._pix_item.setZValue(-1000)  # 保证在最底层
        self._pix_item.setAcceptedMouseButtons(Qt.MouseButton.NoButton)
        self.scene().addItem(self._pix_item)

        # === 状态变量 ===
        self._tool = self.TOOL_NONE
        self._dragging = False
        self._start_img_pt = None
        self._img_path = None
        self._ai_issues = []
        self._draw_color = "#FF0000"
        self._draw_width = 6
        self._base_img_size = (1, 1)

        # === 视图设置 ===
        self.setRenderHints(QPainter.RenderHint.Antialiasing | QPainter.RenderHint.SmoothPixmapTransform)
        self.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setMouseTracking(True)
        self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)

    # ... (鼠标事件保持不变，与你原代码一致，略去以节省篇幅，请保留原有的 mousePressEvent 等) ...

    # 只需要替换 mousePressEvent, mouseDoubleClickEvent, mouseMoveEvent, mouseReleaseEvent
    # 如果你没有修改过这部分，可以保留原文件中的事件处理代码。
    # 重点在于下面的功能函数修复。

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            item = self.itemAt(event.position().toPoint())

            # 优先处理文字编辑
            if isinstance(item, QGraphicsTextItem):
                if self.dragMode() != QGraphicsView.DragMode.NoDrag:
                    self.setDragMode(QGraphicsView.DragMode.NoDrag)
                super().mousePressEvent(event)
                return

            # 允许移动已有的框（浏览模式下）
            if isinstance(item, QGraphicsItem) and item is not self._pix_item and self._tool == self.TOOL_NONE:
                self.setDragMode(QGraphicsView.DragMode.NoDrag)
                super().mousePressEvent(event)
                return

            # 引用问题工具
            if self._tool == self.TOOL_ISSUE_TAG:
                pos = self._to_img_point(event.position().toPoint())
                self._handle_tag_creation(pos)
                return

            # 绘图工具
            if self._tool != self.TOOL_NONE:
                self._dragging = True
                self._start_img_pt = self._to_img_point(event.position().toPoint())
                self._temp_end_img_pt = self._start_img_pt
                return

            # 浏览模式恢复拖拽
            if self._tool == self.TOOL_NONE:
                if self.dragMode() != QGraphicsView.DragMode.ScrollHandDrag:
                    self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)

        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._dragging and self._tool != self.TOOL_NONE:
            self._temp_end_img_pt = self._to_img_point(event.position().toPoint())
            self.viewport().update()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._dragging and self._tool != self.TOOL_NONE:
            self._finish_drawing(event)
        super().mouseReleaseEvent(event)
        if self._tool == self.TOOL_NONE and not self.scene().focusItem():
            self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)

    def paintEvent(self, event):
        super().paintEvent(event)
        if self._dragging and self._start_img_pt and self._temp_end_img_pt:
            painter = QPainter(self.viewport())
            painter.setPen(QPen(QColor(self._draw_color), 2, Qt.PenStyle.DashLine))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            p1 = self.mapFromScene(self._start_img_pt)
            p2 = self.mapFromScene(self._temp_end_img_pt)
            x = min(p1.x(), p2.x())
            y = min(p1.y(), p2.y())
            w = abs(p1.x() - p2.x())
            h = abs(p1.y() - p2.y())

            if self._tool in (self.TOOL_RECT, self.TOOL_ELLIPSE, self.TOOL_TEXT):
                if self._tool == self.TOOL_ELLIPSE:
                    painter.drawEllipse(x, y, w, h)
                else:
                    painter.drawRect(x, y, w, h)
            elif self._tool == self.TOOL_ARROW:
                painter.drawLine(p1, p2)

    # ... (保留 _finish_drawing, _create_text_annotation, _handle_tag_creation, _open_issue_dialog 逻辑) ...
    # 为节省空间，请确保这几个辅助函数存在，代码逻辑与原文件一致即可。

    def _finish_drawing(self, event):
        self._dragging = False
        start = self._start_img_pt
        end = self._to_img_point(event.position().toPoint())
        self._start_img_pt = None
        self._temp_end_img_pt = None
        self.viewport().update()

        if not start or not end: return
        if abs(start.x() - end.x()) < 5 and abs(start.y() - end.y()) < 5:
            if self._tool == self.TOOL_TEXT: self._create_text_annotation(start)
            return

        data = None
        if self._tool in (self.TOOL_RECT, self.TOOL_ELLIPSE):
            x1, x2 = sorted([start.x(), end.x()])
            y1, y2 = sorted([start.y(), end.y()])
            data = {"type": self._tool, "bbox": [int(x1), int(y1), int(x2), int(y2)],
                    "color": self._draw_color, "width": self._draw_width}
        elif self._tool == self.TOOL_ARROW:
            data = {"type": "arrow", "p1": [int(start.x()), int(start.y())], "p2": [int(end.x()), int(end.y())],
                    "color": self._draw_color, "width": self._draw_width}
        elif self._tool == self.TOOL_TEXT:
            self._create_text_annotation(end)
            return

        if data:
            self._create_graphics_item_from_data(data)
            self.annotation_changed.emit()

    def _create_text_annotation(self, pos):
        text, ok = QInputDialog.getText(self, "输入标注文字", "文字内容:")
        if ok and text:
            data = {"type": "text", "pos": [int(pos.x()), int(pos.y())], "text": text,
                    "color": self._draw_color, "font_size": 36}
            self._create_graphics_item_from_data(data)
            self.annotation_changed.emit()

    def _handle_tag_creation(self, pos):
        if not self._ai_issues:
            QMessageBox.information(self, "提示", "当前没有 AI 识别出的问题可引用。\n请先进行[开始分析]。")
            self.tool_reset.emit()
            return
        safe_pos = QPointF(pos.x(), pos.y())
        QTimer.singleShot(0, lambda: self._open_issue_dialog(safe_pos))

    def _open_issue_dialog(self, pos):
        dlg = IssueSelectionDialog(self, self._ai_issues)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            data = {"type": "text", "pos": [int(pos.x()), int(pos.y())], "text": dlg.selected_text,
                    "color": dlg.selected_color, "font_size": 28}
            self._create_graphics_item_from_data(data)
            self.annotation_changed.emit()

    # ================= 核心修复区域：创建与导出 =================

    def _create_graphics_item_from_data(self, data: Dict[str, Any]):
        t = data.get("type")
        color = QColor(data.get("color", "#FF0000"))
        w = int(data.get("width", 6))
        pen = QPen(color, w)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)

        item = None
        if t == "text":
            x, y = data.get("pos", [0, 0])
            txt = data.get("text", "")
            fs = int(data.get("font_size", 28))
            item = EditableTextItem(txt, callback=self._save_text_item_data)
            font = QFont()
            font.setPointSize(fs)
            font.setBold(True)
            item.setFont(font)
            item.setDefaultTextColor(color)
            item.setPos(x, y)
        elif t == "rect":
            x1, y1, x2, y2 = data.get("bbox", [0, 0, 0, 0])
            # 创建时使用相对坐标，但设置 Pos 为 0,0 (默认)
            item = QGraphicsRectItem(QRectF(x1, y1, x2 - x1, y2 - y1))
        elif t == "ellipse":
            x1, y1, x2, y2 = data.get("bbox", [0, 0, 0, 0])
            item = QGraphicsEllipseItem(QRectF(x1, y1, x2 - x1, y2 - y1))
        elif t == "arrow":
            x1, y1 = data.get("p1", [0, 0])
            x2, y2 = data.get("p2", [0, 0])
            path = QPainterPath()
            path.moveTo(x1, y1)
            path.lineTo(x2, y2)
            item = QGraphicsPathItem(path)

        if item:
            if t != "text":
                item.setPen(pen)
                item.setFlags(
                    QGraphicsItem.GraphicsItemFlag.ItemIsMovable | QGraphicsItem.GraphicsItemFlag.ItemIsSelectable)

            # 使用 data.copy() 防止引用污染
            item.setData(Qt.ItemDataRole.UserRole, data.copy())
            self.scene().addItem(item)

    def set_user_annotations(self, anns):
        # 【核心修复1】加载数据时暂时屏蔽信号，防止清空操作触发“保存为空”
        self.blockSignals(True)
        try:
            self.clear_annotations()
            if not anns: return
            for a in anns:
                self._create_graphics_item_from_data(a)
        finally:
            self.blockSignals(False)

    def get_user_annotations(self):
        self.scene().clearFocus()
        anns = []
        items = list(self.scene().items(Qt.SortOrder.AscendingOrder))
        for item in items:
            if item is self._pix_item: continue

            # 获取并复制原始数据
            raw_data = item.data(Qt.ItemDataRole.UserRole)
            if not raw_data: continue
            data = raw_data.copy()

            # 【核心修复2】使用 sceneBoundingRect 获取绝对坐标，支持拖拽后的位置保存
            if isinstance(item, QGraphicsTextItem):
                data["text"] = item.toPlainText()
                data["pos"] = [int(item.pos().x()), int(item.pos().y())]
            elif isinstance(item, QGraphicsRectItem) or isinstance(item, QGraphicsEllipseItem):
                # 获取在场景中的绝对包围盒
                r = item.sceneBoundingRect()
                data["bbox"] = [int(r.left()), int(r.top()), int(r.right()), int(r.bottom())]
            elif isinstance(item, QGraphicsPathItem) and data.get("type") == "arrow":
                # 箭头通常由点定义，如果支持移动，需要应用偏移量 (这里简化处理，箭头通常不移动或重绘)
                # 如果箭头也被移动了，需要更复杂的逻辑，但在你的代码里箭头是 Path，比较难直接反算点
                # 这里假设箭头移动需求较少，或者通过 pos 偏移修正
                offset = item.pos()
                p1 = data.get("p1", [0, 0])
                p2 = data.get("p2", [0, 0])
                data["p1"] = [int(p1[0] + offset.x()), int(p1[1] + offset.y())]
                data["p2"] = [int(p2[0] + offset.x()), int(p2[1] + offset.y())]

            anns.append(data)
        return anns

    # ... (保持原有的辅助函数) ...
    def set_tool(self, tool: str):
        self._tool = tool
        self._dragging = False
        if tool == self.TOOL_NONE:
            self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
            self.setCursor(Qt.CursorShape.OpenHandCursor)
        else:
            self.setDragMode(QGraphicsView.DragMode.NoDrag)
            self.setCursor(Qt.CursorShape.CrossCursor)

    def set_image(self, path: str):
        self._img_path = path
        reader = QImage(path)
        if reader.isNull(): return
        pix = QPixmap.fromImage(reader)
        self._base_pix = pix
        self._base_img_size = (max(1, pix.width()), max(1, pix.height()))
        self._pix_item.setPixmap(pix)
        self.scene().setSceneRect(QRectF(0, 0, pix.width(), pix.height()))
        self.resetTransform()
        self.fitInView(self.sceneRect(), Qt.AspectRatioMode.KeepAspectRatio)

    def _to_img_point(self, view_pos) -> QPointF:
        sp = self.mapToScene(view_pos)
        x = min(max(sp.x(), 0.0), float(self._base_img_size[0]))
        y = min(max(sp.y(), 0.0), float(self._base_img_size[1]))
        return QPointF(x, y)

    def wheelEvent(self, event):
        if event.angleDelta().y() > 0:
            self.scale(1.25, 1.25)
        else:
            self.scale(0.8, 0.8)

    def clear_annotations(self):
        # 仅删除非底图的元素
        for item in list(self.scene().items()):
            if item is not self._pix_item:
                self.scene().removeItem(item)
        # 注意：这里会发射信号，所以在 set_user_annotations 里必须屏蔽
        self.annotation_changed.emit()

    def delete_selected_items(self):
        for item in self.scene().selectedItems():
            if item is not self._pix_item: self.scene().removeItem(item)
        self.annotation_changed.emit()

    def undo(self):
        items = [i for i in self.scene().items() if i is not self._pix_item]
        if items:
            self.scene().removeItem(items[0])
            self.annotation_changed.emit()

    def set_ai_issues(self, issues):
        self._ai_issues = issues or []

    def _save_text_item_data(self, item):
        self.annotation_changed.emit()

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

        # 简单的颜色匹配逻辑
        bg, bd = colors["蓝"], borders["蓝"]
        if any(x in level for x in ["重大", "严重", "High", "警示", "红线"]):
            bg, bd = colors["红"], borders["红"]
        elif any(x in level for x in ["较大", "一般", "质量", "需整理", "Medium"]):
            bg, bd = colors["橙"], borders["橙"]

        self.setStyleSheet(
            f"RiskCard {{ background-color: {bg}; border-left: 5px solid {bd}; border-radius: 4px; margin-bottom: 6px; padding: 6px; }}"
        )

        layout = QVBoxLayout(self)
        header = QHBoxLayout()

        header.addWidget(QLabel(f"<b>[{level}]</b>"))

        # 【核心修复】限制文本显示长度
        raw_issue = str(item.get("issue", ""))
        display_issue = raw_issue[:200] + "..." if len(raw_issue) > 200 else raw_issue

        lbl_issue = QLabel(display_issue)
        lbl_issue.setWordWrap(True)
        # 增加 Tooltip，鼠标悬停才显示完整内容，防止布局计算卡死
        lbl_issue.setToolTip(raw_issue[:1000])

        header.addWidget(lbl_issue, 1)

        btn_edit = QPushButton("编辑")
        btn_edit.setFixedWidth(70)
        btn_edit.clicked.connect(self.on_edit_clicked)
        header.addWidget(btn_edit)

        btn_del = QPushButton("删除")
        btn_del.setFixedWidth(70)
        btn_del.clicked.connect(self.on_delete_clicked)
        header.addWidget(btn_del)

        layout.addLayout(header)

        bbox = item.get("bbox")
        bbox_text = f"{bbox}" if bbox else "无/未定位"

        # 同样对其他字段做长度保护
        reg_txt = str(item.get('regulation', ''))
        layout.addWidget(QLabel(f"依据: {reg_txt[:100]}"))
        layout.addWidget(QLabel(f"定位 bbox: {bbox_text}"))

        corr_txt = str(item.get('correction', ''))
        lbl_fix = QLabel(f"建议: {corr_txt[:200]}")
        lbl_fix.setStyleSheet("color: #2E7D32; font-weight: bold;")
        lbl_fix.setWordWrap(True)
        layout.addWidget(lbl_fix)

    def on_edit_clicked(self):
        self.edit_requested.emit(self.item)

    def on_delete_clicked(self):
        self.delete_requested.emit(self.item)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # 🔧 添加全局异常处理
        sys.excepthook = self._global_exception_handler
        # 1. 加载配置和业务数据
        self.config = ConfigManager.load()
        self.refresh_business_data()

        # 2. 初始化任务变量
        self.tasks: List[Dict[str, Any]] = []
        self.current_task_id: Optional[str] = None
        self.running_workers: Dict[str, AnalysisWorker] = {}
        self.pending_queue: List[str] = []
        self.total_task = 0
        self.done_task = 0

        # 3. 初始化 UI 界面 (确保之前已经修复了 init_ui 的顺序)
        self.init_ui()

        # 4. 初始化计时器
        self._resize_timer = QTimer(self)
        self._resize_timer.setInterval(200)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.timeout.connect(self._refresh_current_image)

    # --- 以下是独立的方法，不要写在 __init__ 里面 ---
    def _global_exception_handler(self, exc_type, exc_value, exc_traceback):
        """捕获所有未处理的异常"""
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return

        error_msg = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        print(f"❌ 全局异常捕获:\n{error_msg}")

        QMessageBox.critical(
            None,
            "程序错误",
            f"发生未处理的异常：\n{exc_type.__name__}: {exc_value}\n\n详情请查看控制台输出"
        )

    def auto_annotate_current_task(self):
        """根据 AI 识别的 bbox,在图片中心自动生成文字标识（增强版）"""
        task = self._current_task()
        if not task or task.get("status") != "done":
            QMessageBox.warning(self, "提示", "请先完成 AI 分析后再使用自动标识。")
            return

        issues = task.get("edited_issues") if task.get("edited_issues") is not None else task.get("issues", [])

        if not issues:
            QMessageBox.information(self, "提示", "未检测到任何可标注的问题。")
            return

        count = 0
        for idx, item in enumerate(issues, 1):
            bbox = item.get("bbox")
            if bbox and isinstance(bbox, list) and len(bbox) == 4:
                # 1. 计算框的中心点
                cx = (bbox[0] + bbox[2]) / 2
                cy = (bbox[1] + bbox[3]) / 2

                # 2. 获取描述文本（优化：添加序号）
                text = item.get("issue", "未知问题")
                if len(text) > 15:
                    text = text[:15] + "..."

                # ✅ 添加序号便于识别
                text = f"{idx}. {text}"

                # 3. 确定颜色
                level = item.get("risk_level", "")
                if any(x in level for x in ["严重", "红线"]):
                    color = "#FF0000"  # 红色
                elif any(x in level for x in ["文明"]):
                    color = "#2196F3"  # 蓝色
                else:
                    color = "#FF8800"  # 橙色

                # 4. 构造标注并创建
                new_anno = {
                    "type": "text",
                    "pos": [int(cx), int(cy)],
                    "text": text,
                    "color": color,
                    "width": 4,
                    "font_size": 32  # ✅ 增大字号便于编辑
                }
                self.image_view._create_graphics_item_from_data(new_anno)
                count += 1

        if count > 0:
            # 同步更新任务中的标注数据
            task["annotations"] = self.image_view.get_user_annotations()
            self.status_bar.showMessage(f"✅ 成功自动标识 {count} 处问题（双击文字可编辑）", 5000)
        else:
            QMessageBox.information(self, "提示", "AI 识别结果中未包含具体坐标(bbox)，无法自动标识。")


    def refresh_business_data(self):
        self.business_data = self.config.get("business_data", DEFAULT_BUSINESS_DATA)

    def init_ui(self):
        self.setWindowTitle("普洱版纳区域检查报告助手V2.0")
        self.resize(1320, 980)

        # ================= 1. 顶部工具栏 (Toolbar) =================
        toolbar = QToolBar("Main")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        toolbar.addWidget(QLabel("  场景模式: "))
        self.cbo_prompt = QComboBox()
        prompts = self.config.get("prompts", DEFAULT_PROMPTS)
        self.cbo_prompt.addItems(prompts.keys())
        self.cbo_prompt.setCurrentText(self.config.get("last_prompt", list(prompts.keys())[0]))
        self.cbo_prompt.setMinimumWidth(280)
        toolbar.addWidget(self.cbo_prompt)

        toolbar.addSeparator()

        self.act_add = QAction("➕ 添加图片", self)
        self.act_run = QAction("▶ 开始分析", self)
        self.act_pause = QAction("⏸ 暂停", self)
        self.act_clear = QAction("🗑️ 清空队列", self)
        toolbar.addAction(self.act_add)
        toolbar.addAction(self.act_run)
        toolbar.addAction(self.act_pause)
        toolbar.addAction(self.act_clear)

        # 导出报告下拉菜单
        self.btn_export_tool = QToolButton()
        self.btn_export_tool.setText("📄 导出报告 ▼")
        self.btn_export_tool.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        export_menu = QMenu(self)
        self.act_report_check = QAction("通用检查报告 (使用 检查模板.docx)", self)
        self.act_report_notice = QAction("整改通知单 (使用 通知单模板.docx)", self)
        self.act_report_simple = QAction("简报模式 (使用 简报模板.docx)", self)
        export_menu.addActions([self.act_report_check, self.act_report_notice, self.act_report_simple])
        self.btn_export_tool.setMenu(export_menu)
        toolbar.addWidget(self.btn_export_tool)

        empty = QWidget()
        empty.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        toolbar.addWidget(empty)
        self.act_help = QAction("❓ 帮助", self)
        toolbar.addAction(self.act_help)
        self.act_setting = QAction("⚙ 设置", self)
        toolbar.addAction(self.act_setting)

        # ================= 2. 基础信息面板 (Info Group) =================
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        info_group = QGroupBox("📄 报告基础信息")
        info_group.setFixedHeight(210)
        info_layout = QGridLayout(info_group)

        self.input_company = QComboBox()
        self.update_company_combo()
        self.input_project = QComboBox()
        self.input_inspected_unit = QLineEdit()
        self.input_check_content = QComboBox()
        self.update_check_content_combo()
        self.input_check_content.setEditable(True)
        self.input_area = QLineEdit()
        self.input_person = QLineEdit(self.config.get("last_check_person", ""))
        self.input_date = QLineEdit(datetime.now().strftime("%Y-%m-%d"))
        self.input_deadline = QLineEdit()
        self.input_group = QLineEdit()

        # 期限快捷键
        quick_deadline_widget = QWidget()
        quick_box = QHBoxLayout(quick_deadline_widget)
        self.btn_day3 = QPushButton("+3天")
        self.btn_day7 = QPushButton("+7天")
        self.btn_day15 = QPushButton("+15天")
        for b in (self.btn_day3, self.btn_day7, self.btn_day15): b.setFixedWidth(60)
        quick_box.addWidget(self.btn_day3);
        quick_box.addWidget(self.btn_day7);
        quick_box.addWidget(self.btn_day15);
        quick_box.addStretch()

        info_layout.addWidget(QLabel("项目公司:"), 0, 0);
        info_layout.addWidget(self.input_company, 0, 1)
        info_layout.addWidget(QLabel("项目名称:"), 0, 2);
        info_layout.addWidget(self.input_project, 0, 3)
        info_layout.addWidget(QLabel("被检单位:"), 1, 0);
        info_layout.addWidget(self.input_inspected_unit, 1, 1)
        info_layout.addWidget(QLabel("检查内容:"), 1, 2);
        info_layout.addWidget(self.input_check_content, 1, 3)
        info_layout.addWidget(QLabel("检查部位:"), 2, 0);
        info_layout.addWidget(self.input_area, 2, 1)
        info_layout.addWidget(QLabel("检查人员:"), 2, 2);
        info_layout.addWidget(self.input_person, 2, 3)
        info_layout.addWidget(QLabel("检查日期:"), 3, 0);
        info_layout.addWidget(self.input_date, 3, 1)
        info_layout.addWidget(QLabel("整改期限:"), 3, 2);
        info_layout.addWidget(self.input_deadline, 3, 3)
        info_layout.addWidget(QLabel("点位分组:"), 4, 0);
        info_layout.addWidget(self.input_group, 4, 1)
        info_layout.addWidget(QLabel("快捷期限:"), 4, 2);
        info_layout.addWidget(quick_deadline_widget, 4, 3)
        main_layout.addWidget(info_group)

        # ================= 3. 主分割面板 (Splitter) =================
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # --- 左侧列表 ---
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        self.lbl_count = QLabel(f"待审队列 (0/{MAX_IMAGES})")
        self.list_widget = QListWidget()
        self.btn_apply_group = QPushButton("批量设点位")
        self.btn_retry_error = QPushButton("重试失败")
        batch_box = QHBoxLayout()
        batch_box.addWidget(self.btn_apply_group);
        batch_box.addWidget(self.btn_retry_error)
        left_layout.addWidget(self.lbl_count);
        left_layout.addWidget(self.list_widget);
        left_layout.addLayout(batch_box)

        # --- 右侧预览与标注 ---
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)

        # 标注工具栏
        self.image_view = AnnotatableImageView()  # 先初始化它，因为按钮需要连它的方法

        self.btn_tool_none = QPushButton("缩放")
        self.btn_tool_rect = QPushButton("框")
        self.btn_tool_ellipse = QPushButton("圈")
        self.btn_tool_arrow = QPushButton("箭头")
        self.btn_tool_text = QPushButton("文字")
        self.btn_tool_tag = QPushButton("🏷️引用问题")

        self.btn_undo = QPushButton("撤销")
        self.btn_delete_selected = QPushButton("删除选中")
        self.btn_clear_anno = QPushButton("清空")
        self.btn_auto_annotate = QPushButton("🤖 自动标识")
        self.btn_save_marked = QPushButton("保存截图")

        # 按钮样式
        for b in [self.btn_tool_none, self.btn_tool_rect, self.btn_tool_ellipse, self.btn_tool_arrow,
                  self.btn_tool_text]:
            b.setFixedWidth(60)
        self.btn_tool_tag.setStyleSheet("color: blue; font-weight: bold;");
        self.btn_tool_tag.setFixedWidth(80)
        self.btn_auto_annotate.setStyleSheet("background-color: #E8F5E9; color: #2E7D32; font-weight: bold;")

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("绘图:"));
        row1.addWidget(self.btn_tool_none);
        row1.addWidget(self.btn_tool_rect);
        row1.addWidget(self.btn_tool_ellipse)
        row1.addWidget(self.btn_tool_arrow);
        row1.addWidget(self.btn_tool_text);
        row1.addWidget(self.btn_tool_tag);

        row1.addStretch()

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("操作:"));
        row2.addWidget(self.btn_undo);
        row2.addWidget(self.btn_delete_selected);
        row2.addWidget(self.btn_clear_anno)
        row2.addWidget(self.btn_auto_annotate);
        row2.addWidget(self.btn_save_marked);
        row2.addStretch()

        right_layout.addLayout(row1);
        right_layout.addLayout(row2)
        right_layout.addWidget(self.image_view, 2)


        self.result_container = QWidget()
        self.result_layout = QVBoxLayout(self.result_container)
        self.result_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        scroll = QScrollArea();
        scroll.setWidgetResizable(True);
        scroll.setWidget(self.result_container)
        right_layout.addWidget(scroll, 3)

        splitter.addWidget(left_widget);
        splitter.addWidget(right_widget)
        splitter.setSizes([380, 940])
        main_layout.addWidget(splitter)

        # 状态栏
        self.status_bar = self.statusBar()
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False);
        self.progress_bar.setFixedWidth(200)
        self.status_bar.addPermanentWidget(self.progress_bar)

        # ================= 4. 最后：连接所有信号 (Signals) =================
        # 工具栏动作
        self.act_add.triggered.connect(self.add_files)
        self.act_run.triggered.connect(self.start_analysis)
        self.act_pause.triggered.connect(self.pause_analysis)
        self.act_clear.triggered.connect(self.clear_queue)
        self.act_setting.triggered.connect(self.open_settings)
        self.act_help.triggered.connect(self.show_help)
        self.act_report_check.triggered.connect(lambda: self.export_word("检查模板.docx"))
        self.act_report_notice.triggered.connect(lambda: self.export_word("通知单模板.docx"))
        self.act_report_simple.triggered.connect(lambda: self.export_word("简报模板.docx"))
        self.cbo_prompt.currentTextChanged.connect(self.save_prompt_selection)

        # 业务数据联动
        self.input_company.currentTextChanged.connect(self.on_company_changed)
        if self.input_company.count() > 0: self.on_company_changed(self.input_company.currentText())
        self.btn_day3.clicked.connect(lambda: self._set_deadline_days(3))
        self.btn_day7.clicked.connect(lambda: self._set_deadline_days(7))
        self.btn_day15.clicked.connect(lambda: self._set_deadline_days(15))

        # 列表与批量
        self.list_widget.itemClicked.connect(self.on_item_clicked)
        self.btn_apply_group.clicked.connect(self.apply_group_to_all_tasks)
        self.btn_retry_error.clicked.connect(self.retry_errors)

        # 标注工具连接
        self.btn_tool_none.clicked.connect(lambda: self._set_tool(AnnotatableImageView.TOOL_NONE))
        self.btn_tool_rect.clicked.connect(lambda: self._set_tool(AnnotatableImageView.TOOL_RECT))
        self.btn_tool_ellipse.clicked.connect(lambda: self._set_tool(AnnotatableImageView.TOOL_ELLIPSE))
        self.btn_tool_arrow.clicked.connect(lambda: self._set_tool(AnnotatableImageView.TOOL_ARROW))
        self.btn_tool_text.clicked.connect(lambda: self._set_tool(AnnotatableImageView.TOOL_TEXT))
        self.btn_tool_tag.clicked.connect(lambda: self._set_tool(AnnotatableImageView.TOOL_ISSUE_TAG))


        self.btn_undo.clicked.connect(self._undo_annotation)
        self.btn_delete_selected.clicked.connect(self.image_view.delete_selected_items)
        self.btn_clear_anno.clicked.connect(self._clear_annotation)
        self.btn_auto_annotate.clicked.connect(self.auto_annotate_current_task)
        self.btn_save_marked.clicked.connect(self._save_marked_for_current_task)

        # 图像视图回调
        self.image_view.annotation_changed.connect(self._on_annotation_changed)
        self.image_view.tool_reset.connect(lambda: self._set_tool(AnnotatableImageView.TOOL_NONE))

    def _set_tool(self, tool: str):
        self.image_view.set_tool(tool)

    def show_help(self):
        help_content = """
        <h3>普洱版纳区域检查报告助手 使用说明</h3>
        <p>本工具旨在辅助用户快速生成包含 AI 辅助分析和人工标注的检查报告。</p>

        <h4><strong>一、 基本操作流程</strong></h4>
        <ol>
            <li><strong>添加图片</strong>：点击工具栏的“➕ 添加图片”按钮，选择要分析的现场照片。图片会加入左侧的任务队列。</li>
            <li><strong>填写报告信息</strong>：在“报告基础信息”区域填写检查所需的各项内容，如项目公司、项目名称、检查人员、检查日期等。</li>
            <li><strong>选择场景模式</strong>：在工具栏的“场景模式”下拉框中选择合适的 AI 分析模型提示词（如“施工全能扫描”）。</li>
            <li><strong>开始分析</strong>：点击工具栏的“▶ 开始分析”按钮，AI 将对队列中的图片进行分析。分析进度会在底部状态栏显示。</li>
            <li><strong>查看与编辑结果</strong>：
                <ul>
                    <li>点击左侧队列中的图片项，右侧将显示图片和 AI 识别出的问题列表。</li>
                    <li>每个问题都以“卡片”形式展现，可点击“编辑”按钮修改问题描述、风险等级、整改建议等，也可删除不准确的问题。</li>
                </ul>
            </li>
            <li><strong>人工标注（可选）</strong>：
                <ul>
                    <li>在图片预览区上方有绘图工具（框、圈、箭头、文字）。选择工具后，可在图片上直接进行手绘标注。</li>
                    <li>“🏷️引用问题”工具允许您在图片上添加文字标注，内容可从 AI 识别的问题列表中选择，方便快速关联。</li>
                    <li>“🤖 自动标识”按钮可以一键将所有 AI 识别出的、带有坐标的问题自动在图片上生成序号标注。</li>
                    <li>“撤销”、“删除选中”、“清空”用于管理您的标注。</li>
                    <li>点击“保存截图”可将当前带标注的图片保存为PNG文件，用于报告输出。</li>
                </ul>
            </li>
            <li><strong>导出报告</strong>：点击工具栏的“📄 导出报告”按钮，选择合适的报告模板（通用检查报告、整改通知单、简报模式），然后选择保存路径，即可生成 Word 报告。报告将包含所有问题描述和带有标注的图片。</li>
        </ol>

        <h4><strong>二、 高级设置（⚙ 按钮）</strong></h4>
        <ul>
            <li><strong>连接设置</strong>：配置 AI 模型厂商（如阿里百炼、硅基流动等）、API Key、Base URL、模型名称。您也可以选择“自定义”来设置任何兼容 OpenAI API 的服务。同时可设置最大并发数、重试次数和temperature。</li>
            <li><strong>提示词编辑</strong>：可查看和修改不同场景模式下，提供给 AI 的具体分析提示词。高级用户可根据需求调整，以获得更符合期望的分析结果。</li>
            <li><strong>业务数据配置</strong>：以 JSON 格式维护公司、项目、检查内容和项目概况等数据。这些数据会用于报告的基础信息填充。请确保 JSON 格式正确。</li>
            <li><strong>诊断</strong>：查看最近使用的检查人员和检查部位历史记录。</li>
        </ul>

        <h4><strong>三、 提示与注意事项</strong></h4>
        <ul>
            <li>为保证运行稳定，单次排查请控制在 {MAX_IMAGES} 张图片以内。</li>
            <li>AI 分析依赖于其识别能力，结果可能不完全准确，请务必人工复核和编辑。</li>
            <li>请确保在程序目录下存在所需的 Word 模板文件（如 `检查模板.docx`）。若缺失，报告将使用空白格式生成。</li>
            <li>“暂停”功能只会停止新任务的排队，正在分析中的任务仍会继续完成。</li>
            <li>“重试失败”功能可重新尝试分析状态为“失败”的任务。</li>
        </ul>
        """
        # 使用 QDialog 来展示 HTML 内容，提供更好的排版和滚动
        help_dialog = QDialog(self)
        help_dialog.setWindowTitle("帮助信息")
        help_dialog.resize(800, 700)

        dialog_layout = QVBoxLayout(help_dialog)
        text_browser = QTextEdit()
        text_browser.setHtml(help_content.format(MAX_IMAGES=MAX_IMAGES))  # 格式化MAX_IMAGES
        text_browser.setReadOnly(True)
        dialog_layout.addWidget(text_browser)

        close_button = QPushButton("关闭")
        close_button.clicked.connect(help_dialog.accept)
        dialog_layout.addWidget(close_button)

        help_dialog.exec()

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

            # 【修复】不要重新创建对象，而是清理现有场景
            self.image_view.scene().clear()
            # 重新添加底图 Item（因为 clear 会把它也删了）
            self.image_view._pix_item = QGraphicsPixmapItem()
            self.image_view._pix_item.setZValue(-1000)
            self.image_view.scene().addItem(self.image_view._pix_item)

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
        # 1. 检查 API Key
        if not self.config.get("api_key"):
            QMessageBox.warning(self, "缺 Key", "请在右上角设置中填写 API Key")
            return

        # 2. 保存当前界面的输入习惯
        self._remember_fields()

        # 3. 筛选需要处理的任务
        waiting = [t for t in self.tasks if t['status'] in ['waiting', 'error']]
        if not waiting:
            self.status_bar.showMessage("没有待处理的任务")
            return

        # 4. 加入等待队列
        for t in waiting:
            if t["id"] not in self.pending_queue and t["id"] not in self.running_workers:
                self.pending_queue.append(t["id"])
                t["status"] = "queued"
                self.update_list_color(t["id"], "#444444")

        # 5. 更新进度条状态
        self.progress_bar.setVisible(True)
        self.total_task = len([t for t in self.tasks if t["status"] in ["queued", "analyzing"]]) + len(
            self.running_workers)
        self.done_task = len([t for t in self.tasks if t["status"] == "done"])

        # 6. 触发调度器开始工作
        self._kick_scheduler()

    def _kick_scheduler(self):
        # 获取最大并发数配置
        max_conc = int(self.config.get("max_concurrency", 3))

        # 当运行中的任务少于最大并发数，且等待队列不为空时
        while len(self.running_workers) < max_conc and self.pending_queue:
            task_id = self.pending_queue.pop(0)
            task = next((t for t in self.tasks if t['id'] == task_id), None)
            if not task:
                continue

            # 获取提示词配置
            selected_template_name = self.cbo_prompt.currentText()
            prompts_dict = self.config.get("prompts", DEFAULT_PROMPTS)
            prompt_content = prompts_dict.get(selected_template_name, list(DEFAULT_PROMPTS.values())[0])

            # 更新任务状态
            task["status"] = "analyzing"
            task["error"] = None
            task["issues"] = []
            task["edited_issues"] = None
            task["export_image_path"] = None

            # 更新列表颜色为蓝色（进行中）
            self.update_list_color(task_id, "#0000FF")

            # 创建并启动后台线程
            worker = AnalysisWorker(task, self.config, prompt_content)

            # 【关键修复】这里必须连接到存在的 on_worker_done，而不是不存在的 on_worker_finished
            worker.result_ready.connect(self.on_worker_done)

            self.running_workers[task["id"]] = worker
            worker.start()

        # 更新进度条
        total = max(1, self.total_task)
        done = len([t for t in self.tasks if t["status"] == "done"])
        self.progress_bar.setValue(int(done / total * 100))

        # 检查是否全部完成
        if not self.running_workers and not self.pending_queue:
            self.status_bar.showMessage("✅ 队列分析完成")
            self.progress_bar.setValue(100)

    def on_worker_done(self, task_id: str, result):
        """后台线程完成回调（修复版）"""

        # 1. 更新任务数据
        task = next((t for t in self.tasks if t['id'] == task_id), None)
        if task:
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

            # 如果是当前选中的任务，安全渲染
            if self.current_task_id == task_id:
                QTimer.singleShot(50, lambda: self._safe_render_result(task))

        # 2. 安全销毁线程
        if task_id in self.running_workers:
            worker = self.running_workers.pop(task_id, None)
            if worker:
                try:
                    worker.result_ready.disconnect()
                except:
                    pass
                worker.quit()
                worker.wait(1000)  # 等待最多1秒
                worker.deleteLater()

        # 3. 🔧 修复：使用直接调用代替 QTimer
        self._kick_scheduler()

    def render_result(self, task: dict):
        """重新渲染结果面板（修复版：解决图片切换不显示的问题）"""

        # 1. 先清理右侧结果栏 (RiskCard)
        widgets_to_delete = []
        while self.result_layout.count():
            item = self.result_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widgets_to_delete.append(widget)

        for widget in widgets_to_delete:
            try:
                # 必须先断开信号，防止删除时触发回调导致崩溃
                widget.blockSignals(True)
                if isinstance(widget, RiskCard):
                    widget.edit_requested.disconnect()
                    widget.delete_requested.disconnect()
            except:
                pass
            widget.hide()
            widget.setParent(None)
            widget.deleteLater()

        # 强制刷新布局事件，确保旧控件被移除
        QApplication.processEvents()

        # 2. 【核心修复】不要禁用 image_view 的更新！
        # 之前的代码在这里调用了 self.image_view.setUpdatesEnabled(False)
        # 这导致 set_image 里的 fitInView 无法计算正确的缩放比例，导致图片消失。
        self.image_view.setUpdatesEnabled(True)

        # 3. 加载图片 (仅当路径变化时)
        # 确保路径存在且不为空
        img_path = task.get("path", "")
        if img_path and os.path.exists(img_path):
            if self.image_view._img_path != img_path:
                self.image_view.set_image(img_path)
        else:
            # 如果图片不存在（比如被删了），可以清空或显示占位
            pass

        # 4. 更新标注数据 (AI问题框 + 用户手绘)
        issues = task.get("edited_issues") if task.get("edited_issues") is not None else task.get("issues", [])
        self.image_view.set_ai_issues(issues)
        self.image_view.set_user_annotations(task.get("annotations", []) or [])

        # 5. 生成右侧问题卡片 (RiskCard)
        if task['status'] == 'done':
            if not issues:
                self.result_layout.addWidget(QLabel("✅ 未发现明显隐患"))
            else:
                for item_data in issues:
                    new_card = RiskCard(item_data)
                    new_card.edit_requested.connect(
                        lambda data=item_data: self.edit_issue(data)
                    )
                    new_card.delete_requested.connect(
                        lambda data=item_data: self.delete_issue(data)
                    )
                    self.result_layout.addWidget(new_card)

        elif task['status'] == 'analyzing':
            self.result_layout.addWidget(QLabel("⏳ 正在分析中..."))
        elif task['status'] == 'error':
            self.result_layout.addWidget(QLabel(f"❌ 失败: {task.get('error')}"))
        elif task['status'] == 'waiting':
            self.result_layout.addWidget(QLabel("🕒 等待分析..."))

        # 6. 强制刷新视图
        self.image_view.viewport().update()

    def edit_issue(self, item: Dict[str, Any]):
        """编辑问题项（修复版）"""
        task = self._current_task()
        if not task or task.get("status") != "done":
            return

        issues = task.get("edited_issues") if task.get("edited_issues") is not None else (task.get("issues") or [])

        dlg = IssueEditDialog(self, item)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            new_item = dlg.get_value()

            if task.get("edited_issues") is None:
                task["edited_issues"] = [dict(x) for x in issues]

            # 使用索引查找替换（更安全）
            replaced = False
            for i, x in enumerate(task["edited_issues"]):
                if x.get("issue") == item.get("issue") and x.get("risk_level") == item.get("risk_level"):
                    task["edited_issues"][i] = new_item
                    replaced = True
                    break

            if not replaced:
                task["edited_issues"].append(new_item)

            task["export_image_path"] = None

            # 🔧 修复：延迟刷新
            QTimer.singleShot(100, lambda: self._safe_render_result(task))

    def delete_issue(self, item: Dict[str, Any]):
        """安全删除问题项，避免信号槽冲突"""
        task = self._current_task()
        if not task:
            return

        # 🔧 修复：先断开所有信号，再更新数据
        sender_card = self.sender()
        if sender_card and isinstance(sender_card, RiskCard):
            try:
                sender_card.blockSignals(True)  # 阻止后续信号
                sender_card.edit_requested.disconnect()
                sender_card.delete_requested.disconnect()
            except:
                pass

        # 更新数据模型（使用深拷贝避免引用问题）
        issues = task.get("edited_issues") if task.get("edited_issues") is not None else (task.get("issues") or [])
        if task.get("edited_issues") is None:
            task["edited_issues"] = [dict(x) for x in issues]

        # 安全过滤（使用 id() 比较对象身份）
        task["edited_issues"] = [x for x in task["edited_issues"] if id(x) != id(item)]
        task["export_image_path"] = None

        # 更新 ImageView 数据
        self.image_view.set_ai_issues(task["edited_issues"])

        # 🔧 修复：使用更长的延迟确保 Qt 事件循环完全清理
        QTimer.singleShot(150, lambda: self._safe_render_result(task))

        self.status_bar.showMessage("已删除该问题项", 2000)

    def _safe_render_result(self, task: dict):
        """安全的渲染包装器，捕获所有异常"""
        try:
            self.render_result(task)
        except RuntimeError as e:
            print(f"⚠️ 渲染时发生 RuntimeError（对象已销毁）: {e}")
        except Exception as e:
            print(f"❌ 渲染时发生未知错误: {e}\n{traceback.format_exc()}")

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
