import sys
import os
import json
import base64
import time
from datetime import datetime
from openai import OpenAI

# ================= 1. Word 与 UI 库导入 =================
from docx import Document
# 关键：导入 Cm 用于设置固定宽度
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLabel, QListWidget, QListWidgetItem, QSplitter,
                             QScrollArea, QFrame, QFileDialog, QProgressBar, QMessageBox,
                             QDialog, QFormLayout, QLineEdit, QComboBox, QToolBar,
                             QSizePolicy, QTabWidget, QTextEdit, QGroupBox, QGridLayout)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QPixmap, QIcon, QColor, QAction

CONFIG_FILE = "app_config_lec.json"

# ================= 2. 核心默认数据配置 (首次运行时的默认值) =================

DEFAULT_BUSINESS_DATA = {
    # 1. 公司与项目的映射关系 (下拉框二级联动)
    "company_project_map": {
        "勐海县泽兴供水有限公司": ["城乡供水一体化项目", "勐海农村供水保障项目"],
        "勐海县润博水利投资有限公司": ["勐阿水库建设项目"],
        "江城县润成水利投资有限公司": ["热水河水库建设项目"],
        "澜沧县润成水利投资有限公司": ["三道箐水库建设项目"]
    },
    # 2. 公司与被检查单位的映射关系 (自动填充)
    "company_unit_map": {
        "勐海县泽兴供水有限公司": "云南建投第二水利水电建设有限公司",
        "勐海县润博水利投资有限公司": "云南建投第二水利水电建设有限公司",
        "江城县润成水利投资有限公司": "云南建投第二水利水电建设有限公司",
        "澜沧县润成水利投资有限公司": "云南省水利水电工程有限公司"
    },
    # 3. 检查内容预设选项
    "check_content_options": [
        "安全文明施工专项检查",
        "工程质量专项检查",
        "项目综合检查",
        "节前安全生产检查",
        "复工复产专项检查"
    ],
    # 4. 项目概况详细信息映射
    "project_overview_map": {
        "勐海农村供水保障项目": "本工程位于西双版纳州勐海县，主要建设内容包括新建取水坝、输水管网及配套水厂设施，旨在解决周边5个乡镇的农村饮水安全问题，设计供水规模为2.5万吨/日。",
        "城乡供水一体化项目": "勐海县城乡供水一体化项目共包含 9 个片区的供水工程，主要建设内容为取水设施、水处理厂、提水泵站、原水、清水输配水管网及其建筑物等。本次新建水处理厂3座，改扩建1座",
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
必须严格返回 JSON 数组，不要 Markdown 标记。`risk_level` 必须包含“严重”、“一般”或“文明施工”字样以触发颜色警告。

[
    {
        "risk_level": "严重安全隐患", 
        "issue": "画面右侧工人站在移动脚手架顶端作业，未佩戴安全带，且脚手架无护栏，存在极高坠落风险",
        "regulation": "《建筑施工高处作业安全技术规范》JGJ 80-2016 第3.0.5条",
        "correction": "立即停止作业，补齐防护栏杆，作业人员必须正确系挂全身式安全带"
    },
    {
        "risk_level": "一般质量缺陷", 
        "issue": "新砌筑的填充墙顶部，斜砌砖角度过小且灰缝不饱满，易导致后期裂缝",
        "regulation": "《砌体结构工程施工质量验收规范》GB 50203",
        "correction": "拆除顶部不合格砌块，待下部砌体沉实后，采用标准角度斜砌挤紧"
    },
    {
        "risk_level": "文明施工问题", 
        "issue": "钢管扣件随意堆放在通道上，且混有生活垃圾，未进行分类归库，影响通行且形象差",
        "regulation": "《建设工程施工现场环境与卫生标准》JGJ 146",
        "correction": "立即清理通道垃圾，钢管扣件按规格分类堆放并设置标识牌"
    }
]""",

    "🏠 纯日常生活 (整理/健康/居家)": """你是一位资深的生活管家。请以提升生活品质为目标，分析照片中的场景。

### 输出格式 (JSON)
[
    {
        "risk_level": "卫生警示", 
        "issue": "冰箱冷藏室内的剩菜未覆盖保鲜膜，且与新鲜水果混放，存在细菌交叉感染风险",
        "regulation": "食品卫生常识",
        "correction": "使用保鲜盒密封剩菜，并建议划分生熟食存放区域"
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


# ================= 3. 配置管理 (核心优化) =================

class ConfigManager:
    @staticmethod
    def get_default_config():
        """返回完整的默认配置字典"""
        return {
            "current_provider": "阿里百炼 (Qwen-VL)",
            "api_key": "",
            "last_prompt": list(DEFAULT_PROMPTS.keys())[0],
            "custom_provider_settings": {"base_url": "", "model": ""},
            # 将核心业务数据也放入配置
            "business_data": DEFAULT_BUSINESS_DATA,
            "prompts": DEFAULT_PROMPTS,
            "provider_presets": DEFAULT_PROVIDER_PRESETS
        }

    @staticmethod
    def load():
        default = ConfigManager.get_default_config()

        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    saved = json.load(f)

                    # 深度合并逻辑：确保配置项完整
                    if "business_data" not in saved:
                        saved["business_data"] = default["business_data"]
                    else:
                        # 确保业务数据内部的键完整
                        for key in default["business_data"]:
                            if key not in saved["business_data"]:
                                saved["business_data"][key] = default["business_data"][key]

                    if "prompts" not in saved:
                        saved["prompts"] = default["prompts"]

                    if "provider_presets" not in saved:
                        saved["provider_presets"] = default["provider_presets"]

                    return {**default, **saved}
            except Exception as e:
                print(f"配置文件加载失败，使用默认值: {e}")
                pass
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


# ================= 4. Word 报告生成器 =================

class WordReportGenerator:
    @staticmethod
    def set_font(run, font_name_cn='宋体', font_name_en='Times New Roman', size=12, bold=False, color=None):
        """统一设置中英文字体格式"""
        run.font.name = font_name_en
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name_cn)
        if size:
            run.font.size = Pt(size)
        run.font.bold = bold
        if color:
            run.font.color.rgb = color

    @staticmethod
    def _replace_text_in_paragraph(paragraph, replacements):
        """
        核心修复：增加 Fallback 机制
        解决 Word 将 {{占位符}} 切割成多个 Run 导致无法匹配的问题
        """
        if not paragraph.text:
            return

        for key, value in replacements.items():
            if key in paragraph.text:
                val_str = str(value) if value else ""
                replaced_in_run = False

                # 1. 尝试在保留格式的 Run 级别进行替换
                for run in paragraph.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, val_str)
                        WordReportGenerator.set_font(run, size=12, bold=run.font.bold)
                        replaced_in_run = True

                # 2. 【核心修复】如果 Run 级别没换成功（说明占位符被Word切碎了），强制在段落级替换
                if not replaced_in_run:
                    # 强行替换段落文本（注意：这会重置该段落内部分文字的特殊格式，但在表头中通常没问题）
                    paragraph.text = paragraph.text.replace(key, val_str)
                    # 重新给新生成的段落应用字体
                    for run in paragraph.runs:
                        WordReportGenerator.set_font(run, size=12, bold=run.font.bold)

    @staticmethod
    def replace_placeholders(doc, info):
        """遍历文档进行注入"""
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

        # 1. 遍历正文段落
        for para in doc.paragraphs:
            WordReportGenerator._replace_text_in_paragraph(para, replacements)

        # 2. 遍历表格 (绝大多数表头信息都在这里)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        WordReportGenerator._replace_text_in_paragraph(para, replacements)

    @staticmethod
    def generate(tasks, save_path, project_info):
        template_name = "模板.docx"
        if os.path.exists(template_name):
            doc = Document(template_name)
        else:
            doc = Document()
            # 设置默认页边距等...
            section = doc.sections[0]
            section.top_margin = Cm(2.0)
            section.bottom_margin = Cm(2.0)
            section.left_margin = Cm(2.0)
            section.right_margin = Cm(2.0)
            doc.add_paragraph("【注意】未找到模板文件，表头信息未填入。请在同级目录下放入 模板.docx")

        # 执行替换
        WordReportGenerator.replace_placeholders(doc, project_info)

        # 移动到文末
        doc.add_paragraph()

        # 循环写入点位 (紧凑模式 + 问题 X 标题)
        for idx, task in enumerate(tasks, 1):
            table = doc.add_table(rows=1, cols=1)
            table.style = 'Table Grid'
            table.autofit = False
            cell = table.cell(0, 0)
            cell.width = Cm(17.0)

            # 标题
            p_title = cell.paragraphs[0]
            p_title.paragraph_format.space_before = Pt(4)
            p_title.paragraph_format.space_after = Pt(4)
            p_title.paragraph_format.left_indent = Cm(0.2)

            run_title = p_title.add_run(f"问题 {idx}")
            WordReportGenerator.set_font(run_title, size=12, bold=True)

            # 数据处理
            data = task.get('data', [])
            safety_texts = []
            quality_texts = []
            all_corrections = []

            if not data or isinstance(data, str) or len(data) == 0:
                safety_texts.append("无明显隐患")
                all_corrections.append("无")
            else:
                for item in data:
                    r_level = item.get("risk_level", "")
                    issue = item.get("issue", "").strip()
                    reg = item.get("regulation", "").strip()
                    corr = item.get("correction", "").strip()

                    full_desc = issue
                    if reg and reg not in ["无", "常识"]:
                        full_desc += f"（违反 {reg}）"

                    if "质量" in r_level:
                        quality_texts.append(full_desc)
                    else:
                        safety_texts.append(full_desc)
                    all_corrections.append(corr)

            # 写入内容 - 安全
            if safety_texts:
                p = cell.add_paragraph()
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.left_indent = Cm(0.2)
                p.paragraph_format.right_indent = Cm(0.2)
                p.paragraph_format.line_spacing = 1.2
                run_label = p.add_run("安全/文明施工问题：")
                WordReportGenerator.set_font(run_label, bold=True, size=11)
                merged_txt = "；".join(safety_texts) + "。"
                run_text = p.add_run(merged_txt)
                WordReportGenerator.set_font(run_text, size=11)

            # 写入内容 - 质量
            if quality_texts:
                p = cell.add_paragraph()
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.left_indent = Cm(0.2)
                p.paragraph_format.right_indent = Cm(0.2)
                p.paragraph_format.line_spacing = 1.2
                run_label = p.add_run("质量问题：")
                WordReportGenerator.set_font(run_label, bold=True, size=11)
                merged_txt = "；".join(quality_texts) + "。"
                run_text = p.add_run(merged_txt)
                WordReportGenerator.set_font(run_text, size=11)

            # 写入内容 - 整改要求
            p_corr = cell.add_paragraph()
            p_corr.paragraph_format.space_before = Pt(2)
            p_corr.paragraph_format.space_after = Pt(2)
            p_corr.paragraph_format.left_indent = Cm(0.2)
            p_corr.paragraph_format.right_indent = Cm(0.2)
            p_corr.paragraph_format.line_spacing = 1.2
            run_label = p_corr.add_run("整改要求：")
            WordReportGenerator.set_font(run_label, bold=True, size=11)
            merged_corr = "；".join(all_corrections) + "。"
            run_text = p_corr.add_run(merged_corr)
            WordReportGenerator.set_font(run_text, size=11, color=RGBColor(0, 100, 0))

            # 插入图片
            if os.path.exists(task['path']):
                p_img = cell.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(4)
                p_img.paragraph_format.space_after = Pt(4)
                try:
                    p_img.add_run().add_picture(task['path'], width=Cm(13.5))
                except:
                    p_img.add_run("[图片加载失败]")

            # 紧凑空行
            if idx < len(tasks):
                spacer = doc.add_paragraph()
                spacer.paragraph_format.space_after = Pt(10)

        try:
            doc.save(save_path)
        except Exception as e:
            raise e


# ================= 5. 后台分析线程 =================

class AnalysisWorker(QThread):
    finished = pyqtSignal(str, object)

    def __init__(self, task, config, prompt_text):
        super().__init__()
        self.task = task
        self.config = config
        self.prompt_text = prompt_text

    def run(self):
        try:
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

            if not api_key:
                self.finished.emit(self.task['id'], {"error": "未配置 API Key"})
                return
            if not base_url or not model:
                self.finished.emit(self.task['id'], {"error": "未配置模型 URL 或 名称"})
                return

            client = OpenAI(api_key=api_key, base_url=base_url)

            with open(self.task['path'], "rb") as f:
                b64 = base64.b64encode(f.read()).decode()

            resp = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": self.prompt_text},
                    {"role": "user", "content": [
                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}},
                        {"type": "text", "text": "请分析"}
                    ]}
                ],
                temperature=0.1
            )

            content = resp.choices[0].message.content
            clean = content.replace("```json", "").replace("```", "").strip()
            s = clean.find('[')
            e = clean.rfind(']') + 1
            if s != -1 and e != -1:
                self.finished.emit(self.task['id'], json.loads(clean[s:e]))
            else:
                self.finished.emit(self.task['id'], [])

        except Exception as e:
            self.finished.emit(self.task['id'], {"error": str(e)})


# ================= 6. UI 组件 =================

class RiskCard(QFrame):
    def __init__(self, item):
        super().__init__()
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
            f"RiskCard {{ background-color: {bg}; border-left: 5px solid {bd}; border-radius: 4px; margin-bottom: 5px; }}")

        layout = QVBoxLayout(self)
        header = QHBoxLayout()
        header.addWidget(QLabel(f"<b>[{level}]</b>"))
        lbl_issue = QLabel(item.get("issue", ""))
        lbl_issue.setWordWrap(True)
        header.addWidget(lbl_issue, 1)
        layout.addLayout(header)

        layout.addWidget(QLabel(f"依据: {item.get('regulation', '')}"))
        lbl_fix = QLabel(f"建议: {item.get('correction', '')}")
        lbl_fix.setStyleSheet("color: #2E7D32; font-weight: bold;")
        lbl_fix.setWordWrap(True)
        layout.addWidget(lbl_fix)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = ConfigManager.load()
        self.refresh_business_data()  # 初始化加载业务数据

        self.tasks = []
        self.queue_workers = []
        self.current_task_id = None
        self.init_ui()

    def refresh_business_data(self):
        """从配置刷新本地业务数据缓存"""
        self.business_data = self.config.get("business_data", DEFAULT_BUSINESS_DATA)

    def init_ui(self):
        self.setWindowTitle("普洱版纳区域检查报告助手")
        self.resize(1300, 950)

        # --- 工具栏 ---
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

        btn_clear = QAction("🗑️ 清空队列", self)
        btn_clear.triggered.connect(self.clear_queue)
        toolbar.addAction(btn_clear)

        btn_export = QAction(QIcon(), "📄 导出Word报告", self)
        btn_export.triggered.connect(self.export_word)
        toolbar.addAction(btn_export)

        empty = QWidget()
        empty.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        toolbar.addWidget(empty)

        btn_setting = QAction("⚙ 设置", self)
        btn_setting.triggered.connect(self.open_settings)
        toolbar.addAction(btn_setting)

        # =========================================================
        # 顶部：基础信息输入区 (从配置加载)
        # =========================================================
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        info_group = QGroupBox("📄 报告基础信息 (数据源可配置)")
        info_group.setFixedHeight(180)
        info_layout = QGridLayout(info_group)
        info_layout.setContentsMargins(10, 10, 10, 10)

        # 1. 公司名称
        self.input_company = QComboBox()
        # 初始化加载
        self.update_company_combo()
        self.input_company.setEditable(False)

        # 2. 项目名称
        self.input_project = QComboBox()
        self.input_project.setEditable(False)

        # 3. 被检查单位
        self.input_inspected_unit = QLineEdit()
        self.input_inspected_unit.setPlaceholderText("自动生成，也可手动修改")

        # 4. 检查内容
        self.input_check_content = QComboBox()
        self.update_check_content_combo()
        self.input_check_content.setEditable(True)

        # 5. 其他字段
        self.input_area = QLineEdit()
        self.input_area.setPlaceholderText("例如：乡镇或者枢纽、隧洞等")

        self.input_person = QLineEdit()
        self.input_person.setPlaceholderText("请输入检查人姓名")

        self.input_date = QLineEdit()
        self.input_date.setText(datetime.now().strftime("%Y-%m-%d"))

        self.input_deadline = QLineEdit()
        self.input_deadline.setPlaceholderText("例如：2025-12-30 ")
        # 信号连接
        self.input_company.currentTextChanged.connect(self.on_company_changed)
        # 初始触发
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
        main_layout.addWidget(info_group)

        # =========================================================
        # 下方：列表 + 结果
        # =========================================================
        splitter = QSplitter(Qt.Orientation.Horizontal)

        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        self.lbl_count = QLabel("待审队列 (0/20)")
        left_layout.addWidget(self.lbl_count)
        self.list_widget = QListWidget()
        self.list_widget.itemClicked.connect(self.on_item_clicked)
        left_layout.addWidget(self.list_widget)

        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)

        self.lbl_image = QLabel("请从左侧选择图片")
        self.lbl_image.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_image.setStyleSheet("background-color: #333; color: #AAA; border-radius: 6px;")
        self.lbl_image.setMinimumHeight(400)
        right_layout.addWidget(self.lbl_image, 1)

        self.result_container = QWidget()
        self.result_layout = QVBoxLayout(self.result_container)
        self.result_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self.result_container)
        right_layout.addWidget(scroll, 1)

        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([350, 950])
        main_layout.addWidget(splitter)

        self.status_bar = self.statusBar()
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedWidth(200)
        self.status_bar.addPermanentWidget(self.progress_bar)

    # --- 辅助刷新 UI ---
    def update_company_combo(self):
        current_text = self.input_company.currentText()
        self.input_company.blockSignals(True)
        self.input_company.clear()
        company_map = self.business_data.get("company_project_map", {})
        self.input_company.addItems(company_map.keys())
        # 尝试恢复之前的选择
        index = self.input_company.findText(current_text)
        if index >= 0:
            self.input_company.setCurrentIndex(index)
        elif self.input_company.count() > 0:
            self.input_company.setCurrentIndex(0)
        self.input_company.blockSignals(False)

    def update_check_content_combo(self):
        current_text = self.input_check_content.currentText()
        self.input_check_content.clear()
        check_options = self.business_data.get("check_content_options", [])
        self.input_check_content.addItems(check_options)
        self.input_check_content.setEditText(current_text)

    # --- 逻辑 ---
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
        if not text: return
        self.config["last_prompt"] = text
        ConfigManager.save(self.config)

    def add_files(self):
        current_count = len(self.tasks)
        if current_count >= 20:
            QMessageBox.warning(self, "数量限制", "为保证运行稳定，单次排查请控制在 20 张图片以内。\n建议先清空队列。")
            return

        remaining = 20 - current_count
        paths, _ = QFileDialog.getOpenFileNames(self, f"选择图片 (还能选 {remaining} 张)", "",
                                                "Images (*.jpg *.png *.jpeg)")

        if not paths: return

        if len(paths) > remaining:
            QMessageBox.warning(self, "超限提示", f"你选择了 {len(paths)} 张，自动截取前 {remaining} 张。")
            paths = paths[:remaining]

        for path in paths:
            if any(t['path'] == path for t in self.tasks): continue
            task_id = str(time.time()) + os.path.basename(path)
            self.tasks.append(
                {"id": task_id, "path": path, "name": os.path.basename(path), "status": "waiting", "data": None})
            item = QListWidgetItem(os.path.basename(path))
            item.setData(Qt.ItemDataRole.UserRole, task_id)
            self.list_widget.addItem(item)
        self.lbl_count.setText(f"待审队列 ({len(self.tasks)}/20)")

    def clear_queue(self):
        if any(t['status'] == 'analyzing' for t in self.tasks):
            QMessageBox.warning(self, "警告", "任务正在分析中，请等待完成后再清空！")
            return
        reply = QMessageBox.question(self, '确认', '确定要清空所有待审任务吗？',
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.tasks.clear()
            self.list_widget.clear()
            self.lbl_count.setText("待审队列 (0/20)")
            self.lbl_image.clear()
            self.lbl_image.setText("请从左侧选择图片")
            self.current_task_id = None
            while self.result_layout.count():
                child = self.result_layout.takeAt(0)
                if child.widget(): child.widget().deleteLater()
            self.status_bar.showMessage("队列已清空")

    def start_analysis(self):
        if not self.config.get("api_key"):
            QMessageBox.warning(self, "缺 Key", "请在右上角设置中填写 API Key")
            return

        waiting = [t for t in self.tasks if t['status'] in ['waiting', 'error']]
        if not waiting:
            self.status_bar.showMessage("没有待处理的任务")
            return

        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.total_task = len(waiting)
        self.done_task = 0

        selected_template_name = self.cbo_prompt.currentText()
        prompts_dict = self.config.get("prompts", DEFAULT_PROMPTS)
        prompt_content = prompts_dict.get(selected_template_name, list(DEFAULT_PROMPTS.values())[0])

        for task in waiting:
            task['status'] = 'analyzing'
            self.update_list_color(task['id'], "#0000FF")
            worker = AnalysisWorker(task, self.config, prompt_content)
            worker.finished.connect(self.on_worker_done)
            worker.start()
            self.queue_workers.append(worker)

    def on_worker_done(self, task_id, result):
        task = next((t for t in self.tasks if t['id'] == task_id), None)
        if task:
            if isinstance(result, dict) and "error" in result:
                task['status'] = 'error'
                task['data'] = result["error"]
                self.update_list_color(task_id, "#FF0000")
            else:
                task['status'] = 'done'
                task['data'] = result
                self.update_list_color(task_id, "#008000")

            if self.current_task_id == task_id:
                self.render_result(task)

        self.done_task += 1
        self.progress_bar.setValue(int(self.done_task / self.total_task * 100))

        if self.done_task == self.total_task:
            self.status_bar.showMessage("✅ 队列分析完成")
            self.queue_workers.clear()

    def render_result(self, task):
        while self.result_layout.count():
            child = self.result_layout.takeAt(0)
            if child.widget(): child.widget().deleteLater()
        if not task: return
        if task['status'] == 'analyzing':
            self.result_layout.addWidget(QLabel("🚀 正在智能分析中 (全模态)，请稍候..."))
        elif task['status'] == 'done':
            if not task['data']:
                self.result_layout.addWidget(QLabel("✅ 完美：未发现明显隐患或改进项"))
            else:
                for item in task['data']:
                    self.result_layout.addWidget(RiskCard(item))

    def on_item_clicked(self, item):
        task_id = item.data(Qt.ItemDataRole.UserRole)
        self.current_task_id = task_id
        task = next((t for t in self.tasks if t['id'] == task_id), None)
        pix = QPixmap(task['path'])
        scaled = pix.scaled(self.lbl_image.size(), Qt.AspectRatioMode.KeepAspectRatio,
                            Qt.TransformationMode.SmoothTransformation)
        self.lbl_image.setPixmap(scaled)
        self.render_result(task)

    def update_list_color(self, task_id, color):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == task_id:
                item.setForeground(QColor(color))

    def export_word(self):
        if not self.tasks: return
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

        current_time_str = datetime.now().strftime('%Y%m%d_%H%M%S')
        prefix = project_info['project_name'] if project_info['project_name'] else "智能排查报告"
        default_name = f"{prefix}_{current_time_str}.docx"

        path, _ = QFileDialog.getSaveFileName(self, "保存报告", default_name, "Word Files (*.docx)")
        if not path: return

        try:
            WordReportGenerator.generate(self.tasks, path, project_info)
            QMessageBox.information(self, "成功", f"报告已生成！\n路径：{path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))

    def open_settings(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("系统高级设置")
        dlg.resize(700, 600)

        tabs = QTabWidget()

        # --- Tab 1: 连接设置 ---
        tab_conn = QWidget()
        layout_conn = QFormLayout(tab_conn)
        provider_presets = self.config.get("provider_presets", DEFAULT_PROVIDER_PRESETS)

        cbo_provider = QComboBox()
        cbo_provider.addItems(provider_presets.keys())
        curr_prov = self.config.get("current_provider")
        if curr_prov not in provider_presets: curr_prov = list(provider_presets.keys())[0]
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
        tabs.addTab(tab_conn, "🔌 连接设置")

        # --- Tab 2: 提示词编辑 ---
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
        if self._temp_last_selected_prompt: load_prompt(self._temp_last_selected_prompt)

        layout_prompt.addWidget(QLabel("选择模板进行编辑:"))
        layout_prompt.addWidget(cbo_template)
        layout_prompt.addWidget(txt_prompt_edit)
        tabs.addTab(tab_prompt, "📝 提示词编辑")

        # --- Tab 3: [新增] 业务数据配置 (直接修改 JSON) ---
        tab_data = QWidget()
        layout_data = QVBoxLayout(tab_data)

        lbl_info = QLabel(
            "此处配置公司名称、项目名称、被检单位及项目概况。\n请保持 JSON 格式正确 (注意双引号和逗号)。修改后点击保存即可生效。")
        lbl_info.setWordWrap(True)
        txt_data_edit = QTextEdit()

        # 加载当前业务数据并格式化显示
        current_biz_data = self.config.get("business_data", DEFAULT_BUSINESS_DATA)
        txt_data_edit.setText(json.dumps(current_biz_data, indent=4, ensure_ascii=False))

        layout_data.addWidget(lbl_info)
        layout_data.addWidget(txt_data_edit)
        tabs.addTab(tab_data, "📊 业务数据配置")

        # --- 按钮 ---
        btn_box = QHBoxLayout()
        btn_save = QPushButton("保存所有配置")
        btn_save.setMinimumHeight(40)
        btn_save.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; border-radius: 4px;")
        btn_cancel = QPushButton("取消")

        def save_all():
            try:
                # 1. 保存提示词
                save_prompt_to_mem()

                # 2. 尝试解析并保存业务数据 (Tab 3)
                raw_json = txt_data_edit.toPlainText()
                new_biz_data = json.loads(raw_json)  # 校验JSON格式

                # 3. 收集连接设置
                self.config["current_provider"] = cbo_provider.currentText()
                self.config["api_key"] = txt_key.text().strip()
                self.config["prompts"] = local_prompts
                self.config["business_data"] = new_biz_data  # 更新业务数据

                if cbo_provider.currentText() == "自定义 (Custom)":
                    self.config["custom_provider_settings"] = {
                        "base_url": txt_base_url.text().strip(),
                        "model": txt_model.text().strip()
                    }

                ConfigManager.save(self.config)

                # 4. 刷新主界面 UI
                self.refresh_business_data()
                self.update_company_combo()
                self.update_check_content_combo()
                # 触发一次公司变更以更新项目
                self.on_company_changed(self.input_company.currentText())

                # 刷新 Prompt 下拉
                self.cbo_prompt.blockSignals(True)
                curr = self.cbo_prompt.currentText()
                self.cbo_prompt.clear()
                self.cbo_prompt.addItems(self.config["prompts"].keys())
                if curr in self.config["prompts"]: self.cbo_prompt.setCurrentText(curr)
                self.cbo_prompt.blockSignals(False)

                dlg.accept()
                self.status_bar.showMessage("✅ 配置已保存，公司项目列表已更新")

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
