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
                             QSizePolicy, QTabWidget, QTextEdit, QGroupBox)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QPixmap, QIcon, QColor, QAction

CONFIG_FILE = "app_config_lec.json"

# ================= 2. 提示词模板 (LEC 评测法升级版) =================

DEFAULT_PROMPTS = {
    # =========================================================================
    # 第一类：工程建设专项 (工作核心) - 安全用LEC，质量用GB规范
    # =========================================================================

"🏗️ 施工安质全能扫描 (LEC+实体质量)": """你是一位拥有30年经验的“注册安全工程师”及“总监理工程师”。请对施工现场照片进行“安全隐患+工程质量”的全方位深度排查。

### 一、 核心评分标准 (严格执行双轨制)

**1. 安全风险 (必须使用 LEC法 量化)**
   - 公式：D = L(可能性) × E(暴露频率) × C(后果严重度)
   - **L (Likelihood)**: 10(完全可能/常发), 6(相当可能), 3(可能/偶然), 1(可能性小).
   - **E (Exposure)**: 10(连续暴露), 6(每日工作时间/常驻), 3(每周一次), 1(极少). *注：施工现场隐患 E值通常默认为 6 或 10*。
   - **C (Consequence)**: 100分：10人以上死亡。40分：3～9人死亡。15分：1～2人死亡。7分：严重事故。3分：重大伤残。.
   - **分级阈值**:
     - **重大风险 (D ≥ 320)**: 必须立即停工整改。
     - **较大风险 (160≤ D < 320)**: 需限期整改。
     - **一般风险 (70 ≤ D < 160)**: 日常维护问题。
     - **低风险 (D < 70)**: 日常维护问题。
**2. 质量缺陷 (依据 GB验收规范 定性)**
   - **重大质量隐患**: 影响结构安全、承载力或主要使用功能 (例: 严重烂根/露筋、贯穿裂缝、钢筋数量不足、特种设备关键部件缺失)。
   - **较大质量缺陷**: 影响耐久性或外观质量极差 (例: 大面积蜂窝麻面、钢筋间距严重不匀、保护层垫块缺失、连接套筒露丝过长)。
   - **一般质量通病**: 轻微外观瑕疵 (例: 模板拼缝漏浆、砖墙灰缝不直、轻微浮锈)。

### 二、 重点排查清单 (像素级扫描)

**1. 特种设备与危大工程 (红线必查)**
   - **起重机械**: 塔吊/施工升降机/汽车吊。重点查：**支腿是否垫实(防倾覆)**、**钢丝绳断丝/锈蚀**、**吊钩防脱棘爪**、**限位器/力矩限制器**、**附着装置**。
   - **深基坑**: 边坡支护变形、坑边堆载过大、临边防护缺失、降排水失效。
   - **高处作业**: 脚手架(立杆/扫地杆/连墙件/剪刀撑)、吊篮(安全锁/配重/生命绳)。
    - **人员**: 人员防护用品、安全管理人员。
**2. 实体工程质量 (质量必查)**
   - **混凝土**: 蜂窝、麻面、孔洞、夹渣、露筋、烂根、缺棱掉角、裂缝。
   - **钢筋**: 绑扎间距、搭接长度、锚固长度、直螺纹连接(露丝<2扣)、除锈情况、保护层垫块。
   - **砌体/模板**: 马牙槎留置、灰缝饱满度、顶砖斜砌、模板对拉螺栓、支撑体系稳定性。

**3. 通用安全**: 临电(一机一闸一漏)、动火(气瓶间距/灭火器)、PPE佩戴。

### 三、 输出格式 (JSON)
请严格按此格式返回，不要包含 Markdown 标记：
[
    {
        "risk_level": "重大风险 (D=240)", 
        "issue": "【安全-特种设备】汽车吊右后支腿下方土地松软且未垫设枕木，L=6, E=10, C=40，存在极高倾覆风险",
        "regulation": "违反《建筑机械使用安全技术规程》JGJ 33 第4.4.2条",
        "correction": "立即停止吊装，重新平整场地并铺设标准路基箱或枕木"
    },
    {
        "risk_level": "重大质量隐患", 
        "issue": "【质量-混凝土】剪力墙底部存在严重烂根及露筋(长度>30cm)，影响结构承载力",
        "regulation": "违反《混凝土结构工程施工质量验收规范》GB 50204 第8.2.1条",
        "correction": "经设计/监理确认方案后，凿除松散层，用高一等级微膨胀砂浆修补并养护"
    }
]""",

    # =========================================================================
    # 第二类：日常办公与生活专项 (行政/后勤/居家)
    # =========================================================================

    "🏠 纯日常生活 (整理/健康/居家)": """你是一位资深的生活管家、收纳师及营养师。请以提升生活品质为目标，分析照片中的场景。不要过分强调工业安全，而是关注整洁度、生活习惯与健康。

### 一、 评价标准 (生活化分级)

**1. 🔴 卫生/健康警示 (对应重大风险色)**
   - 定义：严重影响健康或生活质量的问题。
   - 场景：食材发霉变质、严重的卫生死角(霉斑/油污)、家里有明显的跌倒/割伤隐患(针对老人儿童)、过期药品。

**2. 🟠 需整理/需改善 (对应较大风险色)**
   - 定义：视觉上杂乱、使用不便或轻度浪费。
   - 场景：衣物堆积如山、桌面杂物过多、收纳逻辑混乱、冰箱生熟不分、电源线缠绕凌乱。

**3. 🔵 生活建议/美化 (对应一般风险色)**
   - 定义：锦上添花的优化建议。
   - 场景：色彩搭配建议、增加绿植、灯光氛围优化、家具摆放调整。

### 二、 检查重点 (生活场景)

**1. 居家环境与收纳**
   - **整洁度**: 地面/桌面是否有大量杂物？床铺是否整理？
   - **收纳逻辑**: 物品是否分类归位？常用物品是否顺手？是否存在“无效堆积”？
   - **家居维护**: 墙面是否有污渍/裂纹？灯泡是否损坏？

**2. 饮食与健康**
   - **食材**: 水果蔬菜是否新鲜？是否存在高糖/高油的不健康食品堆积？
   - **厨房**: 碗筷是否沥水？调料瓶是否油腻？冰箱内部是否杂乱？

**3. 舒适与美学**
   - **氛围**: 光线是否昏暗？是否缺乏生活气息？
   - **布局**: 家具摆放是否阻碍动线？

### 三、 输出格式 (JSON)
[
    {
        "risk_level": "卫生警示", 
        "issue": "【食品健康】冰箱冷藏室内的剩菜未覆盖保鲜膜，且与新鲜水果混放，存在细菌交叉感染风险",
        "regulation": "食品卫生与保鲜常识",
        "correction": "使用保鲜盒密封剩菜，并建议划分生熟食存放区域"
    },
    {
        "risk_level": "需整理", 
        "issue": "【居家收纳】书桌表面堆放了过多的书籍、数据线和水杯，占用作业空间且视觉杂乱",
        "regulation": "断舍离与桌面收纳原则",
        "correction": "建议使用桌面收纳盒归类文具，将不常用的书籍归入书架"
    },
    {
        "risk_level": "生活建议", 
        "issue": "【家居美学】客厅沙发区域色调过于单一，缺乏视觉焦点",
        "regulation": "家居软装搭配技巧",
        "correction": "建议增加两个亮色抱枕或铺设一块暖色地毯，提升温馨感"
    }
]"""
}

# 模型厂商预设
PROVIDER_PRESETS = {
    "阿里百炼 (Qwen-VL)": {
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "model": "qwen-vl-max"
    },
    "硅基流动 (SiliconFlow)": {
        "base_url": "https://api.siliconflow.cn/v1",
        "model": "Qwen/Qwen2-VL-72B-Instruct"
    },
    "字节豆包 (Doubao)": {
        "base_url": "https://ark.cn-beijing.volces.com/api/v3",
        "model": "ep-2024xxxxxx-xxxxx"
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
    def load():
        default = {
            "current_provider": "阿里百炼 (Qwen-VL)",
            "api_key": "",
            "last_prompt": list(DEFAULT_PROMPTS.keys())[0],  # 默认选中第一个
            "prompts": DEFAULT_PROMPTS.copy(),
            "custom_provider_settings": {"base_url": "", "model": ""}
        }

        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    saved = json.load(f)

                    # 关键修复：确保本地配置包含最新的默认模板
                    # 如果saved里的prompts为空，或者缺少核心key，则合并
                    if "prompts" not in saved:
                        saved["prompts"] = DEFAULT_PROMPTS.copy()
                    else:
                        # 强行补充缺失的新模板
                        for k, v in DEFAULT_PROMPTS.items():
                            if k not in saved["prompts"]:
                                saved["prompts"][k] = v

                    return {**default, **saved}
            except:
                pass
        return default

    @staticmethod
    def save(config):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)


# ================= 4. Word 报告生成器 (专业排版) =================

class WordReportGenerator:
    @staticmethod
    def set_font(run, font_name_cn='宋体', font_name_en='Times New Roman', size=10.5, bold=False, color=None):
        run.font.name = font_name_en
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name_cn)
        run.font.size = Pt(size)
        run.font.bold = bold
        if color: run.font.color.rgb = color

    @staticmethod
    def set_cell_shading(cell, hex_color):
        shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), hex_color))
        cell._tc.get_or_add_tcPr().append(shading_elm)

    @staticmethod
    def generate(tasks, save_path):
        # [核心修改] 尝试加载模板
        template_name = "模板.docx"
        if os.path.exists(template_name):
            doc = Document(template_name)
            # 如果有模板，我们通常跳过页面边距设置，沿用模板的设置
            print(f"已加载模板: {template_name}")

            # 可以在模板末尾加个分页符，防止内容紧贴封面
            doc.add_page_break()
        else:
            # 没模板，创建新文档并设置边距
            doc = Document()
            section = doc.sections[0]
            section.top_margin = Cm(2.0)
            section.bottom_margin = Cm(2.0)
            section.left_margin = Cm(2.0)
            section.right_margin = Cm(2.0)

            # 手动添加简易标题（因为没模板）
            title_para = doc.add_paragraph()
            title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = title_para.add_run("智能隐患排查报告")
            WordReportGenerator.set_font(run, size=18, bold=True)
            doc.add_paragraph()

        # --- 概况信息 (追加到文档中) ---
        # 如果你希望封面由模板决定，可以注释掉下面这段概况表代码
        # 或者保留它作为正文第一部分
        info_para = doc.add_paragraph()
        info_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run_time = info_para.add_run(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')} | 点位：{len(tasks)}个")
        WordReportGenerator.set_font(run_time, size=9, color=RGBColor(100, 100, 100))

        doc.add_paragraph()  # 空行

        # --- 循环生成具体内容 (逻辑不变) ---
        for idx, task in enumerate(tasks, 1):
            # 1. 点位标题条
            title_table = doc.add_table(rows=1, cols=1)
            title_table.style = 'Table Grid'
            title_cell = title_table.cell(0, 0)
            WordReportGenerator.set_cell_shading(title_cell, "F2F2F2")

            p = title_cell.paragraphs[0]
            run = p.add_run(f"NO.{idx}  点位名称：{task['name']}")
            WordReportGenerator.set_font(run, size=12, bold=True)
            doc.add_paragraph()

            # 2. 图片
            img_para = doc.add_paragraph()
            img_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if os.path.exists(task['path']):
                try:
                    doc.add_picture(task['path'], height=Cm(6.5))
                except:
                    run = img_para.add_run("[图片损坏]")
                    WordReportGenerator.set_font(run, color=RGBColor(255, 0, 0))
            else:
                img_para.add_run("[图片路径不存在]")
            doc.add_paragraph()

            # 3. 表格
            data = task.get('data', [])
            headers = ["风险/指数等级", "详细描述", "依据标准/常识", "整改或优化建议"]
            widths = [Cm(2.5), Cm(6.0), Cm(3.8), Cm(4.5)]

            table = doc.add_table(rows=1, cols=4)
            table.style = 'Table Grid'
            table.autofit = False
            table.alignment = WD_TABLE_ALIGNMENT.CENTER

            # 表头
            hdr_cells = table.rows[0].cells
            for i, text in enumerate(headers):
                p = hdr_cells[i].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(text)
                WordReportGenerator.set_font(run, bold=True, size=10.5)
                WordReportGenerator.set_cell_shading(hdr_cells[i], "E7E6E6")
                hdr_cells[i].width = widths[i]

            # 内容填充
            if not data or isinstance(data, str) or len(data) == 0:
                row = table.add_row()
                cell = row.cells[0]
                cell.merge(row.cells[3])
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run("AI 未发现明显隐患或改进项。")
                WordReportGenerator.set_font(run, color=RGBColor(0, 128, 0))
            else:
                for item in data:
                    row_cells = table.add_row().cells

                    level = item.get("risk_level", "一般")
                    cell_risk = row_cells[0]
                    p_risk = cell_risk.paragraphs[0]
                    p_risk.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run_risk = p_risk.add_run(level)
                    WordReportGenerator.set_font(run_risk, bold=True, size=10.5)

                    # 智能配色
                    if any(x in level for x in ["重大", "严重", "High", "警示"]):
                        WordReportGenerator.set_cell_shading(cell_risk, "FF0000")
                        run_risk.font.color.rgb = RGBColor(255, 255, 255)
                    elif any(x in level for x in ["较大", "需整理", "需改善", "Medium"]):
                        WordReportGenerator.set_cell_shading(cell_risk, "FFC000")
                        run_risk.font.color.rgb = RGBColor(255, 255, 255)
                    else:
                        run_risk.font.color.rgb = RGBColor(0, 0, 0)

                    contents = [item.get("issue", ""), item.get("regulation", ""), item.get("correction", "")]
                    for j, txt in enumerate(contents):
                        cell = row_cells[j + 1]
                        p = cell.paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        run = p.add_run(txt)
                        WordReportGenerator.set_font(run, size=10)
                        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

                    for k in range(4): row_cells[k].width = widths[k]

            if idx < len(tasks): doc.add_page_break()

        doc.save(save_path)


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

            p_conf = PROVIDER_PRESETS.get(p_name, {})
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
                        {"type": "text", "text": "请按要求分析风险"}
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

        # 颜色定义
        colors = {"红": "#FFE5E5", "橙": "#FFF4E5", "蓝": "#E3F2FD"}
        borders = {"红": "#FF0000", "橙": "#FF8800", "蓝": "#2196F3"}

        # 智能匹配逻辑 (兼容工程LEC标准 和 生活居家标准)
        # 红色：重大风险、严重违规、卫生警示、High
        if any(x in level for x in ["重大", "严重", "High", "警示"]):
            bg, bd = colors["红"], borders["红"]
        # 橙色：较大风险、需整理、需改善、Medium
        elif any(x in level for x in ["较大", "需整理", "需改善", "Medium"]):
            bg, bd = colors["橙"], borders["橙"]
        # 蓝色：一般风险、生活建议、Low
        else:
            bg, bd = colors["蓝"], borders["蓝"]

        self.setStyleSheet(f"RiskCard {{ background-color: {bg}; border-left: 5px solid {bd}; border-radius: 4px; margin-bottom: 5px; }}")

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
        self.tasks = []
        self.queue_workers = []
        self.current_task_id = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("普洱版纳区域AI智能终端 ")
        self.resize(1300, 850)

        # --- 工具栏 ---
        toolbar = QToolBar("Main")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        toolbar.addWidget(QLabel("  场景模式: "))
        self.cbo_prompt = QComboBox()
        # 从配置中加载 Prompt 列表
        self.cbo_prompt.addItems(self.config.get("prompts", DEFAULT_PROMPTS).keys())
        # 设置默认选中项
        self.cbo_prompt.setCurrentText(self.config.get("last_prompt", list(DEFAULT_PROMPTS.keys())[0]))
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

        # [功能] 清空队列
        btn_clear = QAction("🗑️ 清空队列", self)
        btn_clear.triggered.connect(self.clear_queue)
        toolbar.addAction(btn_clear)

        btn_export = QAction(QIcon(), "📄 导出Word报告", self)
        btn_export.triggered.connect(self.export_word)
        toolbar.addAction(btn_export)

        # 弹簧
        empty = QWidget()
        empty.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        toolbar.addWidget(empty)

        btn_setting = QAction("⚙ 设置", self)
        btn_setting.triggered.connect(self.open_settings)
        toolbar.addAction(btn_setting)

        # --- 主布局 ---
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # 左侧列表
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(5,5,5,5)
        self.lbl_count = QLabel("待审队列 (0/20)")
        left_layout.addWidget(self.lbl_count)
        self.list_widget = QListWidget()
        self.list_widget.itemClicked.connect(self.on_item_clicked)
        left_layout.addWidget(self.list_widget)

        # 右侧内容
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)

        # 图片容器
        self.lbl_image = QLabel("请从左侧选择图片")
        self.lbl_image.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_image.setStyleSheet("background-color: #333; color: #AAA; border-radius: 6px;")
        self.lbl_image.setMinimumHeight(400)
        right_layout.addWidget(self.lbl_image, 1)

        # 结果容器
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
        self.setCentralWidget(splitter)

        # 状态栏
        self.status_bar = self.statusBar()
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedWidth(200)
        self.status_bar.addPermanentWidget(self.progress_bar)

    # --- 逻辑 ---

    def save_prompt_selection(self, text):
        # 防止清空列表时触发保存导致配置变空
        if not text: return
        self.config["last_prompt"] = text
        ConfigManager.save(self.config)

    def add_files(self):
        current_count = len(self.tasks)
        if current_count >= 20:
            QMessageBox.warning(self, "数量限制", "为保证运行稳定，单次排查请控制在 20 张图片以内。\n建议先清空队列。")
            return

        remaining = 20 - current_count
        paths, _ = QFileDialog.getOpenFileNames(self, f"选择图片 (还能选 {remaining} 张)", "", "Images (*.jpg *.png *.jpeg)")

        if not paths: return

        if len(paths) > remaining:
            QMessageBox.warning(self, "超限提示", f"你选择了 {len(paths)} 张，自动截取前 {remaining} 张。")
            paths = paths[:remaining]

        for path in paths:
            if any(t['path'] == path for t in self.tasks): continue

            task_id = str(time.time()) + os.path.basename(path)
            self.tasks.append({"id": task_id, "path": path, "name": os.path.basename(path), "status": "waiting", "data": None})

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

        # 获取当前选中的提示词内容 (从Config中读取，确保自定义生效)
        selected_template_name = self.cbo_prompt.currentText()
        prompts_dict = self.config.get("prompts", DEFAULT_PROMPTS)
        # 获取内容，如果找不到则回退到默认
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
            self.result_layout.addWidget(QLabel("🚀 正在智能分析中 (LEC/健康双模)，请稍候..."))
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
        scaled = pix.scaled(self.lbl_image.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        self.lbl_image.setPixmap(scaled)
        self.render_result(task)

    def update_list_color(self, task_id, color):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == task_id:
                item.setForeground(QColor(color))

    def export_word(self):
        if not self.tasks: return

        # [修改] 精确到秒的文件名，防止覆盖
        current_time_str = datetime.now().strftime('%Y%m%d_%H%M%S')
        default_name = f"智能排查报告_{current_time_str}.docx"

        path, _ = QFileDialog.getSaveFileName(self, "保存报告", default_name, "Word Files (*.docx)")
        if not path: return

        try:
            WordReportGenerator.generate(self.tasks, path)
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

        cbo_provider = QComboBox()
        cbo_provider.addItems(PROVIDER_PRESETS.keys())
        # 防止配置文件里的厂商在新版不存在导致报错
        curr_prov = self.config.get("current_provider")
        if curr_prov not in PROVIDER_PRESETS: curr_prov = list(PROVIDER_PRESETS.keys())[0]
        cbo_provider.setCurrentText(curr_prov)

        txt_base_url = QLineEdit()
        txt_model = QLineEdit()
        txt_key = QLineEdit(self.config.get("api_key"))
        txt_key.setEchoMode(QLineEdit.EchoMode.Password)

        def on_provider_change(text):
            preset = PROVIDER_PRESETS.get(text, {})
            if text == "自定义 (Custom)":
                custom_saved = self.config.get("custom_provider_settings", {})
                txt_base_url.setText(custom_saved.get("base_url", ""))
                txt_model.setText(custom_saved.get("model", ""))
                txt_base_url.setPlaceholderText("例如: https://api.xxx.com/v1")
                txt_model.setPlaceholderText("例如: llama3-70b")
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

        # --- Tab 2: 提示词模板 ---
        tab_prompt = QWidget()
        layout_prompt = QVBoxLayout(tab_prompt)

        # 深度拷贝，避免直接修改 config
        local_prompts = self.config.get("prompts", DEFAULT_PROMPTS).copy()

        cbo_template = QComboBox()
        cbo_template.addItems(local_prompts.keys())

        txt_prompt_edit = QTextEdit()

        # 记录上一次选中的模板名
        self._temp_last_selected = cbo_template.currentText()

        def load_template_to_editor(template_name):
            content = local_prompts.get(template_name, "")
            txt_prompt_edit.setText(content)
            self._temp_last_selected = template_name

        def save_editor_to_memory():
            current_text = txt_prompt_edit.toPlainText()
            if self._temp_last_selected and self._temp_last_selected in local_prompts:
                local_prompts[self._temp_last_selected] = current_text

        def on_template_change(new_name):
            save_editor_to_memory()
            load_template_to_editor(new_name)

        cbo_template.currentTextChanged.connect(on_template_change)

        # 初始化显示
        if self._temp_last_selected:
            load_template_to_editor(self._temp_last_selected)

        layout_prompt.addWidget(QLabel("选择模板进行编辑:"))
        layout_prompt.addWidget(cbo_template)
        layout_prompt.addWidget(txt_prompt_edit)
        layout_prompt.addWidget(QLabel("<small style='color:grey'>* 切换模板或点击保存时，修改会自动生效</small>"))

        tabs.addTab(tab_prompt, "📝 提示词编辑")

        # --- 按钮区域 ---
        btn_box = QHBoxLayout()
        btn_save = QPushButton("保存所有配置")
        btn_save.setMinimumHeight(40)
        btn_save.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; border-radius: 4px;")
        btn_cancel = QPushButton("取消")

        def save_all():
            try:
                # 1. 保存当前编辑
                save_editor_to_memory()

                # 2. 更新配置
                self.config["current_provider"] = cbo_provider.currentText()
                self.config["api_key"] = txt_key.text().strip()
                self.config["prompts"] = local_prompts

                if cbo_provider.currentText() == "自定义 (Custom)":
                    self.config["custom_provider_settings"] = {
                        "base_url": txt_base_url.text().strip(),
                        "model": txt_model.text().strip()
                    }

                # 3. 写入文件
                ConfigManager.save(self.config)

                # 4. [关键] 安全刷新主界面下拉框
                self.cbo_prompt.blockSignals(True) # 暂停信号

                current_main_selection = self.cbo_prompt.currentText()
                self.cbo_prompt.clear()
                self.cbo_prompt.addItems(self.config["prompts"].keys())

                if current_main_selection in self.config["prompts"]:
                    self.cbo_prompt.setCurrentText(current_main_selection)
                else:
                    self.cbo_prompt.setCurrentIndex(0)
                    self.config["last_prompt"] = self.cbo_prompt.currentText()
                    ConfigManager.save(self.config)

                self.cbo_prompt.blockSignals(False) # 恢复信号

                dlg.accept()
                self.status_bar.showMessage("✅ 配置已保存")

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
