#!/usr/bin/env python3
# -*- coding: utf-8 -*-


import sys, os, json, base64, time, re, traceback, math
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple
from collections import defaultdict, Counter

os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"
os.environ["QT_API"] = "pyqt6"

import httpx
from openai import OpenAI
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from PyQt6.QtCore import (Qt, QThread, pyqtSignal, QTimer, QPointF, QRectF,
                          QBuffer, QByteArray, QIODevice, QSize, QPropertyAnimation, QEasingCurve, pyqtProperty)
from PyQt6.QtGui import (QPixmap, QColor, QAction, QPainter, QPen, QFont,
                         QImage, QPainterPath, QBrush, QKeySequence, QPalette, QLinearGradient, QIcon)
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QLabel, QListWidget, QListWidgetItem, QSplitter,
                             QScrollArea, QFrame, QFileDialog, QProgressBar, QMessageBox, QDialog,
                             QFormLayout, QLineEdit, QComboBox, QToolBar, QSizePolicy, QTabWidget,
                             QTextEdit, QGroupBox, QGridLayout, QSpinBox, QPlainTextEdit, QDialogButtonBox,
                             QToolButton, QMenu, QInputDialog, QGraphicsView, QGraphicsScene,
                             QGraphicsPixmapItem, QGraphicsRectItem, QGraphicsEllipseItem, QGraphicsPathItem,
                             QGraphicsTextItem, QGraphicsItem, QCheckBox, QRadioButton, QButtonGroup,
                             QSlider, QGraphicsDropShadowEffect)

# ==================== 全局配置 ====================
CONFIG_FILE = "app_config_v4.json"
HISTORY_FILE = "inspection_history_v4.json"
STATS_FILE = "inspection_stats_v4.json"
TEMPLATE_DIR = "templates"
MAX_IMAGES = 50  # 提升到50张
EXPORT_IMG_DIR = "_export_marked"

# ==================== 全局配置 ====================
# ... (其他配置保持不变)

# 主题色配置 (修复版：补充了背景色键值)
THEME_COLORS = {
    "primary": "#1976D2",  # 主色调（蓝色）
    "secondary": "#424242",  # 次要色（深灰）
    "success": "#4CAF50",  # 成功（绿色）
    "warning": "#FF9800",  # 警告（橙色）
    "danger": "#F44336",  # 危险（红色）
    "info": "#2196F3",  # 信息（浅蓝）
    "light": "#F5F5F5",  # 浅色背景
    "dark": "#212121",  # 深色背景

    # 核心等级颜色
    "severe_safety": "#D32F2F",  # 严重安全 (红)
    "general_safety": "#F57C00",  # 一般安全 (橙)
    "severe_quality": "#E64A19",  # 严重质量 (深橙)
    "general_quality": "#FFA726",  # 一般质量 (黄)

    # 缺失的背景色 (这是导致报错的原因)
    "severe_safety_bg": "#FFEBEE",  # 浅红背景
    "general_safety_bg": "#FFF3E0",  # 浅橙背景
    "severe_quality_bg": "#FBE9E7",  # 浅深橙背景
    "general_quality_bg": "#FFF8E1",  # 浅黄背景
}

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
        "安全生产专项检查", "工程质量专项检查", "安全质量综合检查",
        "节前安全检查", "复工安全质量检查", "专项整治检查"
    ],
    "project_overview_map": {
        "勐海农村供水保障项目": "本工程位于西双版纳州勐海县，主要建设内容包括新建取水坝、输水管网及配套水厂设施，旨在解决周边5个乡镇的农村饮水安全问题，设计供水规模为2.5万吨/日。",
        "城乡供水一体化项目": "勐海县城乡供水一体化建设项目涉及勐海县城、勐遮镇、勐混镇、勐阿镇、打洛镇、勐满镇、格朗和乡、勐宋乡8个片区，覆盖现状人口28.53万人。主要建设内容为：新建3座水厂，总建设规模32000m³/d，其中县城三水厂20000m³/d，格朗和乡4000m³/d，勐混镇 8000m³/d。扩建水厂1座，勐遮镇扩容建设5000m³/d 工艺设施，扩容后总处理规模15000m³/d。利用存量水厂7座，现状总供水规模61500m³/d，其中县城一水厂10000m³/d，县城二水厂 20000m³/d，勐遮水厂10000m³/d，打洛镇曼彦水厂7500m³/d，勐阿水厂6000m³/d，勐满水厂4000m³/d，勐宋水厂4000m³/d。建设DN100-DN900输配水管网376.87km，配套建设信息化设施、阀门井、排泥阀、闸阀、入户管及其他附属设施。",
        "勐阿水库建设项目": "勐海县勐阿水库项目总投资7.645亿元。勐海县勐阿水库规模为中型，由枢纽工程、输(供)水工程和水厂工程组成，枢纽工程由大坝、溢洪道和输水(兼导流)隧洞组成。合同工期48个月。",
        "热水河水库建设项目": "江城县热水河水库工程主要由枢纽工程和输水工程两部分组成。枢纽工程主要包括拦河坝、溢洪道、导流输水隧洞等，建成后将有效缓解江城县城供水压力。江城热水河水库项目总投资5.61亿元。江城县热水河水库工程由枢纽工程和输水工程组成，合同工期为48个月。",
        "三道箐水库建设项目": "澜沧县三道箐水库位于澜沧县中北部的东河乡拉巴河上游的三道箐河上，水库工程由枢纽工程及灌区工程组成。枢纽工程主要由大坝、1～2#副坝、溢洪道、输水导流兼放空隧洞及主坝～1#副坝库岸防渗组成。水库为小（1）型水库，总库容406万m3，澜沧县三道箐水库项目总投资2.32808亿元，合同工期24个月。"

    }
}

# ================= 3. V5.0 超级规范知识库（融合V3.5精华） =================

REGULATION_DATABASE_V5 = {
    "管道": {
        "role_desc": "管道与阀门工艺专家",
        "norms": """
### GB 50242-2002《建筑给水排水及采暖工程施工质量验收规范》
**第3.3.13条** 法兰连接螺栓：紧固后露出螺母2-3扣，垫片不突入管内。
**第3.3.15条** 阀门安装：型号/规格/耐压/方向正确，手轮便于操作，止回阀低进高出。
**第3.3.16条** 橡胶软接头：管道处于正负压状态时，应设防拉脱限位装置。
### GB 50268-2008《给水排水管道工程施工及验收规范》
**第5.2.6条** 柔性接口：刚性管道与柔性管道连接处应设柔性连接管。
""",
        "checklist": [
            "法兰螺栓：是否露牙2-3扣？是否对称紧固？是否有双垫片/偏垫？",
            "软接头：是否安装限位螺栓？长度是否被拉伸/压缩过度？",
            "阀门：止回阀方向是否装反？手轮是否朝下（违规）？",
            "支吊架：固定支架与滑动支架是否混用？U型卡是否未锁紧？",
            "异径管：水平管变径是否用偏心（顶平/底平）？焊接有无咬边？"
        ],
        "anti_hallucination": "临时封堵不是缺阀门；试压支撑不是支架不足；保温未施工可能分阶段。"
    },

    "电气": {
        "role_desc": "注册电气工程师",
        "norms": """
### GB 50303-2015《建筑电气工程施工质量验收规范》
**第12.1.1条** 金属桥架及其支架全长应不少于2处与接地干线相连；非镀锌桥架连接板两端跨接铜芯接地线≥4mm²。
**第14.1.1条** 箱内PE线应通过汇流排连接，严禁串联连接。
**第18.2.1条** 电缆弯曲半径：单芯≥20D，多芯≥15D。
### GB 50169-2016
**第3.3.2条** 接地线应为黄绿双色标识，严禁作负载线。
""",
        "checklist": [
            "接地连续性：桥架跨接线是否缺失？配电箱门是否软铜线接地？",
            "线色标准：PE线必须是黄绿双色。相序色标是否混乱？",
            "电缆工艺：弯曲半径是否过小（折死弯）？进箱体是否无护口保护？",
            "配电箱：是否一管一孔？是否存在多股线未压接端子？"
        ],
        "anti_hallucination": "施工中电缆凌乱待整理；旧规范PE线颜色可能不同（但新规范必须黄绿）。"
    },

    "结构": {
        "role_desc": "结构总工程师",
        "norms": """
### GB 50204-2015《混凝土结构工程施工质量验收规范》
**第5.5.1条** 钢筋保护层厚度偏差：梁柱±5mm，墙板±3mm。
**第8.3.2条** 施工缝位置：应留在结构受剪力较小且便于施工的部位。
**第8.4.1条** 拆模强度：悬臂构件≥100%，板(≤2m)≥50%。
### JGJ 130-2011
**第6.3.6条** 立杆步距≤1.8m，扫地杆距底座≤200mm。
""",
        "checklist": [
            "钢筋工程：绑扎间距是否均匀？马凳筋/垫块是否缺失？接头错开率？",
            "模板支撑：立杆间距、扫地杆高度、剪刀撑设置是否合规？",
            "混凝土外观：是否有蜂窝、麻面、孔洞、露筋？施工缝处理是否凿毛？",
            "对拉螺栓：是否未切割？是否未做防锈处理？"
        ],
        "anti_hallucination": "未抹面不是不平整；待绑扎区域钢筋散乱是正常的；温度裂缝<0.3mm通常允许。"
    },

    "机械": {
        "role_desc": "起重机械专家",
        "norms": """
### GB 5144-2006《塔式起重机安全规程》
**第6.1.1条** 力矩限制器、起重量限制器、高度/幅度/回转限位器应灵敏可靠。
**第7.2.1条** 钢丝绳：断丝数在一个节距内超过10%应报废。
### JGJ 33-2012
**第5.2.1条** 吊篮安全锁必须在标定期限内。
""",
        "checklist": [
            "塔吊：限位器是否失效？标准节螺栓是否松动？",
            "钢丝绳：是否有断丝、断股、死弯、严重锈蚀？",
            "吊篮：安全锁是否失效？配重块是否固定？钢丝绳是否垂直？",
            "吊钩：防脱棘爪是否缺失？是否有补焊痕迹（严禁补焊）？"
        ],
        "anti_hallucination": "停工状态吊钩无荷载是正常的；钢丝绳表面油污可能是润滑脂。"
    },

    "基坑": {
        "role_desc": "岩土工程师",
        "norms": """
### JGJ 120-2012《建筑基坑支护技术规程》
**第8.1.1条** 基坑周边施工材料堆放距离坑边不应小于2m。
### GB 50497-2009
**第4.2.1条** 监测报警值：水平位移累计值达到设计限值70%-80%。
""",
        "checklist": [
            "临边堆载：坑边2m内是否有重物/机械？",
            "边坡防护：喷锚是否脱落？是否有裂缝？排水沟是否畅通？",
            "降水设施：抽水泵是否工作？水位是否控制在基底以下？",
            "临边防护：是否有1.2m高防护栏杆和密目网？"
        ],
        "anti_hallucination": "雨后积水需及时抽排，但不代表降水失效；土方开挖阶段临时边坡未支护是过程态。"
    },

    "消防": {
        "role_desc": "注册消防工程师",
        "norms": """
### GB 50720-2011《建设工程施工现场消防安全技术规范》
**第5.3.7条** 氧气瓶与乙炔瓶工作间距≥5m，距离明火作业点≥10m。
**第6.2.1条** 临时消防设施应与主体结构施工同步设置。
### GB 50016-2014
**第8.1.2条** 灭火器配置：每处不少于2具，且保护距离符合要求。
""",
        "checklist": [
            "动火作业：是否有接火盆？是否有监护人？是否有灭火器？",
            "气瓶管理：是否直立固定？是否有防震圈？氧乙间距是否够？",
            "灭火器：压力指针是否在绿区？是否过期？",
            "消防通道：是否被材料堵塞？宽度是否<4m？"
        ],
        "anti_hallucination": "空瓶横放等待清运是可以的；监护人可能在画面外。"
    },

    "安全": {
        "role_desc": "注册安全工程师",
        "norms": """
### JGJ 59-2011《建筑施工安全检查标准》
**第3.2.5条** 进入施工现场必须正确佩戴安全帽，系好下颌带。
**第5.1.1条** 高处作业（≥2m）必须系安全带，挂点必须牢固，严禁低挂高用。
### JGJ 46-2005
**第5.1.1条** 施工现场临时用电必须采用TN-S系统（三级配电两级保护）。
""",
        "checklist": [
            "三宝：安全帽（带子）、安全带（高挂）、安全网（完整）。",
            "四口五临边：楼梯口/电梯井口/通道口/预留洞口防护是否缺失？",
            "脚手架：作业层是否有脚手板？是否有挡脚板？",
            "违章作业：是否有人坐在栏杆上？是否酒后作业？"
        ],
        "anti_hallucination": "管理人员在安全通道内检查可短时摘帽；2m以下作业无需安全带。"
    },

    "暖通": {
        "role_desc": "暖通工程师",
        "norms": """
### GB 50243-2016《通风与空调工程施工质量验收规范》
**第4.2.1条** 风管法兰垫片材质符合要求，不凸入管内，不突出法兰外。
**第4.2.7条** 风管支吊架间距：水平管直径>400mm间距≤3m。
**第7.3.1条** 防腐与绝热施工前，管道/风管表面应平整、无油脂锈蚀。
""",
        "checklist": [
            "风管安装：法兰连接是否严密？支架间距是否过大？",
            "保温工程：保温层是否开裂、脱落？厚度是否达标？",
            "软管连接：长度是否>300mm（违规）？是否扭曲变形？",
            "管道坡度：冷凝水管坡度是否足够？"
        ],
        "anti_hallucination": "测试用临时管线；保温未施工完毕。"
    },

    "给排水": {
        "role_desc": "给排水工程师",
        "norms": """
### GB 50268-2008
**第5.3.1条** 管道基础砂垫层厚度不应小于100mm。
### GB 50015-2019
**第3.3.5条** 排水管道坡度：DN50标准坡度25‰，DN75标准坡度15‰。
**第3.5.8条** 地漏水封深度不得小于50mm。
""",
        "checklist": [
            "管道敷设：排水管坡度是否倒坡？支墩设置是否合理？",
            "地漏安装：水封高度是否达标？是否便于清通？",
            "管沟回填：是否分层夯实？是否含有大石块？",
            "闭水试验：是否按规定蓄水？有无渗漏？"
        ],
        "anti_hallucination": "临时排水管可降低标准；开挖未回填正在施工中。"
    },

    "防水": {
        "role_desc": "防水工程师",
        "norms": """
### GB 50207-2012《屋面工程质量验收规范》
**第4.3.1条** 卷材搭接宽度：短边≥150mm，长边≥100mm。
**第5.2.1条** 涂膜防水层厚度应符合设计要求，无裂纹、皱折、流淌、鼓泡。
### GB 50108-2008
**第4.1.14条** 地下工程施工缝应设置止水带/止水钢板。
""",
        "checklist": [
            "卷材施工：搭接宽度是否不足？收头是否密封？是否有空鼓？",
            "涂膜防水：涂刷是否均匀？是否存在露底？",
            "细部构造：阴阳角是否做圆弧处理？管根部是否加强？",
            "止水带：埋设位置是否居中？固定是否牢固？"
        ],
        "anti_hallucination": "防水层未做保护层前不能上人；养护期积水是正常的。"
    },

    "环保": {
        "role_desc": "环境工程师",
        "norms": """
### GB 12523-2011《建筑施工场界环境噪声排放标准》
**第4.1条** 噪声限值：昼间70dB，夜间55dB。
### GB 50325-2020
**第3.1.2条** 施工现场必须实施封闭管理，围挡高度≥2.5m（市区）。
**第6.1.1条** 裸露土方应采取覆盖、绿化或固化措施。
""",
        "checklist": [
            "扬尘控制：裸土是否覆盖？是否有洒水降尘措施？车辆冲洗？",
            "噪声控制：是否夜间强噪声施工？",
            "废水处理：是否有三级沉淀池？是否直排污水？",
            "固废管理：建筑垃圾是否分类堆放？是否有危废标识？"
        ],
        "anti_hallucination": "短时扬尘（如正在倒土）需结合雾炮使用判断；雨天无需洒水。"
    }
}
ROUTER_SYSTEM_PROMPT = """
你是一名拥有25年经验的工程建设总监。请扫描施工现场图片，快速识别核心施工内容，指派 **3-4 名** 最对口的硬核技术专家。

### 必须从以下 10 个角色中选择（严禁编造其他角色）：
1. **管道** (涉及阀门/法兰/水泵/管道工艺)
2. **电气** (涉及配电箱/电缆/桥架/接地/防雷)
3. **结构** (涉及钢筋/模板/混凝土/脚手架/螺栓)
4. **机械** (涉及塔吊/施工电梯/吊篮/钢丝绳)
5. **基坑** (涉及边坡/支护/土方/降水/监测)
6. **消防** (涉及动火作业/气瓶/灭火器)
7. **暖通** (涉及风管/空调水管/保温/风口)
8. **给排水** (涉及市政管网/检查井/排水沟)
9. **防水** (涉及卷材/涂膜/止水带)
10. **安全** (兜底角色，涉及人员行为/临边防护)

### 指派逻辑
- 看到 **法兰/阀门/软接头** -> 必派 **管道**
- 看到 **电箱/电线/桥架** -> 必派 **电气**
- 看到 **钢筋/浇筑/支模** -> 必派 **结构**
- 看到 **塔吊/吊篮** -> 必派 **机械**
- 看到 **深坑/护坡** -> 必派 **基坑**
- 看到 **电焊/气瓶** -> 必派 **消防**
- 看到 **风管/保温** -> 必派 **暖通**


**强制规则**：
1. 始终包含 **安全** 专家。
2. 如果画面模糊或无特定专业内容，仅输出 ["安全"]。
3. 输出必须是 JSON 字符串列表。

示例：`["管道", "电气", "安全"]`
"""
# ==================== V4.0 聚焦安全质量的提示词 ====================
ROUTER_PROMPT_V4 = """你是施工总监（25年经验），专注识别【安全隐患】和【质量问题】。

## 核心任务
快速识别现场风险，指派2-4名专家团队。

## 场景识别
判断施工类型：
- 🏗️ 基础：基坑/桩基/地下室
- 🧱 主体：钢筋/模板/混凝土
- 🔧 安装：管道/电气/机械
- ⚠️ 危险作业：高处/动火/受限空间

## 关键要素（聚焦安全质量）
扫描是否存在：

**安全隐患要素**：
- 人员：未戴安全帽、未系安全带、违章操作
- 临边：无防护栏杆、洞口未封闭
- 机械：限位器失效、钢丝绳断丝、防坠器过期
- 用电：配电箱裸露、电缆破损、无漏保
- 基坑：边坡裂缝、临边无防护、超载堆放
- 消防：动火无监护、气瓶间距不足

**质量问题要素**：
- 管道：法兰螺栓、软接头、阀门、焊缝
- 电气：接地线、PE线色、电缆半径
- 结构：钢筋间距、保护层、混凝土蜂窝
- 机械：钢丝绳规格、吊篮配重
- 基坑：喷锚网厚度、锚杆拉拔

## 专家指派规则
| 画面内容 | 必派专家 | 原因 |
|---------|---------|------|
| 法兰/管道/阀门 | 管道 + 安全 | 质量+安全 |
| 配电箱/电缆/接地 | 电气 + 安全 | 触电风险 |
| 钢筋/模板/混凝土 | 结构 + 安全 | 坍塌风险 |
| 塔吊/电梯/吊篮 | 机械 + 安全 | 坠落风险 |
| 基坑/边坡 | 基坑 + 安全 | 坍塌风险 |
| 动火/气瓶 | 消防 + 安全 | 火灾风险 |
| 仅人员行为 | 安全 | 行为风险 |

**强制规则**：
1. 安全专家永远在列（生命第一）
2. 有专科设备必派对应专家
3. 最多4名专家（聚焦核心）

输出JSON数组：["管道", "安全"]
禁止：解释文字、超4人"""

SPECIALIST_PROMPT_TEMPLATE = """
你现在是一名【{role}】（{role_desc}），拥有30年一线经验。
请对图片进行**工艺级找茬**。不要讲大道理，只找具体的**技术通病**和**违规细节**。

### 1. 核心规范依据 (必须引用)
{norms}

### 2. 你的深度检查清单 (Checklist)
请重点扫描以下细节：
{checklist}

### 3. 互斥协议
- 如果你是【安全】，只管人的不安全行为（未戴帽/带）和临边洞口防护，不要管具体的设备工艺。
- 如果你是【专科专家】(如管道/电气)，只管你专业内的**技术参数**、**安装工艺**和**实体质量**，不要管通用的安全问题。

### 4. 误判警示 (Anti-Hallucination)
{anti_hallucination}

### 5. 输出格式严格要求 (JSON)
你必须输出一个 JSON 数组，包含以下字段：

- **risk_level**: 必须严格从以下 4 个选项中选择一个（根据严重程度和类别）：
  - "严重安全隐患" (可能导致伤亡)
  - "一般安全隐患" (一般违章)
  - "严重质量缺陷" (影响结构安全或主要功能，如法兰漏水、钢筋少放)
  - "一般质量缺陷" (影响观感或一般功能，如螺栓生锈、标签缺失)

- **issue**: 【{role}】+ 具体描述。必须包含：部位 + 问题 + (标准值 vs 实际值) + 后果。
  - ✅ 好例子："DN150法兰连接螺栓仅露出1扣（规范要求2-3扣），存在松动泄漏风险"
  - ❌ 坏例子："法兰安装不规范"

- **regulation**: 必须引用上述规范中的具体条文号。例："GB 50242-2002第3.3.13条"

- **correction**: 分步骤的整改措施。

- **bbox**: [x1, y1, x2, y2] (尽可能定位问题主体，无则 null)



**JSON 示例**:
[
  {{
    "risk_level": "严重质量缺陷",
    "issue": "【{role}】DN100止回阀安装方向错误（箭头向下），违反低进高出原则，导致水泵无法正常出水",
    "regulation": "GB 50242-2002第3.3.15条",
    "correction": "拆除重装，调整阀门方向与水流方向一致",
    "bbox": [100, 200, 300, 400],
    "confidence": 0.98
  }}
]
"""

DEFAULT_PROMPTS_V4 = {
    "V4.0 安全质量双聚焦": "聚焦安全隐患+质量问题",
    "安全隐患专项": "仅识别安全隐患（忽略质量）",
    "质量问题专项": "仅识别质量问题（忽略安全）",
    "高危风险筛查": "仅识别严重安全隐患",
}

DEFAULT_PROVIDERS = {
    "阿里百炼(Qwen-VL-Max)": {
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "model": "qwen-vl-max"
    },
    "阿里百炼(Qwen2.5-VL)": {
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "model": "qwen2.5-vl-72b"
    },
    "硅基流动(Qwen2-VL)": {
        "base_url": "https://api.siliconflow.cn/v1",
        "model": "Qwen/Qwen2-VL-72B-Instruct"
    },
    "OpenAI(GPT-4o)": {
        "base_url": "https://api.openai.com/v1",
        "model": "gpt-4o"
    },
    "自定义": {
        "base_url": "",
        "model": ""
    }
}


# ==================== 配置管理 ====================
class ConfigManager:
    @staticmethod
    def get_default():
        return {
            "current_provider": "阿里百炼(Qwen-VL-Max)",
            "api_key": "",
            "last_prompt": "V4.0 安全质量双聚焦",
            "custom_provider_settings": {"base_url": "", "model": ""},
            "business_data": DEFAULT_BUSINESS_DATA,
            "prompts": DEFAULT_PROMPTS_V4,
            "provider_presets": DEFAULT_PROVIDERS,
            "max_concurrency": 2,
            "max_retries": 2,
            "request_timeout_sec": 90,
            "temperature": 0.2,
            "last_check_person": "",
            "recent_check_areas": [],
            "enable_dual_model": False,
            "secondary_provider": "OpenAI(GPT-4o)",
            "secondary_api_key": "",
            "auto_save_history": True,
            "show_confidence": True,
            "theme": "light"
        }

    @staticmethod
    def load():
        default = ConfigManager.get_default()
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                for k, v in default.items():
                    if k not in saved:
                        saved[k] = v
                if "business_data" in saved:
                    for key in default["business_data"]:
                        if key not in saved["business_data"]:
                            saved["business_data"][key] = default["business_data"][key]
                return saved
            except Exception as e:
                print(f"配置加载失败: {e}")
        ConfigManager.save(default)
        return default

    @staticmethod
    def save(config):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"配置保存失败: {e}")


# ==================== 历史记录管理 ====================
class HistoryManager:
    @staticmethod
    def load():
        if os.path.exists(HISTORY_FILE):
            try:
                with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                return {"inspections": []}
        return {"inspections": []}

    @staticmethod
    def save(data):
        try:
            with open(HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"历史保存失败: {e}")

    @staticmethod
    def add_record(project, date, person, stats, tasks):
        history = HistoryManager.load()
        record = {
            "id": str(int(time.time() * 1000)),
            "project": project,
            "date": date,
            "person": person,
            "stats": stats,  # 统计信息
            "task_count": len(tasks),
            "timestamp": datetime.now().isoformat()
        }
        history["inspections"].insert(0, record)
        history["inspections"] = history["inspections"][:100]
        HistoryManager.save(history)


# ==================== 统计分析管理 ====================
class StatsManager:
    @staticmethod
    def analyze_tasks(tasks):
        """分析任务统计数据"""
        stats = {
            "total_images": len(tasks),
            "analyzed_images": 0,
            "total_issues": 0,
            "severe_safety": 0,
            "general_safety": 0,
            "severe_quality": 0,
            "general_quality": 0,
            "by_specialty": defaultdict(int),
            "avg_issues_per_image": 0,
            "detection_rate": 0
        }

        all_issues = []
        for task in tasks:
            if task.get("status") == "done":
                stats["analyzed_images"] += 1
                issues = task.get("edited_issues") or task.get("issues", [])
                all_issues.extend(issues)

        for issue in all_issues:
            stats["total_issues"] += 1
            level = issue.get("risk_level", "")

            if "严重安全" in level:
                stats["severe_safety"] += 1
            elif "一般安全" in level:
                stats["general_safety"] += 1
            elif "严重质量" in level:
                stats["severe_quality"] += 1
            elif "一般质量" in level:
                stats["general_quality"] += 1

            # 专业统计
            issue_text = issue.get("issue", "")
            if "【" in issue_text and "】" in issue_text:
                specialty = issue_text.split("】")[0].replace("【", "")
                stats["by_specialty"][specialty] += 1

        if stats["analyzed_images"] > 0:
            stats["avg_issues_per_image"] = round(stats["total_issues"] / stats["analyzed_images"], 1)
            stats["detection_rate"] = round((stats["analyzed_images"] / stats["total_images"]) * 100, 1)

        return stats


# ==================== JSON解析工具 ====================
def parse_json_safe(raw: str) -> Tuple[List[Dict], Optional[str]]:
    """增强型JSON解析"""
    if not raw:
        return [], "空响应"

    text = raw.strip().replace("```json", "").replace("```JSON", "").replace("```", "")
    start = text.find("[")
    end = text.rfind("]")
    if start == -1 or end == -1:
        return [], "未找到JSON数组"

    text = text[start:end + 1]
    text = text.replace("'", '"').replace("None", "null").replace("True", "true").replace("False", "false")

    for attempt in range(10):
        try:
            data = json.loads(text)
            break
        except json.JSONDecodeError as e:
            if "Expecting ',' delimiter" in str(e):
                text = text[:e.pos] + "," + text[e.pos:]
            elif "Expecting property name" in str(e):
                prev = text[:e.pos].rstrip()
                if prev.endswith(","):
                    comma_idx = text.rfind(",", 0, e.pos)
                    text = text[:comma_idx] + text[e.pos:]
                else:
                    break
            else:
                break
    else:
        objects = re.findall(r'\{[^{}]+\}', text)
        if objects:
            data = []
            for obj_str in objects:
                try:
                    data.append(json.loads(obj_str))
                except:
                    continue
            if data:
                return _normalize_issues(data), None
        return [], "JSON解析失败"

    if not isinstance(data, list):
        return [], "非数组格式"

    return _normalize_issues(data), None


def _normalize_issues(data: List[Dict]) -> List[Dict]:
    """标准化问题数据"""
    result = []
    for item in data:
        if not isinstance(item, dict):
            continue

        bbox = item.get("bbox")
        if bbox and len(bbox) == 4:
            try:
                bbox = [int(float(x)) for x in bbox]
                bbox = [max(-10000, min(100000, x)) for x in bbox]
                x1, x2 = sorted([bbox[0], bbox[2]])
                y1, y2 = sorted([bbox[1], bbox[3]])
                if x2 - x1 > 1 and y2 - y1 > 1:
                    bbox = [x1, y1, x2, y2]
                else:
                    bbox = None
            except:
                bbox = None
        else:
            bbox = None

        try:
            conf = float(item.get("confidence", 0.9))
        except:
            conf = 0.9

        # 自动判定category
        level = str(item.get("risk_level", ""))
        if "安全" in level:
            category = "安全隐患"
        else:
            category = "质量问题"

        result.append({
            "risk_level": level.strip(),
            "category": item.get("category", category),
            "issue": str(item.get("issue", "")).strip(),
            "regulation": str(item.get("regulation", "")).strip(),
            "correction": str(item.get("correction", "")).strip(),
            "bbox": bbox,
            "confidence": conf
        })

    return result


def calc_iou(box1, box2):
    """计算IoU"""
    if not box1 or not box2:
        return 0.0
    x1 = max(box1[0], box2[0])
    y1 = max(box1[1], box2[1])
    x2 = min(box1[2], box2[2])
    y2 = min(box1[3], box2[3])
    if x2 <= x1 or y2 <= y1:
        return 0.0
    inter = (x2 - x1) * (y2 - y1)
    area1 = (box1[2] - box1[0]) * (box1[3] - box1[1])
    area2 = (box2[2] - box2[0]) * (box2[3] - box2[1])
    union = area1 + area2 - inter
    return inter / union if union > 0 else 0.0


# ==================== 动态提示词生成 ====================
def build_specialist_prompt(role: str) -> str:
    """根据角色动态生成提示词"""
    db = REGULATION_DATABASE_V4.get(role, {})

    safety_focus = db.get("安全要点", "无特殊安全要求")
    quality_focus = db.get("质量要点", "参考相关规范")

    checklist_items = db.get("检查清单", [])
    checklist = "\n".join(checklist_items)

    false_positive = db.get("误判警示", "")

    return SPECIALIST_PROMPT_V4.format(
        role=role,
        safety_focus=safety_focus,
        quality_focus=quality_focus,
        role_specific_checklist=checklist,
        false_positive_examples=false_positive
    )


# 继续在第二部分...（由于字数限制）
# ==================== 图片导出工具 ====================
def ensure_export_dir():
    os.makedirs(EXPORT_IMG_DIR, exist_ok=True)
    return EXPORT_IMG_DIR


# ==================== 图片导出工具 (修复版) ====================
def ensure_export_dir():
    if not os.path.exists(EXPORT_IMG_DIR):
        os.makedirs(EXPORT_IMG_DIR, exist_ok=True)
    return EXPORT_IMG_DIR


def draw_on_image(img: QImage, issues: List[Dict], anns: List[Dict]) -> QImage:
    """核心绘制函数：只绘制人工标注，不绘制AI识别框"""
    if img.isNull():
        return img

    out = img.copy()
    painter = QPainter(out)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)

    # 🔴 问题3修复：移除AI识别框的绘制，只保留人工标注
    # 原先的 issues 识别框代码已删除

    # --- 绘制 人工标注 (anns) ---
    if anns:
        for ann in anns:
            ann_type = ann.get("type")
            color = QColor(ann.get("color", "#FF0000"))
            width = int(ann.get("width", 6))

            pen = QPen(color, width)
            pen.setCapStyle(Qt.PenCapStyle.RoundCap)
            pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
            painter.setPen(pen)
            painter.setBrush(Qt.BrushStyle.NoBrush)

            if ann_type == "rect":
                x1, y1, x2, y2 = ann.get("bbox", [0, 0, 0, 0])
                painter.drawRect(QRectF(x1, y1, x2 - x1, y2 - y1))
            elif ann_type == "ellipse":
                x1, y1, x2, y2 = ann.get("bbox", [0, 0, 0, 0])
                painter.drawEllipse(QRectF(x1, y1, x2 - x1, y2 - y1))
            elif ann_type == "arrow":
                x1, y1 = ann.get("p1", [0, 0])
                x2, y2 = ann.get("p2", [0, 0])
                painter.drawLine(QPointF(x1, y1), QPointF(x2, y2))
                # 画箭头头部
                angle = math.atan2(y2 - y1, x2 - x1)
                head_len = 30
                head_ang = math.radians(25)
                p1 = QPointF(x2 - head_len * math.cos(angle - head_ang), y2 - head_len * math.sin(angle - head_ang))
                p2 = QPointF(x2 - head_len * math.cos(angle + head_ang), y2 - head_len * math.sin(angle + head_ang))
                painter.drawLine(QPointF(x2, y2), p1)
                painter.drawLine(QPointF(x2, y2), p2)
            elif ann_type == "text":
                x, y = ann.get("pos", [0, 0])
                text = ann.get("text", "")
                font_size = int(ann.get("font_size", 32))

                font = QFont("SimHei", font_size, QFont.Weight.Bold)
                painter.setFont(font)

                # 文字描边效果（提升可读性）
                path = QPainterPath()
                path.addText(QPointF(x, y), font, text)

                # 描边
                painter.setPen(QPen(QColor(255, 255, 255), 6))
                painter.drawPath(path)

                # 填充
                painter.setPen(QPen(color))
                painter.drawText(QPointF(x, y), text)

    painter.end()
    return out


def export_marked_image(orig_path, issues, anns, out_path):
    """导出带标注图片（入口函数）"""
    if not os.path.exists(orig_path):
        return False

    img = QImage(orig_path)
    if img.isNull():
        return False

    # 调用统一的绘制函数
    final_img = draw_on_image(img, issues, anns)

    return final_img.save(out_path, "PNG")


# ==================== Word报告生成 ====================
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
    def replace_placeholders(doc, info):
        """替换模板占位符（增强兼容性）"""

        # 获取数值
        s_safe = str(info.get("severe_safety", 0))
        g_safe = str(info.get("general_safety", 0))
        s_qual = str(info.get("severe_quality", 0))
        g_qual = str(info.get("general_quality", 0))
        total = str(info.get("total_issues", 0))

        replacements = {
            "{{项目公司名称}}": info.get("project_company", ""),
            "{{项目名称}}": info.get("project_name", ""),
            "{{检查部位}}": info.get("check_area", ""),
            "{{检查人员}}": info.get("check_person", ""),
            "{{被检查单位}}": info.get("inspected_unit", ""),
            "{{检查内容}}": info.get("check_content", ""),
            "{{项目概况}}": info.get("project_overview", ""),
            "{{检查日期}}": info.get("check_date", ""),
            "{{整改期限}}": info.get("rectification_deadline", ""),

            # === 数值统计 (同时支持"问题"和"缺陷"两种写法，防止模板不匹配) ===
            "{{严重安全隐患数}}": s_safe,
            "{{一般安全隐患数}}": g_safe,

            "{{严重质量问题数}}": s_qual,  # 兼容写法 A
            "{{严重质量缺陷数}}": s_qual,  # 兼容写法 B (推荐)

            "{{一般质量问题数}}": g_qual,  # 兼容写法 A
            "{{一般质量缺陷数}}": g_qual,  # 兼容写法 B (推荐)

            "{{问题总数}}": total,
            "{{隐患总数}}": total  # 兼容写法
        }

        for para in doc.paragraphs:
            for key, val in replacements.items():
                if key in para.text:
                    para.text = para.text.replace(key, val)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for key, val in replacements.items():
                            if key in para.text:
                                para.text = para.text.replace(key, val)

    @staticmethod
    def generate(tasks, save_path, info, template="模板.docx"):
        """生成Word报告（V4.0优化：分类显示）"""
        if os.path.exists(template):
            doc = Document(template)
        else:
            doc = Document()
            doc.add_paragraph(f"【注意】未找到模板 {template}，使用空白格式")

        WordReportGenerator.replace_placeholders(doc, info)

        valid_tasks = [t for t in tasks if t.get("status") == "done" or t.get("annotations")]

        if not valid_tasks:
            doc.add_paragraph("【提示】没有已完成分析的任务")
            doc.save(save_path)
            return

        for idx, task in enumerate(valid_tasks, 1):
            table = doc.add_table(rows=1, cols=1)
            table.style = 'Table Grid'
            cell = table.cell(0, 0)

            p_title = cell.paragraphs[0]
            title = f"检查点位 {idx}"
            group = (task.get("meta") or {}).get("group")
            if group:
                title += f"（{group}）"
            run_title = p_title.add_run(title)
            WordReportGenerator.set_font(run_title, size=14, bold=True)

            issues = task.get("edited_issues") or task.get("issues", [])

            # V4.0优化：按类型分类
            severe_safety, general_safety, severe_quality, general_quality = [], [], [], []
            corrections = []

            for item in issues:
                level = item.get("risk_level", "")
                category = item.get("category", "")
                issue = item.get("issue", "")
                reg = item.get("regulation", "")
                corr = item.get("correction", "")

                if not issue:
                    continue

                full_desc = issue
                if reg and reg not in ["无", "常识"]:
                    full_desc += f"（依据：{reg}）"

                if "严重安全" in level:
                    severe_safety.append(full_desc)
                elif "一般安全" in level:
                    general_safety.append(full_desc)
                elif "严重质量" in level:
                    severe_quality.append(full_desc)
                elif "一般质量" in level:
                    general_quality.append(full_desc)

                if corr:
                    corrections.append(corr)

            # 添加分类内容
            def add_section(label, texts, color=None):
                if not texts:
                    return
                p = cell.add_paragraph()
                run_label = p.add_run(label)
                WordReportGenerator.set_font(run_label, bold=True, size=12, color=color)
                merged = "；".join(texts) + "。"
                run_text = p.add_run(merged)
                WordReportGenerator.set_font(run_text, size=11)

            add_section("🔴 严重安全隐患：", severe_safety, RGBColor(211, 47, 47))
            add_section("🟠 一般安全隐患：", general_safety, RGBColor(245, 124, 0))
            add_section("🟡 严重质量问题：", severe_quality, RGBColor(230, 74, 25))
            add_section("🟡 一般质量问题：", general_quality, RGBColor(255, 167, 38))

            # 整改要求
            p_corr = cell.add_paragraph()
            run_label = p_corr.add_run("整改要求：")
            WordReportGenerator.set_font(run_label, bold=True, size=12)

            if corrections:
                for i, corr in enumerate(corrections, 1):
                    p = cell.add_paragraph()
                    run = p.add_run(f"{i}. {corr}")
                    WordReportGenerator.set_font(run, size=11, color=RGBColor(46, 125, 50))
            else:
                run_text = p_corr.add_run("无")
                WordReportGenerator.set_font(run_text, size=11)

            # 插入图片
            img_path = task.get("export_image_path") or task.get("path")
            if img_path and os.path.exists(img_path):
                p_img = cell.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                try:
                    p_img.add_run().add_picture(img_path, width=Cm(14))
                except Exception as e:
                    p_img.add_run(f"[图片加载失败: {e}]")

        doc.save(save_path)


# ==================== 分析线程（修复版：支持日志回传） ====================
class AnalysisWorker(QThread):
    result_ready = pyqtSignal(str, dict)
    # 新增：用于发送日志到主界面的信号 (消息内容, 颜色等级)
    log_signal = pyqtSignal(str, str)

    PRIORITY_MAP = {
        "管道": 10, "暖通": 10, "给排水": 10,
        "电气": 10, "机械": 10,
        "结构": 10, "防水": 10, "基坑": 10,
        "消防": 9, "环保": 8,
        "安全": 5
    }

    def __init__(self, task, config, prompt):
        super().__init__()
        self.task = task
        self.config = config
        self.prompt = prompt

    def run(self):
        start_time = time.time()

        # 辅助日志发送函数
        def log(msg, level="info"):
            self.log_signal.emit(msg, level)

        provider = self.config["current_provider"]
        api_key = self.config["api_key"]
        preset = self.config["provider_presets"].get(provider, {})
        base_url = preset.get("base_url")
        model = preset.get("model")

        if provider == "自定义":
            custom = self.config.get("custom_provider_settings", {})
            base_url = custom.get("base_url")
            model = custom.get("model")

        if not all([api_key, base_url, model]):
            self.result_ready.emit(self.task['id'], {
                "ok": False, "error": "配置缺失", "elapsed_sec": 0
            })
            log(f"[{self.task['name']}] 配置缺失，无法启动", "error")
            return

        img_b64 = self._compress_image(self.task["path"])
        if not img_b64:
            self.result_ready.emit(self.task['id'], {
                "ok": False, "error": "图片加载失败", "elapsed_sec": 0
            })
            log(f"[{self.task['name']}] 图片加载失败", "error")
            return

        try:
            http_client = httpx.Client(http2=False, verify=False, timeout=90)
            client = OpenAI(api_key=api_key, base_url=base_url, http_client=http_client)

            # 1. Router 分诊
            experts = ["安全"]
            log(f"🔍 [{self.task['name']}] 开始智能分诊...", "info")

            router_resp = self._call_llm(client, model, ROUTER_SYSTEM_PROMPT, img_b64, role="Router")

            if router_resp:
                try:
                    detected = json.loads(router_resp.strip())
                    if isinstance(detected, list):
                        for expert in detected:
                            if expert not in experts:
                                experts.append(expert)
                except Exception as e:
                    log(f"⚠️ Router解析异常: {e}", "warning")

            if len(experts) == 1:
                experts.extend(["结构", "电气"])

            experts = experts[:4]
            log(f"✅ [{self.task['name']}] 专家团队: {experts}", "success")

            # 2. Specialist 检查
            all_issues = []

            for role in experts:
                log(f"🔬 [{self.task['name']}] {role}专家正在检查...", "info")
                time.sleep(1.0)

                knowledge = REGULATION_DATABASE_V5.get(role, REGULATION_DATABASE_V5["安全"])
                checklist_str = "\n".join([f"- {item}" for item in knowledge.get("checklist", [])])

                specialist_prompt = SPECIALIST_PROMPT_TEMPLATE.format(
                    role=role,
                    role_desc=knowledge.get("role_desc", "工程专家"),
                    norms=knowledge.get("norms", "相关国家规范"),
                    checklist=checklist_str,
                    anti_hallucination=knowledge.get("anti_hallucination", "无")
                )

                resp = self._call_llm(client, model, specialist_prompt, img_b64, role=role)

                if resp:
                    issues, err = parse_json_safe(resp)
                    if err:
                        log(f"⚠️ {role}结果解析失败: {err}", "warning")
                    else:
                        for item in issues:
                            if not item["issue"].startswith("【"):
                                item["issue"] = f"【{role}】{item['issue']}"
                        all_issues.extend(issues)
                        log(f"    - {role} 发现 {len(issues)} 个问题", "info")

            # 3. 去重
            before_cnt = len(all_issues)
            final_issues = self._deduplicate(all_issues)
            log(f"✅ [{self.task['name']}] 分析完成 (去重: {before_cnt}->{len(final_issues)})", "success")

            elapsed = round(time.time() - start_time, 2)

            self.result_ready.emit(self.task["id"], {
                "ok": True,
                "issues": final_issues,
                "elapsed_sec": elapsed,
                "provider": provider,
                "model": model
            })

        except Exception as e:
            elapsed = round(time.time() - start_time, 2)
            err_msg = str(e)
            log(f"❌ [{self.task['name']}] 线程异常: {err_msg}", "error")
            print(traceback.format_exc())  # 依然保留控制台输出以便调试
            self.result_ready.emit(self.task["id"], {
                "ok": False,
                "error": err_msg,
                "issues": [],
                "elapsed_sec": elapsed
            })

    # ... _compress_image, _call_llm, _dual_model_verify, _deduplicate 等方法保持原样 ...
    # (为了节省篇幅，请确保你保留了 AnalysisWorker 类中其他的辅助方法)
    def _compress_image(self, path):
        try:
            from PyQt6.QtGui import QImageReader
            reader = QImageReader(path)
            if not reader.canRead(): return ""
            size = reader.size()
            if size.width() > 1536 or size.height() > 1536:
                reader.setScaledSize(size.scaled(1536, 1536, Qt.AspectRatioMode.KeepAspectRatio))
            img = reader.read()
            if img.isNull(): return ""
            ba = QByteArray()
            buf = QBuffer(ba)
            buf.open(QIODevice.OpenModeFlag.WriteOnly)
            img.save(buf, "JPEG", 85)
            return ba.toBase64().data().decode()
        except:
            return ""

    def _call_llm(self, client, model, system_prompt, img_b64, role="Unknown"):
        try:
            messages = [
                {"role": "system", "content": system_prompt},
                {"role": "user",
                 "content": [{"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img_b64}"}},
                             {"type": "text", "text": "输出JSON"}]}
            ]
            temp = 0.2 if role == "Router" else 0.3
            resp = client.chat.completions.create(model=model, messages=messages, temperature=temp)
            return resp.choices[0].message.content
        except:
            return None

    def _deduplicate(self, issues):
        # 简单保留原有的去重逻辑
        if not issues: return []
        for item in issues:
            role = item["issue"].split("】")[0].replace("【", "")
            item["_score"] = self.PRIORITY_MAP.get(role, 5)
        issues.sort(key=lambda x: x["_score"], reverse=True)
        unique = []
        for cand in issues:
            is_dup = False
            cand_bbox = cand.get("bbox")
            if not cand_bbox:
                for exist in unique:
                    if cand["issue"][:10] == exist["issue"][:10]:
                        is_dup = True;
                        break
            else:
                for exist in unique:
                    exist_bbox = exist.get("bbox")
                    if exist_bbox and calc_iou(cand_bbox, exist_bbox) > 0.4:
                        is_dup = True;
                        break
            if not is_dup:
                unique.append(cand)
        return unique


# ==================== UI组件：可编辑文字 ====================
class EditableTextItem(QGraphicsTextItem):
    def __init__(self, text, callback=None):
        super().__init__(text)
        self.callback = callback
        self.setFlags(
            QGraphicsItem.GraphicsItemFlag.ItemIsMovable |
            QGraphicsItem.GraphicsItemFlag.ItemIsSelectable |
            QGraphicsItem.GraphicsItemFlag.ItemIsFocusable
        )
        self.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)
        self.setDefaultTextColor(QColor("#FF0000"))
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.setTextInteractionFlags(Qt.TextInteractionFlag.TextEditorInteraction)
            self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, False)
            self.setFocus()
            self.setCursor(Qt.CursorShape.IBeamCursor)
            super().mouseDoubleClickEvent(event)
            if self.scene() and self.scene().views():
                self.scene().views()[0].setDragMode(QGraphicsView.DragMode.NoDrag)

    def focusOutEvent(self, event):
        self.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        cursor = self.textCursor()
        cursor.clearSelection()
        self.setTextCursor(cursor)
        if self.callback:
            self.callback(self)
        if self.scene() and self.scene().views():
            view = self.scene().views()[0]
            if hasattr(view, "_tool") and view._tool == "none":
                view.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
        super().focusOutEvent(event)


# ==================== UI组件：图片标注视图 ====================
class AnnotatableImageView(QGraphicsView):
    annotation_changed = pyqtSignal()
    tool_reset = pyqtSignal()

    TOOL_NONE = "none"
    TOOL_RECT = "rect"
    TOOL_ELLIPSE = "ellipse"
    TOOL_ARROW = "arrow"
    TOOL_TEXT = "text"
    TOOL_ISSUE_TAG = "issue_tag"

    def __init__(self):
        super().__init__(QGraphicsScene())
        self._pix_item = QGraphicsPixmapItem()
        self._pix_item.setZValue(-1000)
        self._pix_item.setAcceptedMouseButtons(Qt.MouseButton.NoButton)
        self.scene().addItem(self._pix_item)
        self._tool = self.TOOL_NONE
        self._dragging = False
        self._start_pt = None
        self._temp_end_pt = None
        self._img_path = None
        self._ai_issues = []
        self._draw_color = "#FF0000"
        self._draw_width = 6
        self._img_size = (1, 1)
        self.setRenderHints(QPainter.RenderHint.Antialiasing | QPainter.RenderHint.SmoothPixmapTransform)
        self.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setMouseTracking(True)
        self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)

    def set_image(self, path):
        img = QImage(path)
        if img.isNull():
            return
        pix = QPixmap.fromImage(img)
        self._img_path = path
        self._img_size = (pix.width(), pix.height())
        self._pix_item.setPixmap(pix)
        self.scene().setSceneRect(QRectF(0, 0, pix.width(), pix.height()))
        self.resetTransform()
        self.fitInView(self.sceneRect(), Qt.AspectRatioMode.KeepAspectRatio)

    def delete_selected_items(self):
        """删除选中的标注项"""
        has_deleted = False
        # 遍历所有被选中的图元
        for item in self.scene().selectedItems():
            # 保护底图不被删除
            if item != self._pix_item:
                self.scene().removeItem(item)
                has_deleted = True

        if has_deleted:
            self.annotation_changed.emit()

    def set_tool(self, tool):
        self._tool = tool
        self._dragging = False
        if tool == self.TOOL_NONE:
            self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
            self.setCursor(Qt.CursorShape.OpenHandCursor)
        else:
            self.setDragMode(QGraphicsView.DragMode.NoDrag)
            self.setCursor(Qt.CursorShape.CrossCursor)

    def set_ai_issues(self, issues):
        self._ai_issues = issues or []

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            item = self.itemAt(event.position().toPoint())
            if isinstance(item, QGraphicsTextItem):
                self.setDragMode(QGraphicsView.DragMode.NoDrag)
                super().mousePressEvent(event)
                return
            if isinstance(item, QGraphicsItem) and item is not self._pix_item and self._tool == self.TOOL_NONE:
                self.setDragMode(QGraphicsView.DragMode.NoDrag)
                super().mousePressEvent(event)
                return
            if self._tool == self.TOOL_ISSUE_TAG:
                pos = self._to_scene_pt(event.position().toPoint())
                self._handle_issue_tag(pos)
                return
            if self._tool != self.TOOL_NONE:
                self._dragging = True
                self._start_pt = self._to_scene_pt(event.position().toPoint())
                self._temp_end_pt = self._start_pt
                return
            if self._tool == self.TOOL_NONE:
                self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._dragging and self._tool != self.TOOL_NONE:
            self._temp_end_pt = self._to_scene_pt(event.position().toPoint())
            self.viewport().update()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._dragging:
            self._finish_drawing(event)
        super().mouseReleaseEvent(event)
        if self._tool == self.TOOL_NONE and not self.scene().focusItem():
            self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)

    def paintEvent(self, event):
        super().paintEvent(event)
        if self._dragging and self._start_pt and self._temp_end_pt:
            painter = QPainter(self.viewport())
            painter.setPen(QPen(QColor(self._draw_color), 2, Qt.PenStyle.DashLine))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            p1 = self.mapFromScene(self._start_pt)
            p2 = self.mapFromScene(self._temp_end_pt)
            x = min(p1.x(), p2.x())
            y = min(p1.y(), p2.y())
            w = abs(p1.x() - p2.x())
            h = abs(p1.y() - p2.y())
            if self._tool in (self.TOOL_RECT, self.TOOL_ELLIPSE):
                if self._tool == self.TOOL_ELLIPSE:
                    painter.drawEllipse(x, y, w, h)
                else:
                    painter.drawRect(x, y, w, h)
            elif self._tool == self.TOOL_ARROW:
                painter.drawLine(p1, p2)

    def wheelEvent(self, event):
        if event.angleDelta().y() > 0:
            self.scale(1.25, 1.25)
        else:
            self.scale(0.8, 0.8)

    def _to_scene_pt(self, view_pos):
        sp = self.mapToScene(view_pos)
        x = min(max(sp.x(), 0), self._img_size[0])
        y = min(max(sp.y(), 0), self._img_size[1])
        return QPointF(x, y)

    def _finish_drawing(self, event):
        self._dragging = False
        start = self._start_pt
        end = self._to_scene_pt(event.position().toPoint())
        self._start_pt = None
        self._temp_end_pt = None
        self.viewport().update()
        if not start or not end:
            return
        if abs(start.x() - end.x()) < 5 and abs(start.y() - end.y()) < 5:
            if self._tool == self.TOOL_TEXT:
                self._create_text(start)
            return
        data = None
        if self._tool in (self.TOOL_RECT, self.TOOL_ELLIPSE):
            x1, x2 = sorted([start.x(), end.x()])
            y1, y2 = sorted([start.y(), end.y()])
            data = {"type": self._tool, "bbox": [int(x1), int(y1), int(x2), int(y2)],
                    "color": self._draw_color, "width": self._draw_width}
        elif self._tool == self.TOOL_ARROW:
            data = {"type": "arrow", "p1": [int(start.x()), int(start.y())],
                    "p2": [int(end.x()), int(end.y())], "color": self._draw_color, "width": self._draw_width}
        elif self._tool == self.TOOL_TEXT:
            self._create_text(end)
            return
        if data:
            self._create_item_from_data(data)
            self.annotation_changed.emit()

    def _create_text(self, pos):
        # 🔴 简化修改：重新设计对话框，使用固定默认值
        dialog = QDialog(self)
        dialog.setWindowTitle("输入标注文字")
        dialog.resize(400, 150)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(15, 15, 15, 15)

        # 标签
        label = QLabel("请输入标注文字内容:")
        label.setStyleSheet("color: #BDBDBD; font-size: 14px; font-weight: bold; margin-bottom: 5px;")
        layout.addWidget(label)

        # 输入框（使用QLineEdit代替QInputDialog）
        text_input = QLineEdit()
        text_input.setStyleSheet("""
            QLineEdit {
                color: #000000;
                background: white;
                font-size: 13px;
                padding: 8px;
                border: 2px solid #BDBDBD;
                border-radius: 4px;
            }
            QLineEdit:focus {
                border: 2px solid #2196F3;
            }
        """)
        text_input.setPlaceholderText("在此输入文字...")
        layout.addWidget(text_input)

        # 按钮
        btn_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        btn_box.setStyleSheet("""
            QPushButton {
                background: #2196F3;
                color: white;
                font-size: 13px;
                font-weight: bold;
                padding: 6px 20px;
                border: none;
                border-radius: 4px;
                min-width: 80px;
            }
            QPushButton:hover {
                background: #1976D2;
            }
            QPushButton:pressed {
                background: #1565C0;
            }
        """)
        btn_box.accepted.connect(dialog.accept)
        btn_box.rejected.connect(dialog.reject)
        layout.addWidget(btn_box)

        # 显示对话框
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        text = text_input.text().strip()
        if not text:
            return

        # 🔴 简化修改：使用固定默认值（红色，32px）
        color = "#FF0000"
        font_size = 32

        data = {"type": "text", "pos": [int(pos.x()), int(pos.y())], "text": text,
                "color": color, "font_size": font_size}
        self._create_item_from_data(data)
        self.annotation_changed.emit()

    def _handle_issue_tag(self, pos):
        if not self._ai_issues:
            QMessageBox.information(self, "提示", "当前没有AI识别的问题可引用")
            self.tool_reset.emit()
            return
        QTimer.singleShot(0, lambda: self._open_issue_dialog(pos))

    def _open_issue_dialog(self, pos):
        dlg = IssueSelectionDialog(self, self._ai_issues)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            data = {"type": "text", "pos": [int(pos.x()), int(pos.y())],
                    "text": dlg.selected_text, "color": dlg.selected_color, "font_size": 28}
            self._create_item_from_data(data)
            self.annotation_changed.emit()

    def _create_item_from_data(self, data):
        item_type = data.get("type")
        color = QColor(data.get("color", "#FF0000"))
        width = int(data.get("width", 6))
        item = None
        if item_type == "text":
            x, y = data.get("pos", [0, 0])
            text = data.get("text", "")
            fs = int(data.get("font_size", 28))
            item = EditableTextItem(text, callback=lambda _: self.annotation_changed.emit())
            font = QFont()
            font.setPointSize(fs)
            font.setBold(True)
            item.setFont(font)
            item.setDefaultTextColor(color)
            item.setPos(x, y)
        elif item_type == "rect":
            x1, y1, x2, y2 = data.get("bbox", [0, 0, 0, 0])
            item = QGraphicsRectItem(QRectF(x1, y1, x2 - x1, y2 - y1))
        elif item_type == "ellipse":
            x1, y1, x2, y2 = data.get("bbox", [0, 0, 0, 0])
            item = QGraphicsEllipseItem(QRectF(x1, y1, x2 - x1, y2 - y1))
        elif item_type == "arrow":
            x1, y1 = data.get("p1", [0, 0])
            x2, y2 = data.get("p2", [0, 0])
            path = QPainterPath()
            path.moveTo(x1, y1)
            path.lineTo(x2, y2)
            item = QGraphicsPathItem(path)
        if item:
            if item_type != "text":
                pen = QPen(color, width)
                pen.setCapStyle(Qt.PenCapStyle.RoundCap)
                pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
                item.setPen(pen)
                item.setFlags(QGraphicsItem.GraphicsItemFlag.ItemIsMovable |
                              QGraphicsItem.GraphicsItemFlag.ItemIsSelectable)
            item.setData(Qt.ItemDataRole.UserRole, data.copy())
            self.scene().addItem(item)

    def set_user_annotations(self, anns):
        self.blockSignals(True)
        try:
            self.clear_annotations()
            for ann in anns:
                self._create_item_from_data(ann)
        finally:
            self.blockSignals(False)

    def get_user_annotations(self):
        self.scene().clearFocus()
        anns = []
        for item in self.scene().items():
            if item == self._pix_item:
                continue
            raw_data = item.data(Qt.ItemDataRole.UserRole)
            if not raw_data:
                continue
            data = raw_data.copy()
            if isinstance(item, EditableTextItem):
                data["text"] = item.toPlainText()
                data["pos"] = [int(item.pos().x()), int(item.pos().y())]
                # 🔴 保存当前颜色和字体大小
                data["color"] = item.defaultTextColor().name()
                data["font_size"] = item.font().pointSize() or 32
            elif isinstance(item, (QGraphicsRectItem, QGraphicsEllipseItem)):
                r = item.sceneBoundingRect()
                data["bbox"] = [int(r.left()), int(r.top()), int(r.right()), int(r.bottom())]
            elif isinstance(item, QGraphicsPathItem) and data.get("type") == "arrow":
                offset = item.pos()
                p1 = data.get("p1", [0, 0])
                p2 = data.get("p2", [0, 0])
                data["p1"] = [int(p1[0] + offset.x()), int(p1[1] + offset.y())]
                data["p2"] = [int(p2[0] + offset.x()), int(p2[1] + offset.y())]
            anns.append(data)
        return anns

    def clear_annotations(self):
        for item in list(self.scene().items()):
            if item != self._pix_item:
                self.scene().removeItem(item)
        self.annotation_changed.emit()

    def delete_selected_items(self):
        for item in self.scene().selectedItems():
            if item != self._pix_item:
                self.scene().removeItem(item)
        self.annotation_changed.emit()

    def undo(self):
        items = [i for i in self.scene().items() if i != self._pix_item]
        if items:
            self.scene().removeItem(items[0])
            self.annotation_changed.emit()


# ==================== UI组件：问题选择对话框 ====================
class IssueSelectionDialog(QDialog):
    """选择问题进行引用"""

    def __init__(self, parent, issues):
        super().__init__(parent)
        self.setWindowTitle("选择要引用的问题")
        self.resize(550, 350)
        self.selected_text = ""
        self.selected_color = "#FF0000"

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("请选择此标注关联的问题："))

        self.list_widget = QListWidget()
        self.list_widget.setStyleSheet("""
            QListWidget {
                color: #000000;
                font-size: 13px;
                background: white;
            }
            QListWidget::item {
                color: #212121;
                padding: 8px;
                border-bottom: 1px solid #F0F0F0;
            }
            QListWidget::item:selected {
                background: #E3F2FD;
                color: #000000;
            }
        """)

        for idx, item in enumerate(issues, 1):
            level = item.get("risk_level", "一般")
            category = item.get("category", "")
            desc = item.get("issue", "未知问题")
            display = f"{idx}. [{category}] {level} - {desc[:50]}"

            list_item = QListWidgetItem(display)
            list_item.setData(Qt.ItemDataRole.UserRole, desc)
            list_item.setData(Qt.ItemDataRole.UserRole + 1, level)
            list_item.setForeground(QColor("#000000"))  # 黑色文字
            self.list_widget.addItem(list_item)

        layout.addWidget(self.list_widget)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        # 🔴 最终修改：优化按钮样式，确保文字清晰可见
        btns.setStyleSheet("""
            QPushButton {
                background: #2196F3;
                color: white;
                font-size: 13px;
                font-weight: bold;
                padding: 6px 20px;
                border: none;
                border-radius: 4px;
                min-width: 80px;
            }
            QPushButton:hover {
                background: #1976D2;
            }
            QPushButton:pressed {
                background: #1565C0;
            }
        """)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def accept(self):
        item = self.list_widget.currentItem()
        if item:
            desc = item.data(Qt.ItemDataRole.UserRole)
            level = item.data(Qt.ItemDataRole.UserRole + 1)

            short_desc = desc[:15] + "..." if len(desc) > 15 else desc
            self.selected_text = short_desc

            # 🔴 修改：使用更亮的颜色
            if "严重安全" in level:
                self.selected_color = "#D32F2F"  # 更亮的红色
            elif "一般安全" in level:
                self.selected_color = "#F57C00"  # 更亮的橙色
            elif "严重质量" in level:
                self.selected_color = "#E64A19"  # 更亮的深橙色
            else:
                self.selected_color = "#FFA726"  # 更亮的黄色

        super().accept()


# ==================== UI组件：问题编辑对话框 ====================
class IssueEditDialog(QDialog):
    """编辑问题详情"""

    def __init__(self, parent, item):
        super().__init__(parent)
        self.setWindowTitle("编辑问题")
        self.resize(650, 550)
        self.item = dict(item)

        layout = QVBoxLayout(self)
        form = QFormLayout()

        # 风险等级
        self.cbo_level = QComboBox()
        self.cbo_level.addItems([
            "严重安全隐患", "一般安全隐患",
            "严重质量问题", "一般质量问题"
        ])
        if self.item.get("risk_level"):
            idx = self.cbo_level.findText(self.item["risk_level"])
            if idx >= 0:
                self.cbo_level.setCurrentIndex(idx)

        # 问题类型
        self.cbo_category = QComboBox()
        self.cbo_category.addItems(["安全隐患", "质量问题"])
        if self.item.get("category"):
            idx = self.cbo_category.findText(self.item["category"])
            if idx >= 0:
                self.cbo_category.setCurrentIndex(idx)

        # 问题描述
        self.txt_issue = QPlainTextEdit()
        self.txt_issue.setPlainText(self.item.get("issue", ""))
        self.txt_issue.setMinimumHeight(100)

        # 规范依据
        self.txt_reg = QLineEdit()
        self.txt_reg.setText(self.item.get("regulation", ""))

        # 整改建议
        self.txt_corr = QPlainTextEdit()
        self.txt_corr.setPlainText(self.item.get("correction", ""))
        self.txt_corr.setMinimumHeight(100)

        form.addRow("风险等级:", self.cbo_level)
        form.addRow("问题类型:", self.cbo_category)
        form.addRow("问题描述:", self.txt_issue)
        form.addRow("规范依据:", self.txt_reg)
        form.addRow("整改建议:", self.txt_corr)

        layout.addLayout(form)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Cancel |
            QDialogButtonBox.StandardButton.Ok
        )
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def get_value(self):
        return {
            "risk_level": self.cbo_level.currentText(),
            "category": self.cbo_category.currentText(),
            "issue": self.txt_issue.toPlainText().strip(),
            "regulation": self.txt_reg.text().strip(),
            "correction": self.txt_corr.toPlainText().strip(),
            "bbox": self.item.get("bbox"),
            "confidence": self.item.get("confidence")
        }


# ==================== UI组件：现代化问题卡片 ====================
class ModernRiskCard(QFrame):
    """V4.0现代化问题卡片"""
    edit_requested = pyqtSignal(dict)
    delete_requested = pyqtSignal(dict)

    def __init__(self, item):
        super().__init__()
        self.item = item
        self.setFrameShape(QFrame.Shape.StyledPanel)

        level = item.get("risk_level", "")
        category = item.get("category", "质量问题")

        # V4.0配色方案
        if "严重安全" in level:
            bg_color = THEME_COLORS["severe_safety_bg"]
            border_color = THEME_COLORS["severe_safety"]
            icon = "🔴"
        elif "一般安全" in level:
            bg_color = THEME_COLORS["general_safety_bg"]
            border_color = THEME_COLORS["general_safety"]
            icon = "🟠"
        elif "严重质量" in level:
            bg_color = THEME_COLORS["severe_quality_bg"]
            border_color = THEME_COLORS["severe_quality"]
            icon = "🟡"
        else:
            bg_color = THEME_COLORS["general_quality_bg"]
            border_color = THEME_COLORS["general_quality"]
            icon = "🟡"

        self.setStyleSheet(f"""
            ModernRiskCard {{
                background-color: {bg_color};
                border-left: 5px solid {border_color};
                border-radius: 6px;
                padding: 8px;
                margin: 4px 2px;
            }}
            ModernRiskCard:hover {{
                background-color: {self._lighten(bg_color)};
            }}
        """)

        # 添加阴影效果
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(10)
        shadow.setColor(QColor(0, 0, 0, 30))
        shadow.setOffset(2, 2)
        self.setGraphicsEffect(shadow)

        layout = QVBoxLayout(self)

        # 头部：图标 + 等级 + 类型标签 + 按钮
        header = QHBoxLayout()
        lbl_level = QLabel(f"{icon} <b>{level}</b>")
        lbl_level.setStyleSheet("font-size: 13px;")
        header.addWidget(lbl_level)

        # 类型标签
        tag = QLabel(category)
        tag.setStyleSheet(f"""
            QLabel {{
                background: {border_color};
                color: white;
                padding: 3px 10px;
                border-radius: 3px;
                font-size: 11px;
                font-weight: bold;
            }}
        """)
        header.addWidget(tag)
        header.addStretch()

        # 编辑按钮
        btn_edit = QPushButton("✏️ 编辑")
        btn_edit.setFixedSize(70, 28)
        btn_edit.setStyleSheet("""
            QPushButton {
                background: white;
                border: 1px solid #E0E0E0;
                border-radius: 4px;
            }
            QPushButton:hover {
                background: #F5F5F5;
            }
        """)
        btn_edit.clicked.connect(lambda: self.edit_requested.emit(item))
        header.addWidget(btn_edit)

        # 删除按钮
        btn_del = QPushButton("🗑️ 删除")
        btn_del.setFixedSize(70, 28)
        btn_del.setStyleSheet("""
            QPushButton {
                background: white;
                border: 1px solid #E0E0E0;
                border-radius: 4px;
                color: #F44336;
            }
            QPushButton:hover {
                background: #FFEBEE;
            }
        """)
        btn_del.clicked.connect(lambda: self.delete_requested.emit(item))
        header.addWidget(btn_del)

        layout.addLayout(header)

        # 问题描述
        issue = item.get("issue", "")
        lbl_issue = QLabel(issue[:250] + "..." if len(issue) > 250 else issue)
        lbl_issue.setWordWrap(True)
        lbl_issue.setStyleSheet("font-size: 13px; color: #212121; margin: 8px 0;")
        layout.addWidget(lbl_issue)

        # 规范依据（灰色小字）
        reg = item.get("regulation", "")
        if reg:
            lbl_reg = QLabel(f"📋 依据：{reg}")
            lbl_reg.setStyleSheet("font-size: 11px; color: #424242; margin: 4px 0;")
            lbl_reg.setWordWrap(True)
            layout.addWidget(lbl_reg)

        # 整改建议（绿色强调）
        corr = item.get("correction", "")
        if corr:
            lbl_corr = QLabel(f"✅ 整改：{corr[:200]}")
            lbl_corr.setWordWrap(True)
            lbl_corr.setStyleSheet("""
                font-size: 12px; 
                color: #2E7D32; 
                font-weight: bold;
                margin: 4px 0;
                padding: 6px;
                background: rgba(76, 175, 80, 0.1);
                border-radius: 4px;
            """)
            layout.addWidget(lbl_corr)

    def _lighten(self, color):
        """提亮颜色"""
        c = QColor(color)
        return QColor(min(255, c.red() + 10), min(255, c.green() + 10), min(255, c.blue() + 10)).name()


# ==================== UI组件：统计卡片 ====================
class StatsCard(QFrame):
    def __init__(self, title, value, color, icon):
        super().__init__()
        # 核心修改：高度固定80px，宽度230px
        self.setFixedSize(QSize(230, 80))

        # 样式：纯色背景+圆角
        self.setStyleSheet(f"""
            StatsCard {{
                background-color: {color};
                border-radius: 8px;
                color: white;
            }}
        """)

        # 布局：水平布局 (左图标 | 右文字)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(15, 10, 15, 10)
        layout.setSpacing(15)

        # 左侧图标
        lbl_icon = QLabel(icon)
        lbl_icon.setStyleSheet("font-size: 32px; font-weight: bold; border: none; background: transparent;")
        lbl_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_icon)

        # 右侧文字容器
        right_container = QWidget()
        right_container.setStyleSheet("background: transparent; border: none;")
        right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(0, 5, 0, 5)
        right_layout.setSpacing(2)

        # 标题
        lbl_title = QLabel(title)
        lbl_title.setStyleSheet("font-size: 13px; opacity: 0.9; font-weight: bold;")

        # 数值
        self.lbl_value = QLabel(str(value))
        self.lbl_value.setStyleSheet("font-size: 26px; font-weight: bold;")

        right_layout.addWidget(lbl_title)
        right_layout.addWidget(self.lbl_value)
        layout.addWidget(right_container)

    def update_value(self, value):
        self.lbl_value.setText(str(value))


# ==================== UI组件：报告信息配置对话框 ====================
class ReportInfoDialog(QDialog):
    """独立的报告信息录入窗口 - 带自动模板匹配功能"""

    def __init__(self, parent, config, current_info):
        super().__init__(parent)
        self.setWindowTitle("配置报告基础信息 & 模板选择")
        self.resize(780, 500)
        self.config = config
        self.business_data = config.get("business_data", {})
        self.info = current_info.copy()

        layout = QVBoxLayout(self)

        group = QGroupBox("项目与报告配置")
        form = QGridLayout(group)
        form.setVerticalSpacing(15)
        form.setHorizontalSpacing(15)

        # === 控件定义 ===
        self.cbo_company = QComboBox()
        self.cbo_project = QComboBox()
        self.txt_unit = QLineEdit()

        # 1. 检查类型 (下拉框 + 可手输)
        self.cbo_check_type = QComboBox()
        self.cbo_check_type.setEditable(True)
        # 读取配置中的选项
        self.cbo_check_type.addItems(self.business_data.get("check_content_options", []))

        # 2. 模板选择 (下拉框)
        self.cbo_template = QComboBox()
        # 扫描文件
        self.found_templates = self._scan_docx_files()
        self.cbo_template.addItems(self.found_templates)

        self.txt_area = QLineEdit()
        self.txt_person = QLineEdit()
        self.txt_date = QLineEdit(datetime.now().strftime("%Y-%m-%d"))
        self.txt_deadline = QLineEdit()

        self.txt_overview = QPlainTextEdit()
        self.txt_overview.setPlaceholderText("选择项目后自动填充...")
        self.txt_overview.setMaximumHeight(70)

        # === 布局排版 ===
        # 第一行
        form.addWidget(QLabel("项目公司:"), 0, 0)
        form.addWidget(self.cbo_company, 0, 1)
        form.addWidget(QLabel("项目名称:"), 0, 2)
        form.addWidget(self.cbo_project, 0, 3)

        # 第二行
        form.addWidget(QLabel("被检单位:"), 1, 0)
        form.addWidget(self.txt_unit, 1, 1)
        form.addWidget(QLabel("检查类型:"), 1, 2)
        form.addWidget(self.cbo_check_type, 1, 3)

        # 第三行：核心修改 - 模板选择
        form.addWidget(QLabel("导出模板:"), 2, 0)
        form.addWidget(self.cbo_template, 2, 1)
        form.addWidget(QLabel("检查部位:"), 2, 2)
        form.addWidget(self.txt_area, 2, 3)

        # 第四行
        form.addWidget(QLabel("检查人员:"), 3, 0)
        form.addWidget(self.txt_person, 3, 1)
        form.addWidget(QLabel("检查日期:"), 3, 2)
        form.addWidget(self.txt_date, 3, 3)

        # 第五行：期限
        form.addWidget(QLabel("整改期限:"), 4, 0)
        # 期限快捷按钮容器
        deadline_widget = QWidget()
        h_dead = QHBoxLayout(deadline_widget)
        h_dead.setContentsMargins(0, 0, 0, 0)
        h_dead.addWidget(self.txt_deadline)
        btn_3d = QPushButton("+3天");
        btn_3d.setFixedWidth(50)
        btn_7d = QPushButton("+7天");
        btn_7d.setFixedWidth(50)
        btn_3d.clicked.connect(lambda: self._calc_deadline(3))
        btn_7d.clicked.connect(lambda: self._calc_deadline(7))
        h_dead.addWidget(btn_3d)
        h_dead.addWidget(btn_7d)
        form.addWidget(deadline_widget, 4, 1)

        # 第六行：概况
        form.addWidget(QLabel("项目概况:"), 5, 0)
        form.addWidget(self.txt_overview, 5, 1, 1, 3)

        layout.addWidget(group)

        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

        # === 初始化数据 ===
        self._init_data()

        # 信号连接
        self.cbo_company.currentTextChanged.connect(self._on_company_changed)
        self.cbo_project.currentTextChanged.connect(self._on_project_changed)
        # 关键：检查类型改变时，自动匹配模板
        self.cbo_check_type.currentTextChanged.connect(self._auto_match_template)

    def _scan_docx_files(self):
        """扫描当前目录下的 docx 文件"""
        import glob
        # 获取所有docx
        files = glob.glob("*.docx")
        # 过滤掉临时文件(~$开头)和生成的报告(检查报告_开头)
        valid_files = []
        for f in files:
            if f.startswith("~$"): continue
            if f.startswith("检查报告_"): continue
            valid_files.append(f)

        if not valid_files:
            return ["模板.docx (未找到文件)"]
        return valid_files

    def _init_data(self):
        companies = list(self.business_data.get("company_project_map", {}).keys())
        self.cbo_company.addItems(companies)

        # 回填缓存的数据
        if self.info.get("project_company"): self.cbo_company.setCurrentText(self.info["project_company"])
        if self.info.get("project_name"): self.cbo_project.setCurrentText(self.info["project_name"])

        self.txt_unit.setText(self.info.get("inspected_unit", ""))
        self.cbo_check_type.setEditText(self.info.get("check_content", "安全质量综合检查"))
        self.txt_area.setText(self.info.get("check_area", ""))
        self.txt_person.setText(self.info.get("check_person", ""))
        self.txt_date.setText(self.info.get("check_date", datetime.now().strftime("%Y-%m-%d")))
        self.txt_deadline.setText(self.info.get("rectification_deadline", ""))
        self.txt_overview.setPlainText(self.info.get("project_overview", ""))

        # 尝试选中上次的模板
        last_tpl = self.info.get("template_name", "")
        if last_tpl and last_tpl in self.found_templates:
            self.cbo_template.setCurrentText(last_tpl)
        else:
            # 如果没有历史记录，触发一次自动匹配
            self._auto_match_template(self.cbo_check_type.currentText())

        if self.cbo_project.currentText():
            self._on_project_changed(self.cbo_project.currentText())

    def _on_company_changed(self, company_name):
        """当公司改变时，联动项目列表和单位"""
        # 1. 暂时屏蔽信号，防止清空时触发多余操作
        self.cbo_project.blockSignals(True)
        self.cbo_project.clear()

        # 2. 重新填充项目
        projects = self.business_data.get("company_project_map", {}).get(company_name, [])
        self.cbo_project.addItems(projects)

        # 3. 恢复信号
        self.cbo_project.blockSignals(False)

        # 4. 更新被检单位
        unit = self.business_data.get("company_unit_map", {}).get(company_name, "")
        self.txt_unit.setText(unit)

        # 5. 🔴 核心修复：强制刷新项目概况
        # 只要项目列表不为空，就手动选中第一个，并强制调用更新函数
        if self.cbo_project.count() > 0:
            self.cbo_project.setCurrentIndex(0)
            # 手动调用，确保概况文本框更新
            self._on_project_changed(self.cbo_project.currentText())
        else:
            # 如果没有项目，清空概况
            self.txt_overview.clear()

    def _on_project_changed(self, project_name):
        overview = self.business_data.get("project_overview_map", {}).get(project_name, "")
        self.txt_overview.setPlainText(overview)

    def _auto_match_template(self, check_type_text):
        """
        根据检查类型关键字，智能匹配文件夹下的 .docx 模板
        支持你指定的 6 种标准检查类型
        """
        if not self.found_templates:
            return

        # 1. 定义关键词优先级映射
        # 字典顺序：{ 检查类型关键词 : 模板文件名应包含的词 }
        mapping = [
            ("复工", "复工"),  # 匹配 "复工安全质量检查" -> 找含 "复工" 的模板
            ("节前", "节前"),  # 匹配 "节前安全检查"
            ("整治", "整治"),  # 匹配 "专项整治检查"
            ("综合", "综合"),  # 匹配 "安全质量综合检查" -> 找含 "综合" 的模板
            ("工程质量", "质量"),  # 匹配 "工程质量专项检查"
            ("安全生产", "安全"),  # 匹配 "安全生产专项检查"
            ("质量", "质量"),  # 兜底
            ("安全", "安全")  # 兜底
        ]

        target_keyword = ""

        # 2. 遍历查找匹配的关键词
        for check_key, template_key in mapping:
            if check_key in check_type_text:
                target_keyword = template_key
                break

        # 3. 如果没找到特定关键词，默认为综合或通用
        if not target_keyword:
            target_keyword = "通用"

        # 4. 在文件列表中查找最佳匹配
        best_match = None

        # 优先找完全包含关键词的
        for tpl in self.found_templates:
            if target_keyword in tpl:
                best_match = tpl
                break

        # 如果没找到，且列表里有叫 "模板.docx" 的，就选它做备胎
        if not best_match and "模板.docx" in self.found_templates:
            best_match = "模板.docx"

        # 5. 执行选中
        if best_match:
            self.cbo_template.setCurrentText(best_match)
            # 可选：在控制台打印匹配结果方便调试
            print(f"检查类型: {check_type_text} -> 匹配关键词: {target_keyword} -> 选中模板: {best_match}")

    def _calc_deadline(self, days):
        try:
            base = datetime.strptime(self.txt_date.text(), "%Y-%m-%d")
            self.txt_deadline.setText((base + timedelta(days=days)).strftime("%Y-%m-%d"))
        except:
            pass

    def get_data(self):
        """返回数据"""
        tpl = self.cbo_template.currentText()
        if "(未找到文件)" in tpl: tpl = "模板.docx"  # 兜底

        return {
            "project_company": self.cbo_company.currentText(),
            "project_name": self.cbo_project.currentText(),
            "inspected_unit": self.txt_unit.text(),
            "check_content": self.cbo_check_type.currentText(),
            "template_name": tpl,  # 返回选中的模板
            "check_area": self.txt_area.text(),
            "check_person": self.txt_person.text(),
            "check_date": self.txt_date.text(),
            "rectification_deadline": self.txt_deadline.text(),
            "project_overview": self.txt_overview.toPlainText()
        }


# ==================== UI组件：带水波纹动画的按钮 ====================
class RippleButton(QPushButton):
    """带有水波纹点击动画和悬停效果的现代化按钮"""

    def __init__(self, text="", parent=None, color=THEME_COLORS["primary"]):
        super().__init__(text, parent)
        self.cursor_pos = QPointF()
        self.radius = 0
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        # 基础样式
        self.base_color = color
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: white;
                border: 1px solid #E0E0E0;
                border-radius: 4px;
                padding: 6px 12px;
                color: #333;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: #F5F5F5;
                border-color: {color};
                color: {color};
            }}
        """)

        # 动画设置
        self.animation = QPropertyAnimation(self, b"radius_prop")
        self.animation.setDuration(400)  # 动画时长
        self.animation.setEasingCurve(QEasingCurve.Type.OutQuad)

        self.radius = 0

    @pyqtProperty(float)
    def radius_prop(self):
        return self.radius

    @radius_prop.setter
    def radius_prop(self, val):
        self.radius = val
        self.update()

    def mousePressEvent(self, event):
        # 记录点击位置，开始动画
        self.cursor_pos = event.position()
        self.radius = 0
        self.animation.stop()
        # 计算最大半径（覆盖整个按钮）
        end_radius = max(self.width(), self.height()) * 1.5
        self.animation.setStartValue(0)
        self.animation.setEndValue(end_radius)
        self.animation.start()
        super().mousePressEvent(event)

    def paintEvent(self, event):
        super().paintEvent(event)
        # 绘制水波纹
        if self.radius > 0:
            painter = QPainter(self)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            # 设置裁剪区域，防止波纹画出按钮外面
            path = QPainterPath()
            path.addRoundedRect(QRectF(self.rect()), 4, 4)
            painter.setClipPath(path)

            # 半透明波纹颜色
            brush = QBrush(QColor(self.base_color))
            color = brush.color()
            color.setAlpha(40)  # 透明度
            painter.setBrush(color)
            painter.setPen(Qt.PenStyle.NoPen)

            painter.drawEllipse(self.cursor_pos, self.radius, self.radius)


# ==================== 主窗口（V4.0完整版）====================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        sys.excepthook = self._global_exception_handler

        self.config = ConfigManager.load()
        self.refresh_business_data()

        self.tasks = []
        self.current_task_id = None
        self.running_workers = {}
        self.pending_queue = []

        self.init_ui()
        self.setup_shortcuts()

    def _global_exception_handler(self, exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        error_msg = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        print(f"❌ 全局异常:\n{error_msg}")
        QMessageBox.critical(None, "程序错误", f"{exc_type.__name__}: {exc_value}")

    def refresh_business_data(self):
        self.business_data = self.config.get("business_data", DEFAULT_BUSINESS_DATA)

    def __init__(self):
        super().__init__()
        sys.excepthook = self._global_exception_handler

        self.config = ConfigManager.load()
        self.refresh_business_data()

        self.tasks = []
        self.current_task_id = None
        self.running_workers = {}
        self.pending_queue = []

        # === 新增：初始化默认报告信息 ===
        self.report_info = {
            "project_company": "",
            "project_name": "",
            "inspected_unit": "",
            "check_content": "安全质量综合检查",
            "template_name": "模板.docx",  # 🔴 新增默认值
            "check_area": "",
            "check_person": self.config.get("last_check_person", ""),
            "check_date": datetime.now().strftime("%Y-%m-%d"),
            "rectification_deadline": "",
            "project_overview": ""
        }

        self.init_ui()
        self.setup_shortcuts()

    def init_ui(self):
        self.setWindowTitle("普洱版纳区域质量安全检查助手V3.0专家版")
        self.resize(1450, 1000)

        # 设置应用样式
        self.setStyleSheet("""
            QMainWindow { background-color: #F5F5F5; }
            QToolBar { background: white; border-bottom: 1px solid #E0E0E0; spacing: 8px; padding: 8px; }
            QPushButton { padding: 6px 12px; border: 1px solid #E0E0E0; border-radius: 4px; background: white; }
            QPushButton:hover { background: #F5F5F5; }
        """)

        # 1. 工具栏
        toolbar = QToolBar("Main")
        toolbar.setMovable(False)
        toolbar.setIconSize(QSize(20, 20))
        self.addToolBar(toolbar)

        toolbar.addWidget(QLabel("  <b>检查模式</b>: "))
        self.cbo_prompt = QComboBox()
        self.cbo_prompt.setMinimumWidth(250)
        prompts = self.config.get("prompts", DEFAULT_PROMPTS_V4)
        self.cbo_prompt.addItems(prompts.keys())
        self.cbo_prompt.setCurrentText(self.config.get("last_prompt", list(prompts.keys())[0]))
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

        toolbar.addSeparator()
        # === 新增：报告信息配置按钮 ===
        self.act_info = QAction("📝 报告信息配置", self)
        self.act_info.setToolTip("配置公司、项目、人员等报告基础信息")
        toolbar.addAction(self.act_info)

        self.act_export = QAction("📄 导出报告", self)
        toolbar.addAction(self.act_export)

        empty = QWidget()
        empty.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        toolbar.addWidget(empty)
        self.act_info.triggered.connect(self.open_report_info_dialog)
        self.act_history = QAction("📚 历史", self)
        self.act_help = QAction("❓ 帮助", self)
        self.act_setting = QAction("⚙ 设置", self)
        toolbar.addAction(self.act_history)
        toolbar.addAction(self.act_help)
        toolbar.addAction(self.act_setting)

        # 2. 主布局区域
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(15, 15, 15, 15)

        # 3. 顶部统计面板（横向布局）
        stats_widget = QWidget()
        stats_widget.setStyleSheet("background: transparent;")
        stats_widget.setFixedHeight(90)
        stats_layout = QHBoxLayout(stats_widget)
        stats_layout.setSpacing(10)  # 间距稍微调小一点以容纳更多卡片
        stats_layout.setContentsMargins(0, 0, 0, 5)

        # 定义5个卡片
        # 1. 严重安全 (红)
        self.card_severe_safety = StatsCard("严重安全隐患", 0, THEME_COLORS["severe_safety"], "🔴")
        # 2. 一般安全 (橙)
        self.card_general_safety = StatsCard("一般安全隐患", 0, THEME_COLORS["general_safety"], "🟠")
        # 3. 严重质量 (深橙/红橙) - 新增
        self.card_severe_quality = StatsCard("严重质量问题", 0, "#E64A19", "🚫")
        # 4. 一般质量 (黄) - 修改原质量卡片
        self.card_general_quality = StatsCard("一般质量问题", 0, THEME_COLORS["general_quality"], "🟡")
        # 5. 已检查 (蓝)
        self.card_checked = StatsCard("检查图像数量", "0/0", THEME_COLORS["info"], "📸")

        # 将卡片加入布局
        stats_layout.addWidget(self.card_severe_safety)
        stats_layout.addWidget(self.card_general_safety)
        stats_layout.addWidget(self.card_severe_quality)
        stats_layout.addWidget(self.card_general_quality)
        stats_layout.addWidget(self.card_checked)
        stats_layout.addStretch()

        main_layout.addWidget(stats_widget)

        # 4. 核心内容区（左右分割）
        splitter_main = QSplitter(Qt.Orientation.Horizontal)

        # --- 4.1 左侧：图片列表 ---
        left_widget = QWidget()
        left_widget.setStyleSheet("background: white; border-radius: 8px;")
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(10, 10, 10, 10)

        self.lbl_count = QLabel(f"<b>待检队列</b> (0/{MAX_IMAGES})")
        self.lbl_count.setStyleSheet("font-size: 14px; color: #212121;")

        self.list_widget = QListWidget()
        self.list_widget.setStyleSheet("""
             QListWidget { border: none; font-size: 13px; }
             QListWidget::item { padding: 8px; border-bottom: 1px solid #F0F0F0; }
             QListWidget::item:selected { background: #E3F2FD; color: #1976D2; }
             QListWidget::item[data-status="done"] { color: #4CAF50; }
        """)

        btn_layout = QHBoxLayout()
        self.btn_retry = QPushButton("🔄 重试失败")
        self.btn_retry.clicked.connect(self.retry_errors)
        btn_layout.addWidget(self.btn_retry)
        btn_layout.addStretch()

        left_layout.addWidget(self.lbl_count)
        left_layout.addWidget(self.list_widget)
        left_layout.addLayout(btn_layout)

        # --- 4.2 右侧：垂直分割 (图片在上，结果在下) ---
        splitter_right_vertical = QSplitter(Qt.Orientation.Vertical)
        splitter_right_vertical.setHandleWidth(8)
        splitter_right_vertical.setStyleSheet("""
            QSplitter::handle:vertical {
                background: #E0E0E0;
                height: 4px;
                margin: 4px 0px;
                border-radius: 2px;
            }
            QSplitter::handle:vertical:hover { background: #BDBDBD; }
        """)

        # [右上] 图片 + 工具栏
        top_container = QWidget()
        top_container.setStyleSheet("background: white; border-radius: 8px;")
        top_layout = QVBoxLayout(top_container)
        top_layout.setContentsMargins(10, 10, 10, 10)

        # 标注工具栏
        tool_widget = QWidget()
        tool_widget.setStyleSheet("background: #FAFAFA; border-radius: 4px; padding: 4px;")
        tool_layout = QHBoxLayout(tool_widget)
        tool_layout.setContentsMargins(5, 5, 5, 5)

        tool_layout.addWidget(QLabel("<b>标注工具:</b>"))
        # 🔴 移除缩放按钮（用鼠标滚轮缩放）
        # self.btn_tool_none = RippleButton("🖱️ 缩放")
        self.btn_tool_rect = RippleButton("⬜ 框")
        self.btn_tool_text = RippleButton("📝 文字")
        self.btn_tool_tag = RippleButton("🏷️ 引用问题", color=THEME_COLORS["primary"])
        self.btn_del_sel = RippleButton("❌ 删除选中", color=THEME_COLORS["danger"])

        # 样式微调：特殊按钮给不同颜色
        self.btn_auto = RippleButton("🤖 自动标识", color=THEME_COLORS["success"])
        self.btn_save = RippleButton("💾 保存截图", color=THEME_COLORS["info"])
        self.btn_clear_anno = RippleButton("🗑️ 清空所有", color=THEME_COLORS["secondary"])

        for btn in [self.btn_tool_rect, self.btn_tool_text]:
            btn.setFixedWidth(70)
        self.btn_tool_tag.setFixedWidth(90)
        self.btn_auto.setStyleSheet("background: #4CAF50; color: white; font-weight: bold;")

        # tool_layout.addWidget(self.btn_tool_none)  # 🔴 已移除
        tool_layout.addWidget(self.btn_tool_rect)
        tool_layout.addWidget(self.btn_tool_text)
        tool_layout.addWidget(self.btn_tool_tag)
        tool_layout.addWidget(self.btn_del_sel)

        # 🔴 最终简化：只保留两个调整按钮
        self.btn_change_color = RippleButton("🎨 调整文字颜色", color=THEME_COLORS["info"])
        self.btn_change_color.setFixedWidth(180)
        self.btn_change_color.setToolTip("调整选中文字的颜色")
        tool_layout.addWidget(self.btn_change_color)

        self.btn_resize_text = RippleButton("📏 调整文字大小", color=THEME_COLORS["info"])
        self.btn_resize_text.setFixedWidth(180)
        self.btn_resize_text.setToolTip("调整选中文字的字号")
        tool_layout.addWidget(self.btn_resize_text)

        tool_layout.addStretch()
        self.btn_resize_text.setFixedWidth(180)
        self.btn_resize_text.setToolTip("调整选中文字的字号")
        tool_layout.addWidget(self.btn_resize_text)

        tool_layout.addStretch()
        tool_layout.addWidget(self.btn_auto)
        tool_layout.addWidget(self.btn_save)
        tool_layout.addWidget(self.btn_clear_anno)

        top_layout.addWidget(tool_widget)

        # 图片显示区
        self.image_view = AnnotatableImageView()
        self.image_view.setStyleSheet("border: 1px solid #E0E0E0; background: #333; border-radius: 4px;")
        top_layout.addWidget(self.image_view)

        # [右下] 结果列表
        bottom_container = QWidget()
        bottom_container.setStyleSheet("background: white; border-radius: 8px;")
        bottom_layout = QVBoxLayout(bottom_container)
        bottom_layout.setContentsMargins(10, 5, 10, 5)

        result_label = QLabel("<b>识别结果</b>")
        result_label.setStyleSheet("font-size: 13px; color: #616161; margin-bottom: 5px;")
        bottom_layout.addWidget(result_label)

        self.result_container = QWidget()
        self.result_layout = QVBoxLayout(self.result_container)
        self.result_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.result_layout.setSpacing(6)
        self.result_layout.setContentsMargins(4, 4, 4, 4)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self.result_container)
        scroll.setStyleSheet("""
            QScrollArea { border: 1px solid #EEEEEE; border-radius: 4px; background: #FAFAFA; }
            QScrollBar:vertical { width: 10px; }
        """)
        bottom_layout.addWidget(scroll)

        # 组装右侧垂直分割
        splitter_right_vertical.addWidget(top_container)
        splitter_right_vertical.addWidget(bottom_container)
        # 初始比例：图片区800px，结果区250px (约3条记录高度)
        splitter_right_vertical.setSizes([800, 250])
        splitter_right_vertical.setStretchFactor(0, 1)  # 窗口拉伸时，主要拉伸图片区

        # 组装主分割
        splitter_main.addWidget(left_widget)
        splitter_main.addWidget(splitter_right_vertical)
        splitter_main.setSizes([280, 1200])  # 左列表窄，右侧宽

        main_layout.addWidget(splitter_main)
        # === 新增：底部日志控制台 ===
        log_group = QGroupBox("运行日志")
        log_group.setFixedHeight(150)  # 固定高度，不占用太多空间
        log_layout = QVBoxLayout(log_group)
        log_layout.setContentsMargins(5, 5, 5, 5)

        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)  # 只读
        self.txt_log.setStyleSheet("""
                    QTextEdit {
                        background-color: #2b2b2b;
                        color: #e0e0e0;
                        font-family: Consolas, Monaco, monospace;
                        font-size: 12px;
                        border: none;
                        border-radius: 4px;
                    }
                """)
        log_layout.addWidget(self.txt_log)

        main_layout.addWidget(log_group)
        # ================================
        # 5. 状态栏
        self.status_bar = self.statusBar()
        self.status_bar.setStyleSheet("background: white; color: #757575;")
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedWidth(200)
        self.status_bar.addPermanentWidget(self.progress_bar)

        # 6. 信号连接
        self.act_add.triggered.connect(self.add_files)
        self.act_run.triggered.connect(self.start_analysis)
        self.act_pause.triggered.connect(self.pause_analysis)
        self.act_clear.triggered.connect(self.clear_queue)
        self.act_export.triggered.connect(self.export_word)
        self.act_history.triggered.connect(self.show_history)
        self.act_setting.triggered.connect(self.open_settings)
        self.act_help.triggered.connect(self.show_help)

        self.cbo_prompt.currentTextChanged.connect(self.save_prompt_selection)
        self.list_widget.itemClicked.connect(self.on_item_clicked)

        # self.btn_tool_none.clicked.connect(lambda: self.image_view.set_tool("none"))  # 🔴 已移除
        self.btn_tool_rect.clicked.connect(lambda: self.image_view.set_tool("rect"))
        self.btn_tool_text.clicked.connect(lambda: self.image_view.set_tool("text"))
        self.btn_tool_tag.clicked.connect(lambda: self.image_view.set_tool("issue_tag"))

        self.btn_auto.clicked.connect(self.auto_annotate_current)
        self.btn_save.clicked.connect(self.save_marked_image)
        self.btn_clear_anno.clicked.connect(self.image_view.clear_annotations)
        self.btn_change_color.clicked.connect(self.change_selected_text_color)  # 🔴 调整颜色
        self.btn_resize_text.clicked.connect(self.resize_selected_text)  # 🔴 调整字号

        self.image_view.annotation_changed.connect(self.on_annotation_changed)
        self.image_view.tool_reset.connect(lambda: self.image_view.set_tool("none"))

    def keyPressEvent(self, event):
        # 监听 Delete 键
        if event.key() == Qt.Key.Key_Delete:
            self.image_view.delete_selected_items()
        else:
            super().keyPressEvent(event)

    def open_report_info_dialog(self):
        """打开报告信息配置窗口"""
        # 传入当前的 config 和 report_info
        dlg = ReportInfoDialog(self, self.config, self.report_info)

        if dlg.exec() == QDialog.DialogCode.Accepted:
            # 获取用户在对话框中填写的最新数据
            self.report_info = dlg.get_data()

            # 保存检查人到配置文件，方便下次自动填充
            if self.report_info.get("check_person"):
                self.config["last_check_person"] = self.report_info["check_person"]
                ConfigManager.save(self.config)

            # 在状态栏显示提示
            p_name = self.report_info.get('project_name', '未命名项目')
            self.status_bar.showMessage(f"✅ 报告信息已更新: {p_name}", 3000)

    def setup_shortcuts(self):
        """设置快捷键"""
        QAction("添加", self, shortcut=QKeySequence("Ctrl+O"), triggered=self.add_files)
        QAction("分析", self, shortcut=QKeySequence("F5"), triggered=self.start_analysis)
        QAction("导出", self, shortcut=QKeySequence("Ctrl+E"), triggered=lambda: self.export_word("模板.docx"))

    def save_prompt_selection(self, text):
        if text:
            self.config["last_prompt"] = text
            ConfigManager.save(self.config)

    def update_stats(self):
        """更新统计卡片数值"""
        stats = StatsManager.analyze_tasks(self.tasks)

        # 更新5个卡片的数值
        self.card_severe_safety.update_value(stats["severe_safety"])
        self.card_general_safety.update_value(stats["general_safety"])

        # 新增：单独显示严重质量缺陷
        self.card_severe_quality.update_value(stats["severe_quality"])

        # 修改：显示一般质量缺陷
        self.card_general_quality.update_value(stats["general_quality"])

        # 进度
        self.card_checked.update_value(f"{stats['analyzed_images']}/{stats['total_images']}")

    # ==================== 日志辅助方法 ====================
    def log(self, message, level="info"):

        timestamp = datetime.now().strftime("[%H:%M:%S]")

        if level == "error":
            color = "#FF5252"  # 红
            icon = "❌"
        elif level == "warning":
            color = "#FFD740"  # 黄
            icon = "⚠️"
        elif level == "success":
            color = "#69F0AE"  # 亮绿
            icon = "✅"
        else:
            color = "#E0E0E0"  # 默认白
            icon = "ℹ️"

        # 使用 HTML 格式化颜色
        html = f'<span style="color:#808080">{timestamp}</span> <span style="color:{color}">{icon} {message}</span>'
        self.txt_log.append(html)

        # 自动滚动到底部
        scrollbar = self.txt_log.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

        # 同时更新状态栏（可选）
        self.status_bar.showMessage(f"{icon} {message}", 3000)

    def add_files(self):
        current_count = len(self.tasks)
        if current_count >= MAX_IMAGES:
            QMessageBox.warning(self, "数量限制", f"单次最多 {MAX_IMAGES} 张")
            return

        remaining = MAX_IMAGES - current_count
        paths, _ = QFileDialog.getOpenFileNames(
            self, f"选择图片 (还能选{remaining}张)", "",
            "Images (*.jpg *.png *.jpeg)"
        )

        if not paths:
            return

        if len(paths) > remaining:
            QMessageBox.warning(self, "超限", f"截取前{remaining}张")
            paths = paths[:remaining]

        for path in paths:
            task_id = str(time.time()) + os.path.basename(path)
            task = {
                "id": task_id, "path": path, "name": os.path.basename(path),
                "status": "waiting", "issues": [], "edited_issues": None,
                "error": None, "elapsed_sec": None, "meta": {},
                "annotations": [], "export_image_path": None
            }
            self.tasks.append(task)

            item = QListWidgetItem(f"📷 {os.path.basename(path)}")
            item.setData(Qt.ItemDataRole.UserRole, task_id)
            self.list_widget.addItem(item)

        self.lbl_count.setText(f"<b>待检队列</b> ({len(self.tasks)}/{MAX_IMAGES})")
        self.update_stats()

    def start_analysis(self):
        if not self.config.get("api_key"):
            QMessageBox.warning(self, "配置缺失", "请先在设置中配置API Key")
            return

        waiting = [t for t in self.tasks if t['status'] in ['waiting', 'error']]
        if not waiting:
            self.status_bar.showMessage("没有待处理任务")
            return

        for t in waiting:
            if t["id"] not in self.pending_queue and t["id"] not in self.running_workers:
                self.pending_queue.append(t["id"])
                t["status"] = "queued"

        self.progress_bar.setVisible(True)
        self._kick_scheduler()

    def pause_analysis(self):
        self.pending_queue.clear()
        for t in self.tasks:
            if t["status"] == "queued":
                t["status"] = "waiting"
        self.status_bar.showMessage("已暂停")

    def _kick_scheduler(self):
        max_conc = int(self.config.get("max_concurrency", 2))

        while len(self.running_workers) < max_conc and self.pending_queue:
            task_id = self.pending_queue.pop(0)
            task = next((t for t in self.tasks if t['id'] == task_id), None)
            if not task:
                continue

            task["status"] = "analyzing"
            self.update_list_status(task_id, "⏳")

            worker = AnalysisWorker(task, self.config, "")
            worker.result_ready.connect(self.on_worker_done)
            worker.log_signal.connect(self.log)

            self.running_workers[task_id] = worker
            worker.start()

        total = len([t for t in self.tasks if t["status"] in ["queued", "analyzing", "done"]])
        done = len([t for t in self.tasks if t["status"] == "done"])
        if total > 0:
            self.progress_bar.setValue(int(done / total * 100))

        if not self.running_workers and not self.pending_queue:
            self.status_bar.showMessage("✅ 分析完成")
            self.progress_bar.setValue(100)
            QTimer.singleShot(2000, lambda: self.progress_bar.setVisible(False))
            if self.config.get("auto_save_history"):
                self.save_to_history()

    def on_worker_done(self, task_id, result):
        task = next((t for t in self.tasks if t['id'] == task_id), None)
        if task:
            task["elapsed_sec"] = result.get("elapsed_sec")

            if result.get("ok"):
                task['status'] = 'done'
                task['issues'] = result.get("issues", [])
                task["error"] = None

                # 根据严重程度显示图标
                severe_count = sum(1 for i in task['issues'] if "严重" in i.get("risk_level", ""))
                if severe_count > 0:
                    self.update_list_status(task_id, "🔴")
                elif len(task['issues']) > 0:
                    self.update_list_status(task_id, "🟡")
                else:
                    self.update_list_status(task_id, "✅")
            else:
                task['status'] = 'error'
                task['issues'] = []
                task["error"] = result.get("error")
                self.update_list_status(task_id, "❌")

            if self.current_task_id == task_id:
                QTimer.singleShot(50, lambda: self.render_result(task))

        if task_id in self.running_workers:
            worker = self.running_workers.pop(task_id)
            try:
                worker.result_ready.disconnect()
            except:
                pass
            worker.quit()
            worker.wait(1000)
            worker.deleteLater()

        self.update_stats()
        self._kick_scheduler()
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == task_id:
                task = next((t for t in self.tasks if t['id'] == task_id), None)
                if task:
                    item.setText(f"✅ {task['name']}")
                    # 🔴 在这里添加：设置文字颜色为绿色
                    item.setForeground(QColor("#4CAF50"))  # 绿色
            # 更新列表显示


    def on_item_clicked(self, item):
        self.current_task_id = item.data(Qt.ItemDataRole.UserRole)
        task = next((t for t in self.tasks if t['id'] == self.current_task_id), None)
        if task:
            self.render_result(task)

    def render_result(self, task):
        """渲染结果（V4.0优化）"""
        # 清空旧内容
        widgets_to_delete = []
        while self.result_layout.count():
            item = self.result_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widgets_to_delete.append(widget)

        for widget in widgets_to_delete:
            try:
                widget.blockSignals(True)
                if isinstance(widget, ModernRiskCard):
                    widget.edit_requested.disconnect()
                    widget.delete_requested.disconnect()
            except:
                pass
            widget.hide()
            widget.setParent(None)
            widget.deleteLater()

        QApplication.processEvents()

        # 显示图片
        img_path = task.get("path", "")
        if img_path and os.path.exists(img_path):
            if self.image_view._img_path != img_path:
                self.image_view.set_image(img_path)

        # 更新标注
        issues = task.get("edited_issues") if task.get("edited_issues") is not None else task.get("issues", [])
        self.image_view.set_ai_issues(issues)
        self.image_view.set_user_annotations(task.get("annotations", []) or [])

        # 显示问题卡片
        if task['status'] == 'done':
            if issues:
                for item in issues:
                    card = ModernRiskCard(item)
                    card.edit_requested.connect(self.edit_issue)
                    card.delete_requested.connect(self.delete_issue)
                    self.result_layout.addWidget(card)
            else:
                lbl_empty = QLabel("✅ 未发现明显安全隐患或质量问题")
                lbl_empty.setStyleSheet("""
                    font-size: 14px; 
                    color: #4CAF50; 
                    padding: 20px;
                    background: #E8F5E9;
                    border-radius: 8px;
                    text-align: center;
                """)
                lbl_empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
                self.result_layout.addWidget(lbl_empty)
        elif task['status'] == 'analyzing':
            lbl_loading = QLabel("⏳ 正在分析中...")
            lbl_loading.setStyleSheet("font-size: 14px; color: #757575; padding: 20px;")
            self.result_layout.addWidget(lbl_loading)
        elif task['status'] == 'error':
            lbl_error = QLabel(f"❌ 分析失败: {task.get('error', '未知错误')}")
            lbl_error.setStyleSheet("font-size: 14px; color: #F44336; padding: 20px;")
            lbl_error.setWordWrap(True)
            self.result_layout.addWidget(lbl_error)

    def edit_issue(self, item):
        task = self._current_task()
        if not task or task.get("status") != "done":
            return

        issues = task.get("edited_issues") if task.get("edited_issues") is not None else task.get("issues", [])

        dlg = IssueEditDialog(self, item)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            new_item = dlg.get_value()

            if task.get("edited_issues") is None:
                task["edited_issues"] = [dict(x) for x in issues]

            for i, x in enumerate(task["edited_issues"]):
                if id(x) == id(item):
                    task["edited_issues"][i] = new_item
                    break

            task["export_image_path"] = None
            QTimer.singleShot(100, lambda: self.render_result(task))

    def delete_issue(self, item):
        sender_card = self.sender()
        if sender_card and isinstance(sender_card, ModernRiskCard):
            try:
                sender_card.blockSignals(True)
                sender_card.edit_requested.disconnect()
                sender_card.delete_requested.disconnect()
            except:
                pass

        task = self._current_task()
        if task:
            issues = task.get("edited_issues") if task.get("edited_issues") is not None else task.get("issues", [])
            if task.get("edited_issues") is None:
                task["edited_issues"] = [dict(x) for x in issues]

            task["edited_issues"] = [x for x in task["edited_issues"] if id(x) != id(item)]
            task["export_image_path"] = None

            self.image_view.set_ai_issues(task["edited_issues"])
            QTimer.singleShot(150, lambda: self.render_result(task))
            self.update_stats()
            self.status_bar.showMessage("已删除该问题", 2000)

    def on_annotation_changed(self):
        task = self._current_task()
        if task:
            task["annotations"] = self.image_view.get_user_annotations()

    def auto_annotate_current(self):
        task = self._current_task()
        if not task or task.get("status") != "done":
            QMessageBox.warning(self, "提示", "请先完成AI分析")
            return

        issues = task.get("edited_issues") if task.get("edited_issues") is not None else task.get("issues", [])
        if not issues:
            QMessageBox.information(self, "提示", "未检测到问题")
            return

        count = 0
        for idx, item in enumerate(issues, 1):
            bbox = item.get("bbox")
            if bbox and len(bbox) == 4:
                cx = (bbox[0] + bbox[2]) / 2
                cy = (bbox[1] + bbox[3]) / 2

                # 获取问题描述
                text = item.get("issue", "")

                # 🔴 重要：先定义 level 变量
                level = item.get("risk_level", "")

                # 🔴 精简处理（一步到位）
                import re

                # 1. 移除【专业名称】
                text = re.sub(r'【[^】]+】', '', text)

                # 2. 移除常见词组（一次性处理）
                remove_words = [
                    "存在", "发现", "有", "未", "没有", "缺少", "应", "需要",
                    "的问题", "的情况", "的现象", "问题", "情况", "现象"
                ]
                for word in remove_words:
                    text = text.replace(word, "")

                # 3. 清理空格和标点
                text = re.sub(r'[，。、；：！？\s]+', '', text)

                # 4. 智能截取（保留关键信息）
                if len(text) > 10:
                    # 查找关键词位置，在其后截断
                    keywords = ["不符", "不足", "不当", "未接", "未设", "缺失", "松动", "破损"]
                    for kw in keywords:
                        if kw in text:
                            pos = text.find(kw) + len(kw)
                            if 6 <= pos <= 12:
                                text = text[:pos]
                                break

                    # 如果仍然太长，直接取前10字
                    if len(text) > 10:
                        text = text[:10]

                 # 格式化显示
                text = f"{idx}.{text}"

                # 根据问题等级选择颜色
                if "严重安全" in level:
                    color = THEME_COLORS["severe_safety"]
                elif "一般安全" in level:
                    color = THEME_COLORS["general_safety"]
                elif "严重质量" in level:
                    color = THEME_COLORS["severe_quality"]
                else:
                    color = THEME_COLORS["general_quality"]

                # 使用固定默认字号
                font_size = 32

                new_anno = {"type": "text", "pos": [int(cx), int(cy)], "text": text,
                            "color": color, "width": 4, "font_size": font_size}
                self.image_view._create_item_from_data(new_anno)
                count += 1

        if count > 0:
            task["annotations"] = self.image_view.get_user_annotations()
            self.status_bar.showMessage(f"✅ 成功自动标识{count}处", 3000)

    def save_marked_image(self):
        task = self._current_task()
        if not task: return
        if not os.path.exists(task.get("path", "")):
            QMessageBox.warning(self, "失败", "图片不存在")
            return

        issues = task.get("edited_issues") if task.get("edited_issues") is not None else task.get("issues", [])
        anns = task.get("annotations", []) or []

        ensure_export_dir()
        base_name = os.path.splitext(os.path.basename(task["path"]))[0]
        out_path = os.path.join(EXPORT_IMG_DIR, f"{base_name}_{task['id'][-6:]}.png")

        # 🔴 修正点：将 build_export_marked_image 改为 export_marked_image
        ok = export_marked_image(task["path"], issues, anns, out_path)

        if ok:
            task["export_image_path"] = out_path
            self.status_bar.showMessage(f"✅ 已保存: {out_path}", 3000)
        else:
            QMessageBox.warning(self, "失败", "生成失败")

    def change_selected_text_color(self):
        """调整选中文字的颜色"""
        # 获取选中的图形项
        selected_items = self.image_view.scene().selectedItems()

        text_items = [item for item in selected_items if isinstance(item, (EditableTextItem, QGraphicsTextItem))]

        if not text_items:
            QMessageBox.information(self, "提示",
                                    "请先选中要调整颜色的文字标注\n\n💡 提示：点击文字可选中，按住Ctrl可多选")
            return

        # 🔴 简化修改：弹出颜色选择对话框
        color_dialog = QDialog(self)
        color_dialog.setWindowTitle("选择颜色")
        color_dialog.resize(350, 200)

        layout = QVBoxLayout(color_dialog)
        layout.setContentsMargins(15, 15, 15, 15)

        # 提示信息
        info_label = QLabel(f"<b>已选中 {len(text_items)} 个文字标注</b>")
        info_label.setStyleSheet("color: #000000; font-size: 14px; margin-bottom: 10px;")
        layout.addWidget(info_label)

        # 颜色选择
        color_label = QLabel("选择新颜色:")
        color_label.setStyleSheet("color: #000000; font-size: 13px; font-weight: bold; margin-bottom: 5px;")
        layout.addWidget(color_label)

        color_combo = QComboBox()
        color_combo.setStyleSheet("""
            QComboBox {
                font-size: 13px;
                padding: 5px;
                border: 2px solid #BDBDBD;
                border-radius: 4px;
            }
        """)
        colors = [
            ("🔴 红色", "#FF0000"),
            ("🟠 橙色", "#FF9800"),
            ("🟡 黄色", "#FFC107"),
            ("🟢 绿色", "#4CAF50"),
            ("🔵 蓝色", "#2196F3"),
            ("🟣 紫色", "#9C27B0"),
            ("⚫ 黑色", "#000000"),
            ("⚪ 白色", "#FFFFFF"),
        ]

        for name, code in colors:
            color_combo.addItem(name, code)

        layout.addWidget(color_combo)

        # 按钮
        btn_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        btn_box.setStyleSheet("""
            QPushButton {
                background: #2196F3;
                color: white;
                font-size: 13px;
                font-weight: bold;
                padding: 6px 20px;
                border: none;
                border-radius: 4px;
                min-width: 80px;
                margin-top: 10px;
            }
            QPushButton:hover {
                background: #1976D2;
            }
            QPushButton:pressed {
                background: #1565C0;
            }
        """)
        btn_box.accepted.connect(color_dialog.accept)
        btn_box.rejected.connect(color_dialog.reject)
        layout.addWidget(btn_box)

        if color_dialog.exec() != QDialog.DialogCode.Accepted:
            return

        new_color_code = color_combo.currentData()
        color_name = color_combo.currentText()
        new_color = QColor(new_color_code)

        # 更换所有选中文字的颜色
        for item in text_items:
            item.setDefaultTextColor(new_color)

            # 更新item的数据
            data = item.data(Qt.ItemDataRole.UserRole)
            if data:
                data["color"] = new_color_code
                item.setData(Qt.ItemDataRole.UserRole, data)

            # 强制刷新
            if hasattr(item, 'update'):
                item.update()

        # 更新标注数据
        task = self._current_task()
        if task:
            task["annotations"] = self.image_view.get_user_annotations()

        self.status_bar.showMessage(f"✅ 已调整 {len(text_items)} 个文字的颜色为{color_name}", 3000)
        self.image_view.annotation_changed.emit()

    def resize_selected_text(self):
        """调整选中文字的字号"""
        # 获取选中的图形项
        selected_items = self.image_view.scene().selectedItems()

        text_items = [item for item in selected_items if isinstance(item, (EditableTextItem, QGraphicsTextItem))]

        if not text_items:
            QMessageBox.information(self, "提示",
                                    "请先选中要调整字号的文字标注\n\n💡 提示：点击文字可选中，按住Ctrl可多选")
            return

        # 🔴 简化修改：弹出字号输入对话框
        font_dialog = QDialog(self)
        font_dialog.setWindowTitle("调整字号")
        font_dialog.resize(350, 200)

        layout = QVBoxLayout(font_dialog)
        layout.setContentsMargins(15, 15, 15, 15)

        # 提示信息
        info_label = QLabel(f"<b>已选中 {len(text_items)} 个文字标注</b>")
        info_label.setStyleSheet("color: #000000; font-size: 14px; margin-bottom: 10px;")
        layout.addWidget(info_label)

        # 字号选择
        size_label = QLabel("选择新字号:")
        size_label.setStyleSheet("color: #000000; font-size: 13px; font-weight: bold; margin-bottom: 5px;")
        layout.addWidget(size_label)

        size_spin = QSpinBox()
        size_spin.setRange(10, 60)
        size_spin.setValue(32)
        size_spin.setSuffix("px")
        size_spin.setStyleSheet("""
            QSpinBox {
                font-size: 13px;
                padding: 5px;
                border: 2px solid #BDBDBD;
                border-radius: 4px;
            }
        """)
        layout.addWidget(size_spin)

        # 按钮
        btn_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        btn_box.setStyleSheet("""
            QPushButton {
                background: #2196F3;
                color: white;
                font-size: 13px;
                font-weight: bold;
                padding: 6px 20px;
                border: none;
                border-radius: 4px;
                min-width: 80px;
                margin-top: 10px;
            }
            QPushButton:hover {
                background: #1976D2;
            }
            QPushButton:pressed {
                background: #1565C0;
            }
        """)
        btn_box.accepted.connect(font_dialog.accept)
        btn_box.rejected.connect(font_dialog.reject)
        layout.addWidget(btn_box)

        if font_dialog.exec() != QDialog.DialogCode.Accepted:
            return

        new_font_size = size_spin.value()

        # 更换所有选中文字的字号
        for item in text_items:
            font = item.font()
            font.setPointSize(new_font_size)
            item.setFont(font)

            # 更新item的数据
            data = item.data(Qt.ItemDataRole.UserRole)
            if data:
                data["font_size"] = new_font_size
                item.setData(Qt.ItemDataRole.UserRole, data)

            # 强制刷新
            if hasattr(item, 'update'):
                item.update()

        # 更新标注数据
        task = self._current_task()
        if task:
            task["annotations"] = self.image_view.get_user_annotations()

        self.status_bar.showMessage(f"✅ 已调整 {len(text_items)} 个文字的字号为 {new_font_size}px", 3000)
        self.image_view.annotation_changed.emit()

    def export_word(self):
        ensure_export_dir()

        valid_tasks = [t for t in self.tasks if t.get("status") == "done" or t.get("annotations")]
        if not valid_tasks:
            QMessageBox.warning(self, "提示", "没有可导出的任务。")
            return

        abs_export_dir = os.path.abspath(EXPORT_IMG_DIR)
        for t in self.tasks:
            if t not in valid_tasks: continue
            if not os.path.exists(t.get("path", "")): continue

            issues = t.get("edited_issues") if t.get("edited_issues") is not None else t.get("issues", [])
            anns = t.get("annotations", []) or []

            base_name = os.path.splitext(os.path.basename(t["path"]))[0]
            safe_name = "".join(c for c in base_name if c.isalnum() or c in (' ', '_', '-'))
            out_path = os.path.join(abs_export_dir, f"{safe_name}_{t['id'][-6:]}.png")

            # 🔴 修正点：将 build_export_marked_image 改为 export_marked_image
            if export_marked_image(t["path"], issues, anns, out_path):
                t["export_image_path"] = out_path

        # 统计数据
        stats = StatsManager.analyze_tasks(self.tasks)
        final_info = self.report_info.copy()

        target_template = final_info.get("template_name", "模板.docx")
        if not os.path.exists(target_template):
            pass

        # 合并统计数据
        final_info.update({
            "severe_safety": str(stats["severe_safety"]),
            "general_safety": str(stats["general_safety"]),
            "severe_quality": str(stats["severe_quality"]),
            "general_quality": str(stats["general_quality"]),
            "total_issues": str(stats["total_issues"])
        })

        if not final_info["project_name"]: final_info["project_name"] = "项目名称"

        # 保存文件
        default_name = f"检查报告_{final_info['project_name']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        path, _ = QFileDialog.getSaveFileName(self, "保存报告", default_name, "Word Files (*.docx)")

        if path:
            try:
                WordReportGenerator.generate(self.tasks, path, final_info, target_template)
                self.status_bar.showMessage(f"✅ 报告已生成: {path}", 5000)
            except Exception as e:
                import traceback
                QMessageBox.critical(self, "导出失败", f"生成失败:\n{e}\n{traceback.format_exc()}")

    def clear_queue(self):
        if self.running_workers:
            QMessageBox.warning(self, "警告", "任务正在分析中")
            return
        reply = QMessageBox.question(self, '确认', '确定清空所有任务吗？',
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.tasks.clear()
            self.list_widget.clear()
            self.lbl_count.setText(f"<b>待检队列</b> (0/{MAX_IMAGES})")
            self.update_stats()
            while self.result_layout.count():
                self.result_layout.takeAt(0).widget().deleteLater()

    def retry_errors(self):
        error_tasks = [t for t in self.tasks if t["status"] == "error"]
        if not error_tasks:
            self.status_bar.showMessage("没有失败任务")
            return
        for t in error_tasks:
            t["status"] = "waiting"
            t["error"] = None
        self.status_bar.showMessage(f"已重置{len(error_tasks)}个任务")
        self.start_analysis()

    def save_to_history(self):
        done_tasks = [t for t in self.tasks if t.get("status") == "done"]
        if not done_tasks:
            return
        stats = StatsManager.analyze_tasks(self.tasks)
        HistoryManager.add_record("项目", datetime.now().strftime("%Y-%m-%d"),
                                  self.config.get("last_check_person", ""), stats, done_tasks)

    def show_history(self):
        history = HistoryManager.load()
        records = history.get("inspections", [])
        if not records:
            QMessageBox.information(self, "历史记录", "暂无历史记录")
            return
        dlg = QDialog(self)
        dlg.setWindowTitle("检查历史")
        dlg.resize(800, 500)
        layout = QVBoxLayout(dlg)
        list_widget = QListWidget()
        for record in records:
            stats = record.get("stats", {})
            text = f"{record['date']} | {record['project']} | {record['person']} | " \
                   f"严重安全:{stats.get('severe_safety', 0)} 质量问题:{stats.get('severe_quality', 0) + stats.get('general_quality', 0)}"
            list_widget.addItem(text)
        layout.addWidget(list_widget)
        btn_close = QPushButton("关闭")
        btn_close.clicked.connect(dlg.accept)
        layout.addWidget(btn_close)
        dlg.exec()

    def open_settings(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("系统设置")
        dlg.resize(700, 400)
        layout = QVBoxLayout(dlg)
        form = QFormLayout()

        presets = self.config.get("provider_presets", DEFAULT_PROVIDERS)
        cbo_prov = QComboBox()
        cbo_prov.addItems(presets.keys())
        cbo_prov.setCurrentText(self.config.get("current_provider", list(presets.keys())[0]))

        txt_key = QLineEdit(self.config.get("api_key", ""))
        txt_key.setEchoMode(QLineEdit.EchoMode.Password)

        form.addRow("模型厂商:", cbo_prov)
        form.addRow("API Key:", txt_key)
        layout.addLayout(form)

        btn_save = QPushButton("保存配置")
        btn_save.setStyleSheet("background: #2196F3; color: white; font-weight: bold; padding: 10px;")

        def save_all():
            self.config["current_provider"] = cbo_prov.currentText()
            self.config["api_key"] = txt_key.text().strip()
            ConfigManager.save(self.config)
            dlg.accept()
            self.status_bar.showMessage("✅ 配置已保存", 3000)

        btn_save.clicked.connect(save_all)
        layout.addWidget(btn_save)
        dlg.exec()

    def show_help(self):
        help_text = """
<h3>V4.0 终极版 - 聚焦安全质量</h3>

<h4>核心特性</h4>
<ul>
<li><b>聚焦重点</b>: 专注安全隐患和质量问题，取消文明施工</li>
<li><b>智能分诊</b>: Router自动识别并指派2-4名专家</li>
<li><b>实时统计</b>: 顶部彩色卡片实时显示各类问题数量</li>
<li><b>分类显示</b>: 🔴严重安全 🟠一般安全 🟡质量问题</li>
</ul>

<h4>快捷键</h4>
<ul>
<li><b>Ctrl+O</b>: 添加图片</li>
<li><b>F5</b>: 开始分析</li>
<li><b>Ctrl+E</b>: 导出报告</li>
</ul>

<h4>操作流程</h4>
<ol>
<li>⚙ 设置 → 配置API Key</li>
<li>➕ 添加图片 → 选择施工现场照片</li>
<li>▶ 开始分析 → 观察顶部统计卡片更新</li>
<li>查看问题卡片 → 编辑/删除</li>
<li>🤖 自动标识 → 生成序号标注</li>
<li>📄 导出报告</li>
</ol>

<p><b>注意</b>: AI识别结果仅供参考，请人工复核。</p>
        """
        QMessageBox.information(self, "帮助", help_text)

    def update_list_status(self, task_id, icon):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == task_id:
                task = next((t for t in self.tasks if t['id'] == task_id), None)
                if task:
                    item.setText(f"{icon} {task['name']}")

                    # 🔴 新增：根据图标设置颜色
                    if icon == "✅":  # 完成 - 绿色
                        item.setForeground(QColor("#4CAF50"))
                    elif icon == "❌":  # 失败 - 红色
                        item.setForeground(QColor("#F44336"))
                    elif icon == "⏳":  # 处理中 - 蓝色
                        item.setForeground(QColor("#2196F3"))
                    else:  # 默认 - 黑色
                        item.setForeground(QColor("#212121"))

    def _current_task(self):
        if not self.current_task_id:
            return None
        return next((t for t in self.tasks if t['id'] == self.current_task_id), None)


# ==================== 主函数 ====================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
