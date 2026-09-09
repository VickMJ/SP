# -*- coding: utf-8 -*-

"""
======================================================================
StockPulse V7.8.1 Industry Behavior Engine：行业趋势 + 加速/减速 + 退潮 + 内部分化：三指数独立实时 + TDX原生市场广度 + 真正行业第二闸门
======================================================================

核心：
    TDX多周期MACD
    + VWAP
    + 动态量能/DAR
    + 上证Market Regime
    + 4/5星报警

本版本重点修复：
    1. 999999 / 399001 使用 get_index_bars()
    2. 上证月K使用 category=6
    3. 指数日K使用 category=4
    4. 月K排除当前未完成月K
    5. MACD增加金叉/死叉结构与事件识别
    6. 指数压力只使用前20个已完成交易日
    7. 突破前高时压力距离=0，不允许出现负值
    8. DAR动态量能不再把当前量计入基准
    9. RS统一使用同周期1M数据
   10. 修复评分过度饱和100分
   11. ★★★★★增加硬条件
   12. ★★★★ / ★★★★★继续语音播报
   13. ★★★★★继续人工确认
   14. 持仓移动止盈使用Excel中的“最高价”
   15. 默认非交易时间不运行实盘扫描

======================================================================
"""

import os
import sys
import time
from collections import deque
import logging
import datetime
import threading
import ctypes

import numpy as np
import pandas as pd

import pygetwindow as gw
import pyttsx3

from pytdx.hq import TdxHq_API
try:
    from pytdx.parser.get_block_info import get_and_parse_block_info
except Exception:
    get_and_parse_block_info = None

# V7.5.4：不再依赖 opentdx。
# 直接使用现有 pytdx 的 TDX行情连接读取原生统计指数 880005。
# 同时兼容某些 pytdx/衍生版本可能已经暴露的 index up_count/down_count 字段。
TDX_BREADTH_INDEX_CODE = "880005"
TDX_BREADTH_INDEX_MARKET = 1


# ======================================================================
# 日志
# ======================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

logger = logging.getLogger("StockPulse")


# ======================================================================
# 完整运行日志自动保存
# ----------------------------------------------------------------------
# 设计原则：
# 1. 控制台原样显示；
# 2. 同时把 print / traceback / logger 输出完整写入每日日志；
# 3. 每个交易日一个独立文件，便于后续上传给 ChatGPT 做全天复盘；
# 4. 默认保存到程序目录下 logs/StockPulse_YYYY-MM-DD.log；
# 5. 可通过环境变量 STOCKPULSE_LOG_DIR 指定日志目录。
# ======================================================================

STOCKPULSE_LOG_DIR = os.getenv(
    "STOCKPULSE_LOG_DIR",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
)

os.makedirs(STOCKPULSE_LOG_DIR, exist_ok=True)

# logger 自身也写入同一份文件。
# print/异常由 TeeStream 保存；logger 由此 FileHandler 保存，避免依赖
# logging.basicConfig 在 TeeStream 安装之前绑定了旧 stderr。
_STOCKPULSE_LOG_PATH = os.path.join(
    STOCKPULSE_LOG_DIR,
    f"StockPulse_{datetime.datetime.now().strftime('%Y-%m-%d')}.log"
)

if not any(
    isinstance(h, logging.FileHandler)
    and getattr(h, "baseFilename", "") == os.path.abspath(_STOCKPULSE_LOG_PATH)
    for h in logger.handlers
):
    _file_handler = logging.FileHandler(
        _STOCKPULSE_LOG_PATH,
        mode="a",
        encoding="utf-8"
    )
    _file_handler.setLevel(logging.INFO)
    _file_handler.setFormatter(
        logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    )
    logger.addHandler(_file_handler)


class TeeStream:
    """同时输出到原控制台和按日期切换的日志文件。"""

    def __init__(self, original_stream, stream_name="stdout"):
        self.original = original_stream
        self.stream_name = stream_name
        self._lock = threading.RLock()
        self._file = None
        self._date = None

    def _get_file(self):
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        if self._file is None or self._date != today:
            if self._file is not None:
                try:
                    self._file.flush()
                    self._file.close()
                except Exception:
                    pass
            path = os.path.join(STOCKPULSE_LOG_DIR, f"StockPulse_{today}.log")
            self._file = open(path, "a", encoding="utf-8", buffering=1)
            self._date = today
        return self._file

    def write(self, data):
        if data is None:
            return 0
        with self._lock:
            try:
                self.original.write(data)
            except Exception:
                pass
            try:
                f = self._get_file()
                f.write(data)
                f.flush()
            except Exception:
                pass
        return len(data)

    def flush(self):
        with self._lock:
            try:
                self.original.flush()
            except Exception:
                pass
            try:
                if self._file is not None:
                    self._file.flush()
            except Exception:
                pass

    def isatty(self):
        try:
            return self.original.isatty()
        except Exception:
            return False

    def fileno(self):
        return self.original.fileno()

    @property
    def encoding(self):
        return getattr(self.original, "encoding", "utf-8")


# 必须在主循环启动前安装，这样启动信息、异常 traceback、print 输出都能保存。
if not isinstance(sys.stdout, TeeStream):
    sys.stdout = TeeStream(sys.stdout, "stdout")
if not isinstance(sys.stderr, TeeStream):
    sys.stderr = TeeStream(sys.stderr, "stderr")

logger.info(
    "============================================================"
)
logger.info(
    f"StockPulse 完整运行日志已启用 | 目录={STOCKPULSE_LOG_DIR} | "
    f"文件=StockPulse_{datetime.datetime.now().strftime('%Y-%m-%d')}.log"
)


# ======================================================================
# 全局配置
# ======================================================================

EXCEL_PATH = "监测股票列表.xlsx"

INTERVAL_SECONDS = 30

# ======================================================================
# 个股趋势分层 V1
# ----------------------------------------------------------------------
# 大周期决定“能不能买”，小周期决定“什么时候买”。
# 使用：月K MA20 + 周K MA20/MA60 + 日K MA20/MA60。
# 只使用已经完成的月/周/日K，避免未来函数。
# 趋势层不替代原有评分，而是在原始评分上进行加减，并作为
# ENTRY_GATE V6.3 的第二层开仓闸门。
# ======================================================================
TREND_LAYER_ENABLED = True
TREND_FETCH_BARS = 120

# 趋势评分调整
TREND_SCORE_A = 12
TREND_SCORE_B = 6
TREND_SCORE_C = 0
TREND_SCORE_D = -8
TREND_SCORE_E = -15

# 趋势闸门：E/D禁止★★★★★；C只观察；A/B允许进入原有位置闸门
TREND_BLOCK_5STAR_GRADES = {"D", "E"}
TREND_WAIT_5STAR_GRADES = {"C"}

# MA位置容差，避免刚好贴线时频繁翻转
TREND_MA_TOLERANCE_PCT = 0.30


# ======================================================================
# ENTRY_GATE V6.2
# ----------------------------------------------------------------------
# 目标：不推翻原有评分/4星/5星逻辑，只在★★★★★真正触发前增加
# “这个位置现在到底能不能买”的第二道闸门。
#
# 核心原则：
# 1. 接近局部平台压力，不追。
# 2. 刚突破但尚未形成有效站稳，不追。
# 3. 突破后重新跌回平台，视为假突破，禁止★★★★★。
# 4. 大盘明显偏弱时，禁止★★★★★主动开仓。
# 5. ENTRY_GATE只影响★★★★★，不改变原有评分、不删除★★★★。
# ======================================================================
ENTRY_GATE_ENABLED = True

# 最近多少根1M K线寻找“局部平台压力”
ENTRY_GATE_LOOKBACK = 30

# 当前价格距离压力位小于该值：进入WAIT，不追涨；突破/站稳后再PASS
ENTRY_GATE_NEAR_PRESSURE_PCT = 0.80

# 突破压力至少超过该比例，才算“真正站上”
ENTRY_GATE_BREAKOUT_BUFFER_PCT = 0.15

# 有效突破要求最近两根已完成K线都收在压力上方
ENTRY_GATE_CONFIRM_BARS = 2

# 有效突破时，最近确认K线至少有一定量能
ENTRY_GATE_BREAKOUT_VOL_MIN = 1.20

# V6.2：保留短线失败确认窗口，但“当日突破失败”不再受8根K线限制。
# 8根仅用于快速状态参考；真正的当日失败从当天全部已完成K线寻找。
ENTRY_GATE_FAILURE_LOOKBACK = 8
ENTRY_GATE_FAILURE_BUFFER_PCT = 0.15

# 当日突破事件至少需要超过压力多少，才记录为“突破事件”。
ENTRY_GATE_DAY_BREAKOUT_BUFFER_PCT = 0.15

# 当日突破后，价格重新跌回突破压力下方多少，判定为失败。
ENTRY_GATE_DAY_FAILURE_RETRACE_PCT = 0.15

# 当日突破失败后，在重新站回压力之前，不允许★★★★★。
ENTRY_GATE_DAY_FAILURE_BLOCK = True

# 市场明显偏弱时，不允许★★★★★主动开仓
ENTRY_GATE_MARKET_SCORE_BLOCK = -3


# ======================================================================
# TDX K线类别
# ======================================================================

CATEGORY_1M = 7
CATEGORY_15M = 1
CATEGORY_30M = 2

# 指数日K
CATEGORY_DAY = 4

# 指数月K
CATEGORY_MONTH = 6


# ======================================================================
# 指数代码
# ======================================================================

SH_INDEX_CODE = "999999"
SZ_INDEX_CODE = "399001"
CY_INDEX_CODE = "399006"

# V7.5.1：三指数独立实时数据
V751_INDEX_CONFIG = {
    "SH": {"code": SH_INDEX_CODE, "market": 1, "name": "上证指数"},
    "SZ": {"code": SZ_INDEX_CODE, "market": 0, "name": "深证成指"},
    "CY": {"code": CY_INDEX_CODE, "market": 0, "name": "创业板指"},
}

# V7.5.3：市场广度改为TDX指数概况直接提供的上涨/下跌家数。
# 不再枚举 get_security_list()，不再用 1494/3000 只证券冒充全市场。
V756_BREADTH_ENABLED = True
V756_BREADTH_CACHE_SECONDS = 20
V756_BREADTH_MIN_ACTIVE = 1000
V756_BREADTH_STRICT = True
V756_BREADTH_INDEX_CODE = TDX_BREADTH_INDEX_CODE
V756_BREADTH_INDEX_MARKET = TDX_BREADTH_INDEX_MARKET


# ======================================================================
# 月K Market Regime
# ======================================================================

MONTH_REGIME_ENABLED = True

# 月K默认不直接进入普通个股评分
MONTH_SCORE_ENABLED = False

# 月K风险过滤
MONTH_5STAR_RISK_FILTER = True

# 以下状态禁止★★★★★
MONTH_BLOCK_5STAR_STATES = {
    "BEARISH",
    "IMPROVING"
}

# 多头减速是否禁止★★★★★
MONTH_BLOCK_TURNING = False


# ======================================================================
# 全局报警记录
# ======================================================================

ALERTED_STOCKS = {}

# ======================================================================
# StockPulse V7.1 PRO
# ----------------------------------------------------------------------
# 四层结构：
#   ① 原始机会评分：回答“这只票有没有机会”
#   ② 个股趋势 + ENTRY_GATE：回答“位置/趋势是否合格”
#   ③ 五星连续确认：回答“信号是不是瞬时噪声”
#   ④ 市场总闸门：回答“今天这个市场允许不允许主动开仓”
#
# 关键修复：
#   V7.0 曾把 stars==5 放在市场闸门之后，导致市场BLOCK时五星计数
#   永远无法累积。V7.1 改成 raw五星候选 → 连续确认 → 市场闸门 → 最终五星。
# ======================================================================
V71_ENABLED = True
V71_CONFIRM_SCANS = 2
V71_STABILITY_COUNT = {}

# ======================================================================
# V7.6：五星“时间确认”状态机
# ----------------------------------------------------------------------
# 不再把30秒轮询当成一次独立确认。
# 首次五星 = PENDING_CONFIRM；之后必须跨越最小确认间隔再次成立。
# 连续2次时间确认后，才进入 CONFIRMED_BUY 候选。
# ======================================================================
V76_FIVE_STAR_CONFIRM_ENABLED = True
V76_FIVE_STAR_CONFIRM_INTERVAL_MIN = 30       # 两次有效确认至少间隔30分钟
V76_FIVE_STAR_PENDING_EXPIRE_MIN = 180        # 超过3小时仍未确认，自动过期
V76_FIVE_STAR_LAST_CHECK_HOUR = 14             # 14点做加强检查，但不是强制等到14点
V76_FIVE_STAR_STATE = {}

# 市场广度独立状态：不再把“指数弱、广度强”粗暴压成同一个结论
V76_BREADTH_BULLISH_RATIO = 0.60
V76_BREADTH_NEUTRAL_RATIO = 0.45

# 市场评分：强弱/偏弱/极弱
V71_MARKET_SCORE_BLOCK = -8
V71_MARKET_SCORE_WAIT = -2

# 指数当日跌幅
V71_INDEX_CHANGE_BLOCK = -2.0
V71_INDEX_CHANGE_WAIT = -1.0

# 指数距离前20日压力
V71_INDEX_PRESSURE_BLOCK = 0.30
V71_INDEX_PRESSURE_WAIT = 0.80

# 月K BEARISH：默认禁止主动追涨；只有极强反转结构才进入人工确认
V71_BEARISH_SPECIAL_ENABLED = True
V71_SPECIAL_MIN_SCORE = 90
V71_SPECIAL_MIN_VOL = 1.20
V71_SPECIAL_MIN_PRESSURE = 0.50
V71_SPECIAL_ALLOWED_TRENDS = {"A", "B"}

# ======================================================================
# StockPulse V7.5.1：Market Gate 五级状态 + 三指数 + 全市场广度
# ----------------------------------------------------------------------
# A = 强势进攻：指数、广度、题材同步偏强
# B = 结构性进攻：指数一般，但市场内部有明显赚钱效应
# C = 震荡观察：方向不清，只允许观察，不主动追涨
# D = 弱势防守：广度/指数明显偏弱，禁止主动开仓
# E = 风险模式：指数急跌、广度恶化或压力破位失败，硬阻断
#
# 月K只描述长期风险，不再直接决定当日市场状态。
# 月KBEARISH：允许A/B中的结构性机会，但五星必须人工确认，
# 并将市场最高等级限制在B，防止长期空头环境下过度乐观。
# ======================================================================
V75_ENABLED = True
V75_MARKET_A_MIN_SCORE = 8.0
V75_MARKET_B_MIN_SCORE = 2.0
V75_MARKET_D_MIN_SCORE = -5.0
V75_MARKET_E_MIN_SCORE = -10.0
V75_INDEX_E_CHANGE = -2.0
V75_INDEX_D_CHANGE = -1.0
V75_PRESSURE_E_PCT = 0.20
V75_PRESSURE_D_PCT = 0.50
V75_BREADTH_A = 0.65
V75_BREADTH_B = 0.50
V75_BREADTH_D = 0.35
V75_BREADTH_E = 0.25
V75_STRONG_UP_PCT = 2.0
V75_STRONG_DOWN_PCT = -2.0
V75_SECTOR_STRONG_PCT = 1.0
V75_SECTOR_WEAK_PCT = -1.0
V75_SECTOR_STRONG_BREADTH = 0.55
V75_SECTOR_WEAK_BREADTH = 0.40
V75_SECTOR_GATE_ENABLED = True
V75_SECTOR_BLOCK_WEAK = True
V75_SECTOR_WAIT_NEUTRAL = True

# V7.7真实板块第二闸门
V77_REAL_SECTOR_ENABLED = True
V77_SECTOR_FILES = ("block_hy.dat",)  # 仅允许真正行业板块文件；禁止block.dat/block_gn.dat冒充行业
V77_SECTOR_CACHE_SECONDS = 120
V77_SECTOR_MIN_MEMBERS = 5
V77_SECTOR_QUOTE_CHUNK = 80
V77_SECTOR_MAX_QUOTE_MEMBERS = 5000
V77_SECTOR_PREFER_EXCEL_GROUP = True
V77_INDUSTRY_BEHAVIOR_ENABLED = True
V78_BEHAVIOR_HISTORY_SIZE = 120          # 最多保留120个1分钟行业快照
V78_BEHAVIOR_MIN_SAMPLES = 3
V78_BEHAVIOR_SAMPLE_SECONDS = 60          # 行业行为每60秒采样一次
V78_BEHAVIOR_TREND_30M_MIN = 30
V78_BEHAVIOR_TREND_60M_MIN = 60
V78_BEHAVIOR_ACCEL_WINDOW_MIN = 10        # 前后各10分钟比较斜率
V78_BEHAVIOR_ACCEL_THRESHOLD_PCT_10M = 0.20
V78_BEHAVIOR_DECEL_THRESHOLD_PCT_10M = -0.20
V78_BEHAVIOR_RETREAT_PCT = 0.80
V78_BEHAVIOR_DUMPING_RETREAT_PCT = 1.50
V78_BEHAVIOR_DIVERGENCE_BREADTH_LOW = 0.35
V78_BEHAVIOR_DIVERGENCE_BREADTH_HIGH = 0.65
V78_BEHAVIOR_DIVERGENCE_STD = 1.80
V78_BEHAVIOR_DIVERGENCE_SPREAD = 4.00
V78_BEHAVIOR_BREADTH_CONTRACTION = -0.08
V78_BEHAVIOR_BREADTH_EXPANSION = 0.08
V78_BEHAVIOR_CACHE_SECONDS = 60
V78_MAPPING_CACHE_SECONDS = 300
V78_BEHAVIOR_ADVERSE = ("RETREATING", "DUMPING", "DIVERGING")
V78_BEHAVIOR_FAVORABLE = ("DRIVING", "ACCELERATING", "TREND_UP", "RECOVERING")
V77_INDUSTRY_CFG_ENABLED = True
# V7.7.2：明确指定本机通达信行业映射文件，避免自动搜索不到导致 0/13。
V77_TDX_CACHE_DIR = os.getenv("TDX_CACHE_DIR", r"D:\zd_zxzq_gm\T0002\hq_cache")
V77_TDX_HY_CFG = os.getenv("TDX_HY_CFG", os.path.join(V77_TDX_CACHE_DIR, "tdxhy.cfg"))
V77_TDX_INDUSTRY_NAMES = (
    os.path.join(V77_TDX_CACHE_DIR, "tdxzs.cfg"),
    os.path.join(V77_TDX_CACHE_DIR, "incon.dat"),
)
V77_INDUSTRY_CFG_NAME = "tdxhy.cfg"
V77_INDUSTRY_NAME_FILES = ("tdxzs.cfg", "incon.dat")
V77_INDUSTRY_UNKNOWN_BLOCK = False
V77_INDUSTRY_WEAK_BLOCK = False
V77_TRADE_PERMISSION_ENABLED = True



# ======================================================================
# StockPulse V7.3：Capital Behavior V2.1 状态稳定器
# ======================================================================
CAPITAL_BEHAVIOR_STABLE_ENABLED = True
CAPITAL_BEHAVIOR_CONFIRM_BARS = 3
CAPITAL_BEHAVIOR_HYSTERESIS = True
CAPITAL_BEHAVIOR_3M_CONFIRM = True
CAPITAL_BEHAVIOR_3M_MIN_BARS = 35
CB_EXIT_ABSORPTION = 6.0
CB_ENTER_ABSORPTION = 10.0
CB_EXIT_DRIVE_EFFICIENCY = 0.12
CB_ENTER_DRIVE_EFFICIENCY = 0.20
CB_EXIT_CHASE = 3.0
CB_ENTER_CHASE = 6.0
CB_EXIT_DISTRIBUTION = 6.0
CB_ENTER_DISTRIBUTION = 9.0
CB_EXIT_DUMP = 7.0
CB_ENTER_DUMP = 9.0
CB_STABILITY_STATE = {}



# ======================================================================
# 安全浮点
# ======================================================================

def calculate_v71_market_gate(market_info, month_regime, trend_info, entry_gate, dar_info,
                              raw_five_star, stable_five_star, final_score,
                              macd_1m=None, macd_15m=None, macd_30m=None,
                              vwap_info=None, sector_gate=None):
    """
    StockPulse V7.5.1 Market Gate。

    第一层：市场当日环境（A/B/C/D/E）。
    第二层：个股所属板块/题材强弱。
    第三层：个股趋势 + ENTRY_GATE + 资金行为。

    月K不再硬阻断当日市场环境；它只作为长期风险层。
    BEARISH时：市场最高只允许到B，五星进入人工确认专区。
    """
    result = {
        "enabled": V75_ENABLED, "passed": False, "state": "E",
        "decision_state": "BLOCK",
        "special": False, "auto_open": False, "reasons": [],
        "market_state": market_info.get("state5", "E"),
        "long_term_risk": "UNKNOWN",
        "sector_gate": sector_gate or {}
    }
    if not V75_ENABLED:
        result.update(passed=True, state="A", auto_open=True)
        return result

    market_state = market_info.get("state5", "E")
    month_state = month_regime.get("state", "UNKNOWN")
    trend_grade = trend_info.get("grade", "UNKNOWN")
    entry_passed = bool(entry_gate.get("passed", False))
    dar_overheat = bool(dar_info.get("overheat", False))
    hard_block = False
    wait = False

    if market_state == "E":
        hard_block = True; result["reasons"].append("市场E级风险模式，禁止主动开仓")
    elif market_state == "D":
        hard_block = True; result["reasons"].append("市场D级弱势防守，禁止主动开仓")
    elif market_state == "C":
        wait = True; result["reasons"].append("市场C级震荡，只观察")
    elif market_state == "B":
        wait = False; result["reasons"].append("市场B级结构性进攻：允许筛选强势个股")
    elif market_state == "A":
        result["reasons"].append("市场A级强势进攻")

    # 长期风险：只限速，不否定当日环境。
    if month_state == "BEARISH":
        result["long_term_risk"] = "HIGH"
        if market_state == "A":
            result["market_state"] = "B"
            result["reasons"].append("月KBEARISH：长期风险高，市场最高等级限制为B")
        else:
            result["reasons"].append("月KBEARISH：长期风险高，但不再直接阻断当日结构性机会")
    elif month_state == "IMPROVING":
        result["long_term_risk"] = "MEDIUM"
        result["reasons"].append("月K空头改善：长期风险中等")
    elif month_state in {"BULLISH", "TURNING"}:
        result["long_term_risk"] = "LOW"
    else:
        result["long_term_risk"] = "UNKNOWN"

    if trend_grade in {"D", "E"}:
        hard_block = True; result["reasons"].append(f"个股趋势{trend_grade}级，禁止主动开仓")
    elif trend_grade == "C":
        wait = True; result["reasons"].append("个股趋势C级，仅观察")

    if not entry_passed:
        est = entry_gate.get("state", "BLOCK")
        if est == "BLOCK": hard_block = True
        else: wait = True
        result["reasons"].append("ENTRY_GATE未通过")

    if dar_overheat:
        wait = True; result["reasons"].append("DAR量价过热，禁止追涨")

    # V7.7.2 行业第二闸门：行业只做“确认/降级”，不再把错误的block.dat分类当行业。
    # WEAK默认降级为WAIT，不直接BLOCK；UNKNOWN只记录，不阻断。
    sg = sector_gate or {}
    sg_state = sg.get("state", "UNKNOWN")
    sg_source = str(sg.get("source", "NONE") or "NONE")
    if V75_SECTOR_GATE_ENABLED:
        if sg_state == "WEAK":
            if V77_INDUSTRY_WEAK_BLOCK and sg_source.startswith("TDX_INDUSTRY"):
                hard_block = True
                result["reasons"].append("所属行业明显偏弱，第二层闸门阻断")
            else:
                wait = True
                result["reasons"].append("所属行业偏弱：降级观察，不直接阻断")
        elif sg_state == "NEUTRAL":
            if V75_SECTOR_WAIT_NEUTRAL:
                wait = True
                result["reasons"].append("所属行业中性：不足以支持主动追涨")
        elif sg_state == "UNKNOWN":
            result["reasons"].append("所属行业数据不可用：不阻断，仅作为未知风险")
        elif sg_state == "STRONG":
            result["reasons"].append("所属行业强势，第二层闸门通过")

    # 月K空头下的极强结构，进入人工确认专区；不自动开仓。
    if month_state == "BEARISH":
        special = (
            V71_BEARISH_SPECIAL_ENABLED and raw_five_star and stable_five_star
            and final_score >= V71_SPECIAL_MIN_SCORE
            and trend_grade in V71_SPECIAL_ALLOWED_TRENDS
            and not dar_overheat
            and safe_float(dar_info.get("vol_ratio"), 0) >= V71_SPECIAL_MIN_VOL
            and safe_float(entry_gate.get("distance_to_pressure"), 0) >= V71_SPECIAL_MIN_PRESSURE
            and sg_state == "STRONG"
            and not hard_block
        )
        if special:
            result.update(passed=True, state="B", decision_state="HUMAN_CONFIRM", special=True, auto_open=False)
            result["reasons"].append("月K空头 + 当日结构性进攻 + 板块强势：仅人工确认，不自动开仓")
            return result

    # V7.7：state只表示市场环境；decision_state表示个股最终闸门。
    result["state"] = market_state
    if hard_block:
        result["decision_state"] = "BLOCK"
        result["passed"] = False
    elif wait:
        result["decision_state"] = "WAIT"
        result["passed"] = False
    else:
        result["decision_state"] = "PASS"
        result["passed"] = market_state in {"A", "B"}
        result["auto_open"] = False
    return result


# ======================================================================
# 安全浮点
# ======================================================================


    try:

        if value is None:
            return default

        if pd.isna(value):
            return default

        x = float(value)

        if not np.isfinite(x):
            return default

        return x

    except (ValueError, TypeError):

        return default


# ======================================================================
# DataFrame标准化
# ======================================================================

def normalize_df_columns(df):

    if df is None or df.empty:
        return df

    d = df.copy()

    d.columns = [
        str(c).lower()
        for c in d.columns
    ]

    if (
        "vol" in d.columns
        and "volume" not in d.columns
    ):
        d.rename(
            columns={
                "vol": "volume"
            },
            inplace=True
        )

    for col in [
        "open",
        "high",
        "low",
        "close",
        "volume",
        "amount"
    ]:

        if col in d.columns:

            d[col] = pd.to_numeric(
                d[col],
                errors="coerce"
            )

    if (
        "amount" not in d.columns
        and "volume" in d.columns
        and "close" in d.columns
    ):

        d["amount"] = (
            d["close"]
            * d["volume"]
            * 100
        )

    if "datetime" in d.columns:

        d["datetime"] = pd.to_datetime(
            d["datetime"],
            errors="coerce"
        )

        d = d.dropna(
            subset=["datetime"]
        )

        d = (
            d
            .sort_values("datetime")
            .reset_index(drop=True)
        )

    return d


# ======================================================================
# 安全数值转换
# ======================================================================
def safe_float(value, default=np.nan):
    """安全转换为 float；None/空值/NaN/异常值时返回 default。"""
    try:
        if value is None:
            return default
        if isinstance(value, str) and not value.strip():
            return default
        x = float(value)
        if np.isnan(x):
            return default
        return x
    except (TypeError, ValueError, OverflowError):
        return default


# ======================================================================
# 市场代码
# ======================================================================

def get_market_code(code):

    c = (
        str(code)
        .strip()
        .replace(".0", "")
        .zfill(6)
    )

    if c.startswith(
        ("6", "9", "5", "7")
    ):
        return 1

    return 0


# ======================================================================
# 窗口聚焦
# ======================================================================

def focus_terminal_window():

    try:

        wins = (
            gw.getWindowsWithTitle(
                os.path.basename(sys.argv[0])
            )
            or
            gw.getWindowsWithTitle("cmd")
            or
            gw.getWindowsWithTitle("PowerShell")
        )

        if wins:

            win = wins[0]

            if win.isMinimized:
                win.restore()

            win.activate()

    except Exception as e:

        logger.debug(
            f"窗口聚焦异常: {e}"
        )


# ======================================================================
# ★★★★★ 人工确认
# ======================================================================

def ask_human_confirmation_cmd(
    code,
    name
):

    focus_terminal_window()

    result = ctypes.windll.user32.MessageBoxW(
        0,
        (
            f"标的 [{name} ({code})] "
            f"触发最高级 ★★★★★ 买入信号！\n\n"
            f"是否确认写入Excel？"
        ),
        "五星级高优先级信号确认",
        0x24
    )

    return result == 6


# ======================================================================
# 语音
# ======================================================================

def speak_text_async(text):

    def run():

        try:

            engine = pyttsx3.init()

            engine.setProperty(
                "rate",
                180
            )

            engine.setProperty(
                "volume",
                1.0
            )

            engine.say(text)

            engine.runAndWait()

        except Exception as e:

            logger.debug(
                f"语音异常: {e}"
            )

    threading.Thread(
        target=run,
        daemon=True
    ).start()


# ======================================================================
# 交易时间
# ======================================================================

def is_trading_time():

    now = datetime.datetime.now()

    if now.weekday() >= 5:

        return False, "非交易日 (周末休市)"

    t = now.time()

    if (
        datetime.time(9, 30)
        <= t
        <= datetime.time(11, 30)
        or
        datetime.time(13, 0)
        <= t
        <= datetime.time(15, 0)
    ):

        return True, "交易进行中"

    if t < datetime.time(9, 30):

        return False, "等待早盘开盘 (09:30)"

    if t < datetime.time(13, 0):

        return False, "午间休市中 (11:30-13:00)"

    return False, "今日已收盘 (15:00)"


# ======================================================================
# TDX服务器池
# ======================================================================

class TDXServerPool:

    SERVERS = [

        (
            "218.75.126.9",
            7709,
            "浙江电信"
        ),

        (
            "119.147.212.81",
            7709,
            "深圳电信"
        ),

        (
            "47.103.48.45",
            7709,
            "华东阿里云"
        ),

        (
            "106.120.74.86",
            7709,
            "北京联通"
        ),

        (
            "180.153.39.51",
            7709,
            "上海电信"
        ),

        (
            "123.125.108.23",
            7709,
            "北京移动"
        )
    ]

    def __init__(self):

        self.api = TdxHq_API()

        self.connected = False

        self.idx = 0

    # ------------------------------------------------------------------
    # 连接
    # ------------------------------------------------------------------

    def connect(self):

        if self.connected:
            return True

        for _ in range(
            len(self.SERVERS)
        ):

            ip, port, name = (
                self.SERVERS[self.idx]
            )

            try:

                if self.api.connect(
                    ip,
                    port,
                    time_out=3.0
                ):

                    self.connected = True

                    logger.info(
                        f"成功连接至 TDX 行情服务器: "
                        f"{name} ({ip})"
                    )

                    return True

            except Exception as e:

                logger.warning(
                    f"节点{name}连接异常: {e}"
                )

            self.idx = (
                self.idx + 1
            ) % len(self.SERVERS)

        self.connected = False

        logger.error(
            "所有通达信行情服务器连接失败"
        )

        return False

    # ------------------------------------------------------------------
    # 通用读取
    # ------------------------------------------------------------------

    def _get(
        self,
        function,
        *args
    ):

        if not self.connect():
            return None

        try:

            bars = function(*args)

            if bars:

                return normalize_df_columns(
                    pd.DataFrame(bars)
                )

            return None

        except Exception as e:

            logger.error(
                f"TDX行情读取异常: {e}"
            )

            self.connected = False

            try:

                if self.connect():

                    bars = function(*args)

                    if bars:

                        return normalize_df_columns(
                            pd.DataFrame(bars)
                        )

            except Exception:

                pass

        return None

    # ------------------------------------------------------------------
    # 个股K线
    # ------------------------------------------------------------------

    def get_bars(
        self,
        category,
        market,
        code,
        start,
        count
    ):

        return self._get(
            self.api.get_security_bars,
            category,
            market,
            code,
            start,
            count
        )

    # ------------------------------------------------------------------
    # 指数K线
    #
    # 999999 / 399001必须走这里
    # ------------------------------------------------------------------

    def get_index_bars(
        self,
        category,
        market,
        code,
        start,
        count
    ):

        return self._get(
            self.api.get_index_bars,
            category,
            market,
            code,
            start,
            count
        )


    def get_block_members(self, blockfile):
        """读取TDX真实板块文件；失败返回None，不伪造。"""
        if get_and_parse_block_info is None or not self.connect():
            return None
        try:
            return get_and_parse_block_info(self.api, blockfile)
        except Exception as e:
            logger.debug(f"TDX板块文件 {blockfile} 读取失败: {e}")
            return None


# ======================================================================
# MACD空结果
# ======================================================================

def _empty_macd():

    return {
        "state": "数据不足",
        "valid": False,
        "score": 0,

        "dif": np.nan,
        "dea": np.nan,
        "hist": np.nan,
        "prev_hist": np.nan,

        "golden_cross": False,
        "death_cross": False,

        "golden_cross_event": False,
        "death_cross_event": False
    }


# ======================================================================
# MACD
# ======================================================================

def calculate_macd(df):

    if (
        df is None
        or df.empty
        or "close" not in df.columns
    ):

        return _empty_macd()

    close = pd.to_numeric(
        df["close"],
        errors="coerce"
    ).dropna()

    if len(close) < 35:

        return _empty_macd()

    ema12 = (
        close
        .ewm(
            span=12,
            adjust=False
        )
        .mean()
    )

    ema26 = (
        close
        .ewm(
            span=26,
            adjust=False
        )
        .mean()
    )

    dif = ema12 - ema26

    dea = (
        dif
        .ewm(
            span=9,
            adjust=False
        )
        .mean()
    )

    hist = (
        dif - dea
    ) * 2

    d = safe_float(
        dif.iloc[-1]
    )

    e = safe_float(
        dea.iloc[-1]
    )

    h = safe_float(
        hist.iloc[-1]
    )

    ph = safe_float(
        hist.iloc[-2]
    )

    pdif = safe_float(
        dif.iloc[-2]
    )

    pdea = safe_float(
        dea.iloc[-2]
    )

    golden = d > e
    death = d < e

    golden_event = (
        pdif <= pdea
        and
        d > e
    )

    death_event = (
        pdif >= pdea
        and
        d < e
    )

    if (
        d > e
        and
        h > 0
        and
        h > ph
    ):

        state = "🟢多头加速"
        score = 12

    elif (
        d < e
        and
        h < 0
        and
        h < ph
    ):

        state = "🔴空头加速"
        score = -12

    elif (
        d < e
        and
        h < 0
        and
        h > ph
    ):

        state = "🟡空头改善"
        score = 0

    elif (
        d > e
        and
        h > 0
        and
        h < ph
    ):

        state = "🟠多头减速"
        score = 0

    elif golden:

        state = "🟢多头结构"
        score = 6

    elif death:

        state = "🔴空头结构"
        score = -6

    else:

        state = "⚪震荡平缓"
        score = 0

    return {

        "state": state,
        "valid": True,
        "score": score,

        "dif": d,
        "dea": e,
        "hist": h,
        "prev_hist": ph,

        "golden_cross": golden,
        "death_cross": death,

        "golden_cross_event":
            golden_event,

        "death_cross_event":
            death_event
    }


# ======================================================================
# 上证月K Market Regime
# ======================================================================

def calculate_month_market_regime(df):

    base = {

        "valid": False,

        "state": "UNKNOWN",

        "text": "数据不足",

        "source": "NONE",

        "bars": 0,

        "dif": np.nan,

        "dea": np.nan,

        "hist": np.nan,

        "prev_hist": np.nan,

        "structure": "数据不足",

        "risk_block_5star": False,

        "golden_cross_event": False,

        "death_cross_event": False
    }

    if (
        df is None
        or df.empty
        or "close" not in df.columns
    ):

        return base

    d = df.copy()

    if "datetime" in d.columns:

        d["datetime"] = pd.to_datetime(
            d["datetime"],
            errors="coerce"
        )

        d = (
            d
            .dropna(subset=["datetime"])
            .sort_values("datetime")
        )

    close = pd.to_numeric(
        d["close"],
        errors="coerce"
    ).dropna()

    # ----------------------------------------------------------
    # 关键修复：
    #
    # 当前月份尚未收盘时，绝对不能拿它判断月K死叉。
    #
    # 否则月末盘中：
    #
    # DIF > DEA
    #      ↓
    # 盘中下跌
    #      ↓
    # DIF < DEA
    #
    # 程序就会误判成“月K死叉”。
    # ----------------------------------------------------------

    if (
        len(close) > 1
        and
        "datetime" in d.columns
    ):

        last_dt = (
            d
            .loc[
                close.index,
                "datetime"
            ]
            .iloc[-1]
        )

        now = datetime.datetime.now()

        if (
            last_dt.year == now.year
            and
            last_dt.month == now.month
        ):

            close = close.iloc[:-1]

    if len(close) < 35:

        return {
            **base,
            "source": "TDX_MONTHLY",
            "bars": len(close)
        }

    ema12 = (
        close
        .ewm(
            span=12,
            adjust=False
        )
        .mean()
    )

    ema26 = (
        close
        .ewm(
            span=26,
            adjust=False
        )
        .mean()
    )

    dif = ema12 - ema26

    dea = (
        dif
        .ewm(
            span=9,
            adjust=False
        )
        .mean()
    )

    hist = (
        dif - dea
    ) * 2

    x = safe_float(
        dif.iloc[-1]
    )

    y = safe_float(
        dea.iloc[-1]
    )

    h = safe_float(
        hist.iloc[-1]
    )

    ph = safe_float(
        hist.iloc[-2]
    )

    px = safe_float(
        dif.iloc[-2]
    )

    py = safe_float(
        dea.iloc[-2]
    )

    golden_event = (
        px <= py
        and
        x > y
    )

    death_event = (
        px >= py
        and
        x < y
    )

    # ----------------------------------------------------------
    # 月K状态
    # ----------------------------------------------------------

    if (
        x > y
        and
        h > 0
        and
        h > ph
    ):

        state = "BULLISH"

        text = "🟢月K多头加速"

        structure = "🟢MACD多头结构"

    elif (
        x > y
        and
        h > 0
    ):

        state = "TURNING"

        text = "🟡月K多头减速"

        structure = "🟡MACD多头减速"

    elif (
        x < y
        and
        h < 0
        and
        h > ph
    ):

        state = "IMPROVING"

        text = "🟡月K空头改善"

        structure = "🟡MACD死叉/空头结构"

    elif (
        x < y
        and
        h < 0
    ):

        state = "BEARISH"

        text = "🔴月K空头加速"

        structure = "🔴MACD死叉/空头结构"

    else:

        state = "UNKNOWN"

        text = "⚪月K震荡"

        structure = "⚪MACD震荡"

    # ----------------------------------------------------------
    # 五星风险
    #
    # 这里非常重要：
    #
    # 只要 DIF < DEA，
    # 即使 HIST 开始改善，
    # 仍然认为月线处于空头结构。
    #
    # “空头改善” ≠ “重新进入多头”。
    # ----------------------------------------------------------

    risk_block = False

    if MONTH_5STAR_RISK_FILTER:

        if x < y:

            risk_block = True

        if state in MONTH_BLOCK_5STAR_STATES:

            risk_block = True

        if (
            MONTH_BLOCK_TURNING
            and
            state == "TURNING"
        ):

            risk_block = True

    return {

        "valid": True,

        "state": state,

        "text": text,

        "source":
            "TDX_MONTHLY_COMPLETED",

        "bars": len(close),

        "dif": x,

        "dea": y,

        "hist": h,

        "prev_hist": ph,

        "structure": structure,

        "risk_block_5star":
            risk_block,

        "golden_cross_event":
            golden_event,

        "death_cross_event":
            death_event
    }


# ======================================================================
# VWAP
# ======================================================================

def calculate_vwap(df):

    result = {

        "valid": False,

        "vwap": np.nan,

        "distance": np.nan,

        "reason": "数据不足",

        "score": 0
    }

    if df is None or df.empty:

        return result

    d = df.copy()

    # ----------------------------------------------------------
    # VWAP只计算当前交易日
    # ----------------------------------------------------------

    if "datetime" in d.columns:

        d["datetime"] = pd.to_datetime(
            d["datetime"],
            errors="coerce"
        )

        d = d.dropna(
            subset=["datetime"]
        )

        if not d.empty:

            latest_date = (
                d["datetime"]
                .dt.date
                .max()
            )

            d = d[
                d["datetime"].dt.date
                ==
                latest_date
            ]

    required = [
        "high",
        "low",
        "close",
        "volume"
    ]

    if not all(
        c in d.columns
        for c in required
    ):

        result["reason"] = "字段不足"

        return result

    for c in required:

        d[c] = pd.to_numeric(
            d[c],
            errors="coerce"
        )

    d = d.dropna(
        subset=required
    )

    if d.empty:

        return result

    volume = d["volume"].clip(
        lower=0
    )

    total_volume = volume.sum()

    if total_volume <= 0:

        result["reason"] = "成交量为0"

        return result

    typical_price = (
        d["high"]
        +
        d["low"]
        +
        d["close"]
    ) / 3.0

    vwap = (
        typical_price
        *
        volume
    ).sum() / total_volume

    current = safe_float(
        d["close"].iloc[-1]
    )

    if (
        current <= 0
        or
        not np.isfinite(vwap)
        or
        vwap <= 0
    ):

        result["reason"] = "当前价格异常"

        return result

    distance = (
        current - vwap
    ) / vwap * 100

    # ----------------------------------------------------------
    # VWAP异常熔断
    # ----------------------------------------------------------

    if abs(distance) > 15:

        result["reason"] = (
            "VWAP偏离异常熔断"
        )

        return result

    return {

        "valid": True,

        "vwap": float(vwap),

        "distance": float(distance),

        "reason": "有效",

        "score":
            8
            if distance > 0
            else
            -8
    }


# ======================================================================
# DAR + 动态量能
# ======================================================================

def calculate_dar_and_volume(
    df,
    sell_buy_ratio=1.2
):

    if (
        df is None
        or
        len(df) < 6
        or
        "volume" not in df.columns
    ):

        return {

            "dar_score": 0,

            "dar_state": "数据不足",

            "overheat": False,

            "vol_ratio": np.nan
        }

    d = df.copy()

    for c in [
        "volume",
        "open",
        "high",
        "low",
        "close"
    ]:

        if c in d.columns:

            d[c] = pd.to_numeric(
                d[c],
                errors="coerce"
            )

    d = d.dropna(
        subset=["volume"]
    )

    if len(d) < 6:

        return {

            "dar_score": 0,

            "dar_state": "数据不足",

            "overheat": False,

            "vol_ratio": np.nan
        }

    latest = d.iloc[-1]

    current_volume = safe_float(
        latest.get("volume"),
        0
    )

    # ----------------------------------------------------------
    # 关键修复：
    #
    # 原版：
    #
    # MA = 最近5根包含当前K线
    #
    # 现在：
    #
    # MA = 当前K线之前5根
    #
    # 避免“放量越大，自己的基准也越大”的自污染。
    # ----------------------------------------------------------

    previous_volume = d[
        "volume"
    ].iloc[-6:-1]

    avg_volume = safe_float(
        previous_volume.mean(),
        0
    )

    vol_ratio = (
        current_volume / avg_volume
        if avg_volume > 0
        else 0
    )

    open_price = safe_float(
        latest.get("open")
    )

    high_price = safe_float(
        latest.get("high")
    )

    low_price = safe_float(
        latest.get("low")
    )

    close_price = safe_float(
        latest.get("close")
    )

    if (
        np.isfinite(high_price)
        and
        np.isfinite(low_price)
    ):

        rng = max(
            high_price - low_price,
            0
        )

    else:

        rng = 0

    if (
        rng > 0
        and
        np.isfinite(close_price)
    ):

        upper_pressure = (
            high_price
            -
            close_price
        ) / rng

    else:

        upper_pressure = 0

    # ----------------------------------------------------------
    # 量价过热
    # ----------------------------------------------------------

    overheat = (
        vol_ratio >= 3.5
        and
        upper_pressure >= 0.60
    )

    sell_buy_ratio = safe_float(
        sell_buy_ratio,
        1.2
    )

    # ----------------------------------------------------------
    # DAR评分
    # ----------------------------------------------------------

    if (
        vol_ratio >= 2.0
        and
        close_price > open_price
        and
        sell_buy_ratio < 1.5
    ):

        dar_state = (
            "🟢主力砸盘吸收反转"
        )

        dar_score = 15

    elif sell_buy_ratio > 3.0:

        dar_state = (
            "🔴高抛压压制"
        )

        dar_score = -10

    elif (
        vol_ratio >= 1.5
        and
        close_price >= open_price
    ):

        dar_state = (
            "🟢放量偏强"
        )

        dar_score = 10

    elif (
        vol_ratio < 0.7
        and
        close_price < open_price
    ):

        dar_state = (
            "🟡缩量偏弱"
        )

        dar_score = -5

    else:

        dar_state = (
            "⚪量能正常"
        )

        dar_score = 5

    return {

        "dar_score":
            dar_score,

        "dar_state":
            dar_state,

        "overheat":
            overheat,

        "vol_ratio":
            vol_ratio
    }


# ======================================================================
# V7.5 板块/题材第二层闸门
# ----------------------------------------------------------------------
# Excel可选列：板块 / 行业 / 题材。优先级：题材 > 板块 > 行业。
# 若Excel没有这些列，则退化为“个股相对强弱”，不伪造板块数据。
# ======================================================================
V75_SECTOR_COLUMNS = ("题材", "板块", "行业")
V75_SECTOR_FALLBACK_ENABLED = False  # V7.5.1：彻底删除个股相对强弱假板块


def _get_sector_name(row):
    if row is None:
        return "未分类"
    for col in V75_SECTOR_COLUMNS:
        v = str(row.get(col, "") or "").strip()
        if v and v.lower() not in {"nan", "none"}:
            return v
    return "未分类"


def _build_sector_gate_result(sector, item):
    ret = safe_float(item.get("mean_change"), np.nan)
    breadth = safe_float(item.get("up_ratio"), np.nan)
    members = int(item.get("members", 0))
    source = item.get("source", "NONE")
    behavior = str(item.get("behavior_state", "UNKNOWN") or "UNKNOWN")
    if members < V77_SECTOR_MIN_MEMBERS or not np.isfinite(ret) or not np.isfinite(breadth):
        state = "UNKNOWN"
    elif behavior in V78_BEHAVIOR_ADVERSE:
        state = "WEAK" if behavior in ("RETREATING", "DUMPING") else "NEUTRAL"
    elif ret >= V75_SECTOR_STRONG_PCT and breadth >= V75_SECTOR_STRONG_BREADTH:
        state = "STRONG"
    elif ret <= V75_SECTOR_WEAK_PCT or breadth < V75_SECTOR_WEAK_BREADTH:
        state = "WEAK"
    else:
        state = "NEUTRAL"
    behavior_state = str(item.get("behavior_state", "UNKNOWN") or "UNKNOWN")
    return {"state": state, "sector": sector, "mean_change": ret, "up_ratio": breadth,
            "members": members, "source": source, "behavior_state": behavior_state,
            "behavior_trend": item.get("behavior_trend", "UNKNOWN"),
            "behavior_trend_30m": safe_float(item.get("behavior_trend_30m"), np.nan),
            "behavior_trend_60m": safe_float(item.get("behavior_trend_60m"), np.nan),
            "behavior_accel": safe_float(item.get("behavior_accel"), np.nan),
            "behavior_retreat": safe_float(item.get("behavior_retreat"), np.nan),
            "behavior_dispersion": safe_float(item.get("behavior_dispersion"), np.nan),
            "behavior_spread": safe_float(item.get("behavior_spread"), np.nan),
            "behavior_breadth_delta": safe_float(item.get("behavior_breadth_delta"), np.nan),
            "behavior_samples": int(item.get("behavior_samples", 0) or 0),
            "behavior_sampled": bool(item.get("behavior_sampled", False))}


def calculate_sector_gate(code, row, market_info, stock_df=None):
    """V7.7.2：真正“行业”第二闸门。

    优先级：
      1) Excel显式“行业”列（若存在有效实时统计）
      2) TDX tdxhy.cfg -> 通达信行业代码 -> tdxzs.cfg/incon.dat名称
      3) 无可靠行业数据 => UNKNOWN

    明确禁止：block.dat、block_gn.dat作为行业来源；也禁止个股相对强弱冒充行业。
    """
    code = str(code).replace(".0", "").strip().zfill(6)
    stats = market_info.get("sector_stats", {}) if isinstance(market_info, dict) else {}

    excel_industry = ""
    if row is not None:
        excel_industry = str(row.get("行业", "") or "").strip()
        if excel_industry.lower() in {"nan", "none", "未分类"}:
            excel_industry = ""

    if V77_SECTOR_PREFER_EXCEL_GROUP and excel_industry:
        item = stats.get(excel_industry)
        if item and str(item.get("source", "")).startswith("WATCHLIST_INDUSTRY"):
            return _build_sector_gate_result(excel_industry, item)

    sector_map = market_info.get("sector_map", {}) if isinstance(market_info, dict) else {}
    sector = sector_map.get(code)
    if sector:
        item = stats.get(sector)
        if item:
            return _build_sector_gate_result(sector, item)
        return {"state":"UNKNOWN", "sector":sector, "mean_change":np.nan,
                "up_ratio":np.nan, "members":0,
                "source":market_info.get("sector_source", "TDX_INDUSTRY_UNKNOWN")}

    return {"state":"UNKNOWN", "sector":excel_industry or "未分类",
            "mean_change":np.nan, "up_ratio":np.nan, "members":0,
            "source":"NONE"}


# ======================================================================
# 市场指数分析
# ======================================================================

class MarketAnalyzer:
    def __init__(self, tdx):
        self.tdx = tdx
        self._a_share_codes_cache = []
        self._a_share_codes_cache_time = 0.0
        self._sector_map_cache = {}
        self._sector_stats_cache = {}
        self._sector_source_cache = "NONE"
        self._sector_cache_time = 0.0
        self._industry_behavior_history = {}
        self._industry_behavior_cache_time = 0.0
        self._industry_behavior_date = None
        self._industry_behavior_last_sample = 0.0
        self._sector_mapping_cache = {}
        self._sector_mapping_members = {}
        self._sector_mapping_source = "TDX_INDUSTRY_UNKNOWN"
        self._sector_mapping_cache_time = 0.0

    def get_index_day_data(self, code):
        cfg = next((v for v in V751_INDEX_CONFIG.values() if v["code"] == code), None)
        market = cfg["market"] if cfg else get_market_code(code)
        return self.tdx.get_index_bars(CATEGORY_DAY, market, code, 0, 120)

    def get_index_intraday_data(self, code):
        cfg = next((v for v in V751_INDEX_CONFIG.values() if v["code"] == code), None)
        market = cfg["market"] if cfg else get_market_code(code)
        return self.tdx.get_index_bars(CATEGORY_1M, market, code, 0, 240)

    def get_sh_month_data(self):
        return self.tdx.get_index_bars(CATEGORY_MONTH, 1, SH_INDEX_CODE, 0, 120)

    def get_market_regime(self):
        if not MONTH_REGIME_ENABLED:
            return {"valid": False, "state": "DISABLED", "text": "月K关闭", "source": "DISABLED", "bars": 0, "risk_block_5star": False}
        return calculate_month_market_regime(self.get_sh_month_data())

    @staticmethod
    def _pool_group(row):
        return _get_sector_name(row)

    def _get_stock_intraday_change(self, code, market):
        # V7.5.1：实时广度使用 quote 的 price/last_close，避免用1M第一根K线
        # 冒充“当日涨跌”。
        try:
            q = self.tdx.get_security_quotes([(market, code)])
            if q is not None and len(q) > 0:
                item = q[0]
                price = safe_float(item.get("price"), np.nan)
                last_close = safe_float(item.get("last_close"), np.nan)
                if np.isfinite(price) and np.isfinite(last_close) and last_close > 0:
                    return (price / last_close - 1.0) * 100.0
        except Exception:
            pass
        return np.nan

    def _get_tdx_market_breadth(self, stock_pool=None):
        """V7.6.0：TDX原生880005作为唯一权威全市场广度源。

        读取顺序：
        1) get_security_quotes(1, 880005) —— 实时统计指数；
        2) 若服务器/pytdx不返回该Quote，则改用 get_index_bars(category=1M, 1, 880005)
           读取最近已返回的统计指数K线；
        3) 两条路径都失败才判定广度不可用。

        绝不再枚举证券全集，也不拼接999999/399001的伪沪深子广度。
        """
        empty = {
            "valid": False, "source": "TDX_NATIVE_BREADTH_UNAVAILABLE",
            "members": 0, "up": 0, "down": 0, "flat": 0,
            "up_ratio": np.nan, "down_ratio": np.nan,
            "strong_up_ratio": np.nan, "strong_down_ratio": np.nan,
            "changes": {}, "sector_stats": {}, "watchlist": {},
            "sh": {}, "sz": {},
        }
        if not V756_BREADTH_ENABLED:
            empty["source"] = "DISABLED"
            return empty

        def _decode(raw_up, raw_down, raw_total, source):
            """将TDX 880005统计指数解码为真实家数，并做严格一致性校验。"""
            try:
                raw_up = float(raw_up)
                raw_down = float(raw_down)
                raw_total = float(raw_total)
            except Exception:
                return None
            if not all(np.isfinite(x) for x in (raw_up, raw_down, raw_total)):
                return None
            if raw_total <= 0:
                return None

            # TDX统计指数通常按1/10家数返回：例如 341.7 -> 3417家。
            candidates = []
            for scale in (10.0, 1.0):
                up = int(round(raw_up * scale))
                down = int(round(raw_down * scale))
                members = int(round(raw_total * scale))
                if members >= V756_BREADTH_MIN_ACTIVE and members <= 10000 \
                        and 0 <= up <= members and 0 <= down <= members \
                        and up + down <= members:
                    candidates.append((members, up, down, scale))

            if not candidates:
                return None

            # 优先选择符合A股全市场典型规模的10倍解码；若服务器直接返回整数则允许1倍。
            members, up, down, scale = max(candidates, key=lambda x: (x[3] == 10.0, x[0]))
            flat = max(0, members - up - down)
            return {
                "valid": True,
                "source": source,
                "members": members,
                "up": up,
                "down": down,
                "flat": flat,
                "neutral_or_suspended": flat,
                "up_ratio": up / members,
                "down_ratio": down / members,
                "strong_up_ratio": np.nan,
                "strong_down_ratio": np.nan,
                "changes": {}, "sector_stats": {}, "watchlist": {},
                "sh": {}, "sz": {},
            }

        # ----------------------------------------------------------
        # 路径A：实时Quote
        # ----------------------------------------------------------
        try:
            quotes = self.tdx.api.get_security_quotes([
                (V756_BREADTH_INDEX_MARKET, V756_BREADTH_INDEX_CODE)
            ])
            if quotes:
                q = quotes[0]
                result = _decode(
                    q.get("price", q.get("close", np.nan)),
                    q.get("open", np.nan),
                    q.get("high", q.get("high_price", np.nan)),
                    "TDX_NATIVE_880005_QUOTE",
                )
                if result:
                    if stock_pool is not None and len(stock_pool) > 0:
                        result["sector_stats"] = self._calculate_watchlist_sector_stats(stock_pool)
                    logger.info(
                        "V7.6.0市场广度："
                        f"全市场 ↑{result['up']} ↓{result['down']} "
                        f"→/停={result['flat']} / {result['members']} | "
                        f"↑占比={result['up_ratio']*100:.1f}% | "
                        f"source={result['source']}"
                    )
                    return result
        except Exception as e:
            logger.warning(f"V7.5.6 TDX 880005实时Quote读取失败：{e}")

        # ----------------------------------------------------------
        # 路径B：统计指数1分钟K线备用
        # ----------------------------------------------------------
        try:
            bars = self.tdx.api.get_index_bars(
                CATEGORY_1M,
                V756_BREADTH_INDEX_MARKET,
                V756_BREADTH_INDEX_CODE,
                0,
                5,
            )
            if bars:
                # 最新一根即可。若当前分钟尚未形成完整值，向前寻找最近有效值。
                for bar in reversed(bars):
                    result = _decode(
                        bar.get("close", np.nan),
                        bar.get("open", np.nan),
                        bar.get("high", np.nan),
                        "TDX_NATIVE_880005_INDEX_BAR",
                    )
                    if result:
                        if stock_pool is not None and len(stock_pool) > 0:
                            result["sector_stats"] = self._calculate_watchlist_sector_stats(stock_pool)
                        logger.info(
                            "V7.6.0市场广度："
                            f"全市场 ↑{result['up']} ↓{result['down']} "
                            f"→/停={result['flat']} / {result['members']} | "
                            f"↑占比={result['up_ratio']*100:.1f}% | "
                            f"source={result['source']}"
                        )
                        return result
        except Exception as e:
            logger.warning(f"V7.5.6 TDX 880005指数K线备用读取失败：{e}")

        logger.warning("V7.6.0市场广度不可用：880005实时Quote及指数K线均未返回有效全市场家数。")
        return empty

    def _calculate_watchlist_sector_stats(self, stock_pool):
        """仅计算Excel监测池内、明确分组的板块统计。"""
        result = {}
        try:
            watch_codes = []
            group_by_code = {}
            for _, row in stock_pool.iterrows():
                code = str(row.get("代码", "")).replace(".0", "").strip().zfill(6)
                if not code:
                    continue
                sector = self._pool_group(row)
                if sector == "未分类":
                    continue
                watch_codes.append((get_market_code(code), code))
                group_by_code[code] = sector

            if not watch_codes:
                return result

            q = self.tdx.api.get_security_quotes(watch_codes)
            groups = {}
            for item in q or []:
                code = str(item.get("code", "")).strip().zfill(6)
                sector = group_by_code.get(code)
                if not sector:
                    continue
                price = safe_float(item.get("price"), np.nan)
                last_close = safe_float(item.get("last_close"), np.nan)
                if np.isfinite(price) and np.isfinite(last_close) and last_close > 0:
                    groups.setdefault(sector, []).append((price / last_close - 1.0) * 100.0)

            for sector, arr in groups.items():
                if len(arr) >= 3:
                    a = np.asarray(arr, dtype=float)
                    result[sector] = {
                        "members": int(len(a)),
                        "mean_change": float(np.mean(a)),
                        "up_ratio": float(np.mean(a > 0)),
                        "strong_up_ratio": float(np.mean(a >= V75_STRONG_UP_PCT)),
                        "strong_down_ratio": float(np.mean(a <= V75_STRONG_DOWN_PCT)),
                    }
        except Exception as e:
            logger.debug(f"V7.5.4监测池板块统计失败：{e}")
        return result

    @staticmethod
    def _candidate_tdx_cache_dirs():
        """寻找通达信本地T0002/hq_cache；不依赖固定安装目录。"""
        candidates = []
        env_names = ("TDX_HOME", "TDX_PATH", "TDX_DIR")
        for name in env_names:
            v = os.getenv(name, "").strip()
            if v:
                p = os.path.abspath(v)
                candidates.extend([p, os.path.join(p, "T0002", "hq_cache")])
        for name in ("TDX_HY_CFG", "TDX_CACHE_DIR"):
            v = os.getenv(name, "").strip()
            if v:
                p = os.path.abspath(v)
                candidates.append(os.path.dirname(p) if os.path.isfile(p) else p)

        here = os.path.abspath(os.getcwd())
        exe = os.path.dirname(os.path.abspath(sys.executable))
        script = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else here
        candidates.extend([here, os.path.join(here, "T0002", "hq_cache"),
                           exe, os.path.join(exe, "T0002", "hq_cache"),
                           script, os.path.join(script, "T0002", "hq_cache")])

        # Windows常见通达信目录；只检查有限候选，不做全盘递归。
        for drive in ("C:", "D:", "E:", "F:"):
            for root in ("new_tdx", "通达信", "TongDaXin", "Tdx", "tdx", "证券之星"):
                candidates.append(os.path.join(drive, root, "T0002", "hq_cache"))

        out=[]; seen=set()
        for x in candidates:
            x=os.path.normpath(x)
            if x in seen: continue
            seen.add(x)
            if os.path.isdir(x): out.append(x)
        return out

    @staticmethod
    def _read_tdx_industry_cfg(path):
        mapping={}
        try:
            with open(path, "rb") as f:
                text=f.read().decode("gbk", errors="replace")
            for line in text.splitlines():
                parts=line.strip().split("|")
                if len(parts)>=3:
                    code=str(parts[1]).strip().zfill(6)
                    hy=str(parts[2]).strip()
                    if len(code)==6 and hy:
                        mapping[code]=hy
        except Exception as e:
            logger.debug(f"V7.7.2 tdxhy.cfg读取失败: {e}")
        return mapping

    @staticmethod
    def _read_tdx_industry_names(cache_dir):
        names={}
        # tdxzs.cfg：第1列名称，第3列类别；类别2=通达信行业，第6列为行业代码。
        path=os.path.join(cache_dir, "tdxzs.cfg")
        try:
            with open(path, "rb") as f:
                text=f.read().decode("gbk", errors="replace")
            for line in text.splitlines():
                p=line.strip().split("|")
                if len(p)>=6 and p[2].strip()=="2":
                    name=p[0].strip(); code=p[5].strip()
                    if name and code: names[code]=name
        except Exception:
            pass

        # incon.dat：#TDXNHY段保存T代码->行业名称。
        path=os.path.join(cache_dir, "incon.dat")
        try:
            with open(path, "rb") as f:
                text=f.read().decode("gbk", errors="replace")
            in_tdx=False
            for line in text.splitlines():
                x=line.strip()
                if x=="#TDXNHY":
                    in_tdx=True; continue
                if in_tdx and x=="######":
                    break
                if in_tdx and "|" in x:
                    p=x.split("|",1)
                    if len(p)==2 and p[0].strip() and p[1].strip():
                        names.setdefault(p[0].strip(), p[1].strip())
        except Exception:
            pass
        return names

    def _load_local_tdx_industry(self):
        """V7.7.2：优先读取用户明确指定的 TDX hq_cache。"""
        """读取股票->通达信行业代码，并尽量解析行业名称。"""
        if not V77_INDUSTRY_CFG_ENABLED:
            return {}, {}, "TDX_INDUSTRY_DISABLED"
        explicit_dirs = [V77_TDX_CACHE_DIR, os.path.dirname(V77_TDX_HY_CFG)]
        candidate_dirs = []
        for d in explicit_dirs + self._candidate_tdx_cache_dirs():
            if d and d not in candidate_dirs:
                candidate_dirs.append(d)
        for cache_dir in candidate_dirs:
            cfg=os.path.join(cache_dir, V77_INDUSTRY_CFG_NAME)
            if not os.path.isfile(cfg):
                continue
            code_map=self._read_tdx_industry_cfg(cfg)
            if not code_map:
                continue
            names=self._read_tdx_industry_names(cache_dir)
            name_map={}
            for code, hy in code_map.items():
                name=names.get(hy)
                if not name:
                    # 精确名称缺失时逐级回退到父行业；仍无名称才使用代码。
                    probe=hy
                    while not name and len(probe)>3:
                        probe=probe[:-2]
                        name=names.get(probe)
                name_map[code]=name or hy
            logger.info(f"V7.8.1真实行业：tdxhy.cfg={cfg} | 股票映射={len(code_map)} | 行业名称={len(set(name_map.values()))}")
            logger.info(
                f"V7.8行业解析诊断：TDX_CACHE_DIR={V77_TDX_CACHE_DIR} | "
                f"tdxhy.cfg={'FOUND' if os.path.isfile(V77_TDX_HY_CFG) else 'NOT_FOUND'} | "
                f"tdxzs.cfg={'FOUND' if os.path.isfile(V77_TDX_INDUSTRY_NAMES[0]) else 'NOT_FOUND'} | "
                f"incon.dat={'FOUND' if os.path.isfile(V77_TDX_INDUSTRY_NAMES[1]) else 'NOT_FOUND'} | "
                f"行业映射={len(code_map)} | 行业名称={len(set(name_map.values()))}"
            )
            for test_code, test_name in (("601898", "中煤能源"), ("600188", "兖矿能源"), ("000400", "许继电气")):
                logger.info(f"V7.8行业测试股票：{test_code} {test_name} -> {name_map.get(test_code, code_map.get(test_code, 'UNKNOWN'))}")
            return code_map, name_map, "TDX_INDUSTRY_TDXHY_CFG"
        return {}, {}, "TDX_INDUSTRY_UNKNOWN"

    def _update_v78_industry_behavior(self, stats):
        """V7.8.1 行业行为引擎。

        与V7.8的关键区别：
        1. 行业静态行情缓存与行为时间序列彻底分离。
        2. 每60秒采样一次，不再因为120秒静态缓存而永远只有1个样本。
        3. 趋势使用真实时间窗口：30M / 60M。
        4. 加速度比较“最近10M斜率”和“此前10M斜率”。
        5. 退潮使用当日历史峰值，开盘新的一天自动清空。
        6. 内部分化同时观察STD、10%-90%价差、上涨家数扩散变化。
        7. 只使用已经看到的历史快照，不使用未来数据。
        """
        if not V77_INDUSTRY_BEHAVIOR_ENABLED or not stats:
            return stats

        now = time.time()
        today = datetime.datetime.now().date().isoformat()
        if self._industry_behavior_date != today:
            self._industry_behavior_history = {}
            self._industry_behavior_date = today
            self._industry_behavior_last_sample = 0.0

        # 行为历史独立于行业行情缓存；60秒内重复扫描不重复计样。
        should_sample = (now - self._industry_behavior_last_sample) >= V78_BEHAVIOR_SAMPLE_SECONDS

        for sec, item in stats.items():
            ret = safe_float(item.get("mean_change"), np.nan)
            breadth = safe_float(item.get("up_ratio"), np.nan)
            if not (np.isfinite(ret) and np.isfinite(breadth)):
                continue

            hist = self._industry_behavior_history.setdefault(
                sec, deque(maxlen=V78_BEHAVIOR_HISTORY_SIZE)
            )

            if should_sample:
                hist.append({
                    "ts": now,
                    "ret": ret,
                    "breadth": breadth,
                    "dispersion": safe_float(item.get("dispersion"), np.nan),
                    "spread": safe_float(item.get("spread"), np.nan),
                })

            samples = len(hist)
            points = list(hist)

            def value_ago(minutes, field):
                target = now - minutes * 60.0
                candidates = [x for x in points if x["ts"] <= target]
                if not candidates:
                    return np.nan
                return safe_float(candidates[-1].get(field), np.nan)

            ret_30 = value_ago(V78_BEHAVIOR_TREND_30M_MIN, "ret")
            ret_60 = value_ago(V78_BEHAVIOR_TREND_60M_MIN, "ret")
            breadth_30 = value_ago(V78_BEHAVIOR_TREND_30M_MIN, "breadth")

            trend_30 = (ret - ret_30) if np.isfinite(ret_30) else np.nan
            trend_60 = (ret - ret_60) if np.isfinite(ret_60) else np.nan
            breadth_delta_30 = (breadth - breadth_30) if np.isfinite(breadth_30) else np.nan

            # 最近10M与此前10M的行业均值斜率比较。单位最终转换为“%/10M”。
            recent_cut = now - V78_BEHAVIOR_ACCEL_WINDOW_MIN * 60.0
            prior_cut = now - 2 * V78_BEHAVIOR_ACCEL_WINDOW_MIN * 60.0
            recent_pts = [x for x in points if x["ts"] >= recent_cut]
            prior_pts = [x for x in points if prior_cut <= x["ts"] < recent_cut]

            def slope_10m(arr):
                if len(arr) < 2:
                    return np.nan
                dt = arr[-1]["ts"] - arr[0]["ts"]
                if dt <= 0:
                    return np.nan
                return (arr[-1]["ret"] - arr[0]["ret"]) / dt * 600.0

            recent_slope = slope_10m(recent_pts)
            prior_slope = slope_10m(prior_pts)
            accel = (recent_slope - prior_slope) if np.isfinite(recent_slope) and np.isfinite(prior_slope) else np.nan

            peak_ret = max((x["ret"] for x in points), default=ret)
            retreat = max(0.0, peak_ret - ret)

            dispersion = safe_float(item.get("dispersion"), np.nan)
            spread = safe_float(item.get("spread"), np.nan)

            # 行业内部分化：中间涨跌家数区间 + 高离散度，或者上涨扩散明显收缩。
            diverging = (
                ((V78_BEHAVIOR_DIVERGENCE_BREADTH_LOW <= breadth <= V78_BEHAVIOR_DIVERGENCE_BREADTH_HIGH) and
                 ((np.isfinite(dispersion) and dispersion >= V78_BEHAVIOR_DIVERGENCE_STD) or
                  (np.isfinite(spread) and spread >= V78_BEHAVIOR_DIVERGENCE_SPREAD)))
                or (np.isfinite(breadth_delta_30) and breadth_delta_30 <= V78_BEHAVIOR_BREADTH_CONTRACTION and
                    ((np.isfinite(dispersion) and dispersion >= V78_BEHAVIOR_DIVERGENCE_STD) or
                     (np.isfinite(spread) and spread >= V78_BEHAVIOR_DIVERGENCE_SPREAD)))
            )

            if samples < V78_BEHAVIOR_MIN_SAMPLES:
                behavior = "INSUFFICIENT"
            elif retreat >= V78_BEHAVIOR_DUMPING_RETREAT_PCT and np.isfinite(trend_30) and trend_30 < 0:
                behavior = "DUMPING"
            elif retreat >= V78_BEHAVIOR_RETREAT_PCT and np.isfinite(trend_30) and trend_30 < 0:
                behavior = "RETREATING"
            elif diverging:
                behavior = "DIVERGING"
            elif np.isfinite(accel) and accel >= V78_BEHAVIOR_ACCEL_THRESHOLD_PCT_10M:
                behavior = "ACCELERATING"
            elif np.isfinite(accel) and accel <= V78_BEHAVIOR_DECEL_THRESHOLD_PCT_10M:
                behavior = "DECELERATING"
            elif np.isfinite(trend_30) and trend_30 > 0.30 and (not np.isfinite(trend_60) or trend_60 >= 0):
                behavior = "DRIVING"
            elif np.isfinite(trend_30) and trend_30 > 0.20 and np.isfinite(trend_60) and trend_60 < 0:
                behavior = "RECOVERING"
            elif np.isfinite(trend_30) and trend_30 > 0.10:
                behavior = "TREND_UP"
            elif np.isfinite(trend_30) and trend_30 < -0.10:
                behavior = "TREND_DOWN"
            else:
                behavior = "NEUTRAL"

            item.update({
                "behavior_state": behavior,
                "behavior_trend": (
                    "UP" if np.isfinite(trend_30) and trend_30 > 0.10 else
                    ("DOWN" if np.isfinite(trend_30) and trend_30 < -0.10 else "FLAT")
                ),
                "behavior_trend_30m": trend_30,
                "behavior_trend_60m": trend_60,
                "behavior_accel": accel,
                "behavior_retreat": retreat,
                "behavior_dispersion": dispersion,
                "behavior_spread": spread,
                "behavior_breadth_delta": breadth_delta_30,
                "behavior_samples": samples,
                "behavior_peak": peak_ret,
                "behavior_recent_slope": recent_slope,
                "behavior_prior_slope": prior_slope,
                "behavior_sampled": bool(should_sample),
            })

        if should_sample:
            self._industry_behavior_last_sample = now
        self._industry_behavior_cache_time = now
        return stats

    def _calculate_sector_snapshot(self, stock_pool):
        """V7.8.1：行业静态映射与行业行为采样彻底解耦。

        映射/成员列表缓存较长；行情快照每60秒重新读取，
        这样行业行为引擎才能真正形成30M/60M时间序列。
        """
        now = time.time()
        codes = []
        for _, row in stock_pool.iterrows():
            c = str(row.get("代码", "")).replace(".0", "").strip().zfill(6)
            if c:
                codes.append(c)
        codes = list(dict.fromkeys(codes))

        # --------------------------------------------------------------
        # A. 行业映射缓存：5分钟；行为行情不跟着这个缓存停滞。
        # --------------------------------------------------------------
        mapping_valid = (
            bool(self._sector_mapping_cache) and
            now - self._sector_mapping_cache_time < V78_MAPPING_CACHE_SECONDS
        )
        if not mapping_valid:
            sector_map = {}
            sector_members = {}
            source = "TDX_INDUSTRY_UNKNOWN"

            excel_rows = {}
            for _, row in stock_pool.iterrows():
                c = str(row.get("代码", "")).replace(".0", "").strip().zfill(6)
                ind = str(row.get("行业", "") or "").strip()
                if c and ind and ind.lower() not in {"nan", "none", "未分类"}:
                    excel_rows[c] = ind

            if V77_SECTOR_PREFER_EXCEL_GROUP and excel_rows:
                sector_map = dict(excel_rows)
                source = "WATCHLIST_INDUSTRY"
                for c, sec in sector_map.items():
                    sector_members.setdefault(sec, set()).add(c)

            if not sector_map:
                code_hy, code_name, cfg_source = self._load_local_tdx_industry()
                if code_hy:
                    hy_to_name = {}
                    for c, hy in code_hy.items():
                        hy_to_name.setdefault(hy, code_name.get(c, hy))
                    for c in codes:
                        hy = code_hy.get(c)
                        if hy:
                            sec = code_name.get(c, hy_to_name.get(hy, hy))
                            sector_map[c] = sec
                    for c, hy in code_hy.items():
                        sec = hy_to_name.get(hy, code_name.get(c, hy))
                        sector_members.setdefault(sec, set()).add(c)
                    source = cfg_source

            if not sector_map and V77_REAL_SECTOR_ENABLED:
                parsed = self.tdx.get_block_members("block_hy.dat")
                if parsed:
                    lm = {}; members = {}
                    for rec in parsed:
                        name = str(rec.get("blockname", "") or "").strip()
                        c = str(rec.get("code", "") or "").strip().zfill(6)
                        if not name or not c:
                            continue
                        members.setdefault(name, set()).add(c)
                        lm.setdefault(c, name)
                    matched = sum(c in lm for c in codes)
                    if matched >= max(1, min(2, len(codes))):
                        sector_map = {c: lm[c] for c in codes if c in lm}
                        sector_members = members
                        source = "TDX_INDUSTRY_BLOCK_HY"

            self._sector_mapping_cache = sector_map
            self._sector_mapping_members = sector_members
            self._sector_mapping_source = source
            self._sector_mapping_cache_time = now
        else:
            sector_map = self._sector_mapping_cache
            sector_members = self._sector_mapping_members
            source = self._sector_mapping_source

        if not sector_map:
            logger.info(f"V7.8.1真实行业：source={source} | 监测池映射=0/{len(codes)} | 有效行业=0 | 行为历史保持")
            return {}, {}, source

        sectors = sorted(set(sector_map.values()))
        relevant = set()
        for sec in sectors:
            mem = sector_members.get(sec, set())
            if len(mem) >= V77_SECTOR_MIN_MEMBERS:
                # 不再依赖监测池；行业行为必须尽可能覆盖整个行业成员。
                relevant.update(list(mem)[:V77_SECTOR_MAX_QUOTE_MEMBERS])

        # --------------------------------------------------------------
        # B. 行业行情：每60秒重新采样。
        # --------------------------------------------------------------
        quotes = []
        pairs = [(get_market_code(c), c) for c in sorted(relevant)]
        for i in range(0, len(pairs), V77_SECTOR_QUOTE_CHUNK):
            try:
                q = self.tdx.api.get_security_quotes(pairs[i:i + V77_SECTOR_QUOTE_CHUNK])
                if q:
                    quotes.extend(q)
            except Exception as e:
                logger.debug(f"V7.8.1行业行情读取失败: {e}")

        c2s = {}
        for sec in sectors:
            for c in sector_members.get(sec, set()):
                c2s[c] = sec

        grouped = {sec: [] for sec in sectors}
        for q in quotes:
            c = str(q.get("code", "")).strip().zfill(6)
            sec = c2s.get(c)
            if not sec:
                continue
            price = safe_float(q.get("price"), np.nan)
            prev = safe_float(q.get("last_close"), np.nan)
            if np.isfinite(price) and np.isfinite(prev) and prev > 0:
                grouped[sec].append((price / prev - 1) * 100)

        stats = {}
        for sec, arr in grouped.items():
            if len(arr) < V77_SECTOR_MIN_MEMBERS:
                continue
            a = np.asarray(arr, float)
            q10 = float(np.percentile(a, 10))
            q90 = float(np.percentile(a, 90))
            stats[sec] = {
                "members": len(a),
                "mean_change": float(a.mean()),
                "up_ratio": float(np.mean(a > 0)),
                "source": source,
                "dispersion": float(np.std(a)),
                "spread": q90 - q10,
            }

        stats = self._update_v78_industry_behavior(stats)
        self._sector_map_cache = sector_map
        self._sector_stats_cache = stats
        self._sector_source_cache = source
        self._sector_cache_time = now
        sampled = any(v.get("behavior_sampled") for v in stats.values()) if stats else False
        logger.info(
            f"V7.8.1真实行业：source={source} | 监测池映射={sum(c in sector_map for c in codes)}/{len(codes)} "
            f"| 有效行业={len(stats)} | 行情成员={sum(v['members'] for v in stats.values())} "
            f"| 行为采样={'YES' if sampled else 'NO'} | 历史分钟={min([len(h) for h in self._industry_behavior_history.values()] or [0])}-{max([len(h) for h in self._industry_behavior_history.values()] or [0])}"
        )
        return sector_map, stats, source

    def _calculate_real_breadth(self, stock_pool=None):
        """V7.5.4兼容入口：使用TDX原生880005/指数涨跌家数。"""
        return self._get_tdx_market_breadth(stock_pool)

    def _analyze_one_index(self, name, code, market):
        """三指数统一计算：实时价、上一交易日收盘、当日涨跌、20日压力。"""
        base = {"name": name, "code": code, "market": market, "valid": False,
                "current": np.nan, "previous_close": np.nan, "change": np.nan,
                "pressure": np.nan, "pressure_distance": np.nan, "ma20": np.nan,
                "intraday_change": np.nan}
        day = self.tdx.get_index_bars(CATEGORY_DAY, market, code, 0, 120)
        if day is None or day.empty or "close" not in day.columns:
            return base
        d = day.copy()
        for c in ("close", "high"):
            if c in d.columns:
                d[c] = pd.to_numeric(d[c], errors="coerce")
        d = d.dropna(subset=["close"]).sort_values("datetime" if "datetime" in d.columns else d.index).reset_index(drop=True)
        if len(d) < 21:
            return base

        intra = self.tdx.get_index_bars(CATEGORY_1M, market, code, 0, 240)
        ic = pd.Series(dtype=float)
        if intra is not None and not intra.empty and "close" in intra.columns:
            ic = pd.to_numeric(intra["close"], errors="coerce").dropna()
        quote_current = np.nan
        try:
            q = self.tdx.api.get_security_quotes([(market, code)])
            if q:
                quote_current = safe_float(q[0].get("price"), np.nan)
        except Exception:
            pass
        current = quote_current if np.isfinite(quote_current) and quote_current > 0 else (safe_float(ic.iloc[-1], np.nan) if len(ic) else safe_float(d["close"].iloc[-1], np.nan))

        # 上一交易日收盘：日K最后日期若为今天，取倒数第二根；否则取最新日K。
        previous_close = safe_float(d["close"].iloc[-1], np.nan)
        if "datetime" in d.columns:
            dd = pd.to_datetime(d["datetime"], errors="coerce")
            last_date = dd.iloc[-1].date() if len(dd) and pd.notna(dd.iloc[-1]) else None
            if last_date == datetime.datetime.now().date() and len(d) >= 2:
                previous_close = safe_float(d["close"].iloc[-2], np.nan)

        change = ((current - previous_close) / previous_close * 100.0
                  if np.isfinite(current) and np.isfinite(previous_close) and previous_close > 0 else np.nan)
        intraday_change = ((ic.iloc[-1] / ic.iloc[-2] - 1.0) * 100.0
                           if len(ic) >= 2 and ic.iloc[-2] > 0 else np.nan)

        recent = d.tail(21).iloc[:-1]
        pressure = safe_float(recent["high"].max() if "high" in recent.columns else recent["close"].max(), current)
        pressure = max(pressure, current)
        pressure_distance = max(0.0, min(100.0, safe_float((pressure-current)/current*100.0, 0.0)))
        ma20 = safe_float(recent["close"].mean(), current)
        base.update(valid=np.isfinite(current), current=current, previous_close=previous_close,
                    change=change, pressure=pressure, pressure_distance=pressure_distance,
                    ma20=ma20, intraday_change=intraday_change)
        return base

    def analyze(self, stock_df=None, stock_pool=None):
        indices = {}
        for key, cfg in V751_INDEX_CONFIG.items():
            indices[key] = self._analyze_one_index(cfg["name"], cfg["code"], cfg["market"])

        valid_idx = [v for v in indices.values() if v.get("valid") and np.isfinite(v.get("change", np.nan))]
        base = {
            "valid": False, "state": "指数数据不足", "state5": "E", "score": 0.0,
            "pressure": np.nan, "pressure_distance": np.nan, "current": np.nan,
            "idx_change": np.nan, "idx_day_change": np.nan, "idx_intraday_change": np.nan,
            "breadth": {}, "indices": indices, "sector_stats": {},
        }
        if len(valid_idx) < 3:
            logger.warning(f"V7.5.6三指数有效数据不足：{len(valid_idx)}/3")
            return base

        sh = indices["SH"]
        # 市场压力仍以上证为主，但指数环境由三指数共同决定。
        pressure = sh["pressure"]; pressure_distance = sh["pressure_distance"]
        avg_change = float(np.mean([x["change"] for x in valid_idx]))
        avg_intraday = float(np.mean([x["intraday_change"] for x in valid_idx if np.isfinite(x["intraday_change"])]) or 0.0)

        breadth = self._get_tdx_market_breadth(stock_pool)
        sector_map, sector_stats, sector_source = self._calculate_sector_snapshot(stock_pool)
        up_ratio = safe_float(breadth.get("up_ratio"), 0.5)
        strong_up_ratio = safe_float(breadth.get("strong_up_ratio"), np.nan)
        strong_down_ratio = safe_float(breadth.get("strong_down_ratio"), np.nan)

        # V7.6：广度单独分层。
        # 这样“指数偏弱但上涨家数61.5%”会被识别为结构性市场，
        # 而不是简单地认为整个市场都是空头。
        if breadth.get("valid"):
            if up_ratio >= V76_BREADTH_BULLISH_RATIO:
                breadth_state = "BULLISH_BREADTH"
            elif up_ratio >= V76_BREADTH_NEUTRAL_RATIO:
                breadth_state = "NEUTRAL_BREADTH"
            else:
                breadth_state = "BEARISH_BREADTH"
        else:
            breadth_state = "UNKNOWN_BREADTH"

        score = 0.0
        # 三指数方向：至少2/3同向才给予明显加减分。
        pos = sum(x["change"] >= 0.30 for x in valid_idx)
        neg = sum(x["change"] <= -0.30 for x in valid_idx)
        if pos >= 2: score += 3.0
        elif neg >= 2: score -= 3.0
        if avg_change >= 1.0: score += 2.0
        elif avg_change <= -1.0: score -= 2.0
        if pressure_distance < 0.5: score -= 2.0
        elif pressure_distance >= 3.0: score += 2.0
        if breadth.get("valid"):
            if up_ratio >= V75_BREADTH_A: score += 4.0
            elif up_ratio >= V75_BREADTH_B: score += 2.0
            elif up_ratio < V75_BREADTH_E: score -= 5.0
            elif up_ratio < V75_BREADTH_D: score -= 3.0
            # V7.5.3不再从逐股报价虚构“强势股占比”。
            # 只有真实强势比例存在时才参与该项评分。
            if np.isfinite(strong_up_ratio) and strong_up_ratio >= 0.20:
                score += 2.0
            elif np.isfinite(strong_down_ratio) and strong_down_ratio >= 0.20:
                score -= 2.0
        else:
            # 真广度不可用时不虚构；直接扣除“广度确认分”。
            score -= 1.0

        score = max(-15.0, min(15.0, score))
        if score >= V75_MARKET_A_MIN_SCORE: state5, text = "A", "🟢A级强势进攻"
        elif score >= V75_MARKET_B_MIN_SCORE: state5, text = "B", "🟢B级结构性进攻"
        elif score >= V75_MARKET_D_MIN_SCORE: state5, text = "C", "🟡C级震荡观察"
        elif score >= V75_MARKET_E_MIN_SCORE: state5, text = "D", "🟠D级弱势防守"
        else: state5, text = "E", "🔴E级风险模式"

        return {
            "valid": True, "state": text, "state5": state5, "score": float(score),
            "pressure": pressure, "pressure_distance": pressure_distance,
            "current": sh["current"], "ma20": sh["ma20"],
            "idx_change": sh["change"], "idx_day_change": sh["change"],
            "idx_intraday_change": sh["intraday_change"],
            "indices": indices, "avg_index_change": avg_change,
            "avg_index_intraday_change": avg_intraday,
            "breadth": breadth, "sector_stats": sector_stats,
            "sector_map": sector_map, "sector_source": sector_source,
            "breadth_state": breadth_state,
            "rs_score": 0, "source": "3INDEX_REALTIME+TDX_NATIVE_880005_BREADTH",
        }


# ======================================================================
# 个股趋势分层 V1
# ======================================================================
def _completed_stock_bars(df, period):
    """只保留已完成的日/周/月K，避免当前未完成K线污染趋势判断。"""
    if df is None or df.empty or "close" not in df.columns:
        return df
    d = df.copy()
    if "datetime" not in d.columns:
        return d
    d["datetime"] = pd.to_datetime(d["datetime"], errors="coerce")
    d = d.dropna(subset=["datetime"]).sort_values("datetime").reset_index(drop=True)
    if d.empty:
        return d
    now = datetime.datetime.now()
    last = d["datetime"].iloc[-1]
    remove = False
    if period == "D":
        remove = last.date() == now.date()
    elif period == "W":
        remove = last.isocalendar()[:2] == now.isocalendar()[:2]
    elif period == "M":
        remove = last.year == now.year and last.month == now.month
    if remove and len(d) > 1:
        d = d.iloc[:-1].copy()
    return d.reset_index(drop=True)


def _trend_ma_state(df, ma):
    """返回 MA 当前斜率：UP / DOWN / FLAT。"""
    if df is None or len(df) < ma + 3 or "close" not in df.columns:
        return "UNKNOWN", np.nan, np.nan
    c = pd.to_numeric(df["close"], errors="coerce").dropna()
    if len(c) < ma + 3:
        return "UNKNOWN", np.nan, np.nan
    ma_now = safe_float(c.rolling(ma).mean().iloc[-1])
    ma_prev = safe_float(c.rolling(ma).mean().iloc[-3])
    if not np.isfinite(ma_now) or not np.isfinite(ma_prev) or ma_prev == 0:
        return "UNKNOWN", ma_now, ma_prev
    pct = (ma_now - ma_prev) / abs(ma_prev) * 100.0
    if pct > TREND_MA_TOLERANCE_PCT:
        state = "UP"
    elif pct < -TREND_MA_TOLERANCE_PCT:
        state = "DOWN"
    else:
        state = "FLAT"
    return state, ma_now, ma_prev


def _trend_price_state(price, ma):
    if not np.isfinite(price) or not np.isfinite(ma) or ma == 0:
        return "UNKNOWN"
    pct = (price - ma) / abs(ma) * 100.0
    if pct > TREND_MA_TOLERANCE_PCT:
        return "ABOVE"
    if pct < -TREND_MA_TOLERANCE_PCT:
        return "BELOW"
    return "NEAR"


def calculate_stock_trend_layer(tdx_pool, market, code):
    """月20 + 周20/60 + 日20/60 -> A/B/C/D/E + score + gate。"""
    result = {
        "valid": False,
        "grade": "UNKNOWN",
        "score": 0,
        "gate": "WAIT",
        "reason": "趋势数据不足",
        "price": np.nan,
        "monthly": {},
        "weekly": {},
        "daily": {},
    }
    if not TREND_LAYER_ENABLED:
        result.update(valid=True, grade="C", score=0, gate="PASS", reason="趋势层关闭")
        return result

    try:
        df_d = tdx_pool.get_bars(4, market, code, 0, TREND_FETCH_BARS)
        df_w = tdx_pool.get_bars(5, market, code, 0, TREND_FETCH_BARS)
        df_m = tdx_pool.get_bars(6, market, code, 0, TREND_FETCH_BARS)

        df_d = _completed_stock_bars(df_d, "D")
        df_w = _completed_stock_bars(df_w, "W")
        df_m = _completed_stock_bars(df_m, "M")

        def pack(df, mas, period):
            x = {"period": period, "valid": False, "states": {}, "price_states": {}, "values": {}}
            if df is None or df.empty or "close" not in df.columns:
                return x
            c = pd.to_numeric(df["close"], errors="coerce").dropna()
            if len(c) < max(mas) + 3:
                return x
            price = safe_float(c.iloc[-1])
            x["valid"] = np.isfinite(price)
            x["close"] = price
            for ma in mas:
                st, now_ma, prev_ma = _trend_ma_state(df, ma)
                x["states"][f"MA{ma}"] = st
                x["values"][f"MA{ma}"] = now_ma
                x["values"][f"MA{ma}_prev"] = prev_ma
                x["price_states"][f"MA{ma}"] = _trend_price_state(price, now_ma)
            return x

        result["daily"] = pack(df_d, [20, 60], "D")
        result["weekly"] = pack(df_w, [20, 60], "W")
        result["monthly"] = pack(df_m, [20], "M")

        if not (result["daily"]["valid"] and result["weekly"]["valid"] and result["monthly"]["valid"]):
            return result

        # 5条MA斜率：月20、周20/60、日20/60
        states = [
            result["monthly"]["states"]["MA20"],
            result["weekly"]["states"]["MA20"],
            result["weekly"]["states"]["MA60"],
            result["daily"]["states"]["MA20"],
            result["daily"]["states"]["MA60"],
        ]
        up = states.count("UP")
        down = states.count("DOWN")

        # 价格是否位于关键均线之上
        above_count = sum([
            result["monthly"]["price_states"]["MA20"] == "ABOVE",
            result["weekly"]["price_states"]["MA20"] == "ABOVE",
            result["weekly"]["price_states"]["MA60"] == "ABOVE",
            result["daily"]["price_states"]["MA20"] == "ABOVE",
            result["daily"]["price_states"]["MA60"] == "ABOVE",
        ])
        below_count = sum([
            result["monthly"]["price_states"]["MA20"] == "BELOW",
            result["weekly"]["price_states"]["MA20"] == "BELOW",
            result["weekly"]["price_states"]["MA60"] == "BELOW",
            result["daily"]["price_states"]["MA20"] == "BELOW",
            result["daily"]["price_states"]["MA60"] == "BELOW",
        ])

        # 分层：A强多、B偏多、C震荡、D偏空、E空头加速
        if down >= 4 or (down >= 3 and result["monthly"]["states"]["MA20"] == "DOWN"):
            grade, score, gate = "E", TREND_SCORE_E, "BLOCK"
            reason = f"5条核心MA中{down}条向下，月K MA20亦偏弱，趋势空头"
        elif down >= 3:
            grade, score, gate = "D", TREND_SCORE_D, "BLOCK"
            reason = f"5条核心MA中{down}条向下，趋势偏空"
        elif up >= 4 and above_count >= 3:
            grade, score, gate = "A", TREND_SCORE_A, "PASS"
            reason = f"5条核心MA中{up}条向上，价格站上{above_count}条关键MA，趋势强多"
        elif up >= 3 and above_count >= 2:
            grade, score, gate = "B", TREND_SCORE_B, "PASS"
            reason = f"5条核心MA中{up}条向上，趋势偏多"
        else:
            grade, score, gate = "C", TREND_SCORE_C, "WAIT"
            reason = f"MA方向多空分化：向上{up}条、向下{down}条，趋势震荡"

        # 防止“斜率向上但价格同时跌破日20/60、周20”的假强势
        critical_below = (
            result["daily"]["price_states"]["MA20"] == "BELOW"
            and result["daily"]["price_states"]["MA60"] == "BELOW"
            and result["weekly"]["price_states"]["MA20"] == "BELOW"
        )
        if grade in {"A", "B"} and critical_below:
            grade, score, gate = "C", TREND_SCORE_C, "WAIT"
            reason = "均线斜率尚可，但价格同时跌破日MA20/MA60及周MA20，降级观察"

        result.update(valid=True, grade=grade, score=score, gate=gate, reason=reason)
        return result
    except Exception as e:
        result["reason"] = f"趋势计算异常: {e}"
        return result


# ======================================================================
# ENTRY_GATE V6.3
# ======================================================================
def calculate_entry_gate(

    df_1m,
    market_info,
    dar_info,
    current_price
):
    """
    ENTRY_GATE V6.2

    在V6.0基础上重点修复两个问题：

    1. “假突破”不能只看最近8根1M K线。
       000973今天就是：先突破12.24 → 冲到12.40 → 后面才跌回12.24。
       如果只看最后8根K线，可能把这次突破事件漏掉。

    2. 未发生突破时，突破量能不能显示“通过”。
       V6.2改为三态：
           NOT_TRIGGERED = 未触发
           PASS          = 通过
           BLOCK         = 不足

    设计原则：
        - 不修改原有评分。
        - 不删除★★★★。
        - ENTRY_GATE主要决定★★★★★是否允许成立。
        - 当★★★★被ENTRY_GATE阻断时，交易建议必须明确写成“观察/不主动开仓”。
        - 压力计算只使用当前信号之前的已完成K线，避免未来函数。
        - 当日突破事件使用“滚动前N根压力”逐根计算，再判断后续是否失败。
    """

    result = {
        "enabled": ENTRY_GATE_ENABLED,
        "valid": False,
        "passed": True,
        "state": "DISABLED",
        "reason": "ENTRY_GATE关闭",

        "pressure": np.nan,
        "distance_to_pressure": np.nan,

        "breakout": False,
        "breakout_confirmed": False,
        "breakout_failed": False,

        # V6.2 新增：当日突破状态
        "today_breakout": False,
        "today_breakout_failed": False,
        "today_breakout_pressure": np.nan,
        "today_breakout_price": np.nan,
        "today_breakout_time": "",
        "today_high": np.nan,

        # V6.2：量能三态
        "volume_ok": None,
        "volume_state": "NOT_TRIGGERED",

        "market_ok": True,
        "intraday_change": np.nan,
        "relative_strength": np.nan,
        "reasons": []
    }

    if not ENTRY_GATE_ENABLED:
        return result

    if (
        df_1m is None
        or df_1m.empty
        or "close" not in df_1m.columns
        or "high" not in df_1m.columns
    ):
        result.update({
            "passed": False,
            "state": "BLOCK",
            "reason": "1M数据不足，ENTRY_GATE禁止五星"
        })
        result["reasons"].append("1M数据不足")
        return result

    d = df_1m.copy()

    for c in ["open", "high", "low", "close", "volume"]:
        if c in d.columns:
            d[c] = pd.to_numeric(
                d[c],
                errors="coerce"
            )

    if "datetime" in d.columns:
        d["datetime"] = pd.to_datetime(
            d["datetime"],
            errors="coerce"
        )
        d = d.dropna(
            subset=["datetime"]
        ).sort_values(
            "datetime"
        )

    d = d.dropna(
        subset=["close", "high"]
    ).reset_index(drop=True)

    if len(d) < 12:
        result.update({
            "passed": False,
            "state": "BLOCK",
            "reason": "1M有效K线不足，ENTRY_GATE禁止五星"
        })
        result["reasons"].append("有效1M K线不足")
        return result

    current = safe_float(
        current_price,
        safe_float(d["close"].iloc[-1])
    )

    if not np.isfinite(current) or current <= 0:
        result.update({
            "passed": False,
            "state": "BLOCK",
            "reason": "当前价格异常，ENTRY_GATE禁止五星"
        })
        result["reasons"].append("当前价格异常")
        return result

    # --------------------------------------------------------------
    # 当前交易日
    # --------------------------------------------------------------
    if "datetime" in d.columns and not d.empty:
        latest_date = d["datetime"].dt.date.max()
        day_df = d[
            d["datetime"].dt.date == latest_date
        ].copy().reset_index(drop=True)
    else:
        day_df = d.copy().reset_index(drop=True)

    if len(day_df) < 3:
        result.update({
            "passed": False,
            "state": "BLOCK",
            "reason": "当前交易日1M数据不足，ENTRY_GATE禁止五星"
        })
        result["reasons"].append("当前交易日1M数据不足")
        return result

    result["today_high"] = safe_float(
        day_df["high"].max(),
        np.nan
    )

    # --------------------------------------------------------------
    # 只用“当前K线之前”的已完成K线。
    # --------------------------------------------------------------
    completed = d.iloc[:-1].copy().reset_index(drop=True)

    if len(completed) < 10:
        result.update({
            "passed": False,
            "state": "BLOCK",
            "reason": "已完成1M K线不足，ENTRY_GATE禁止五星"
        })
        result["reasons"].append("已完成1M K线不足")
        return result

    lookback = min(
        int(ENTRY_GATE_LOOKBACK),
        len(completed)
    )

    # --------------------------------------------------------------
    # 1. 当前局部压力
    # --------------------------------------------------------------
    pressure_df = completed.tail(lookback)

    pressure = safe_float(
        pressure_df["high"].max(),
        np.nan
    )

    if not np.isfinite(pressure) or pressure <= 0:
        result.update({
            "passed": False,
            "state": "BLOCK",
            "reason": "无法计算局部平台压力，ENTRY_GATE禁止五星"
        })
        result["reasons"].append("压力位数据异常")
        return result

    result["pressure"] = pressure

    # --------------------------------------------------------------
    # 2. 当前价格距离压力
    # --------------------------------------------------------------
    distance = (
        pressure - current
    ) / current * 100.0

    result["distance_to_pressure"] = distance

    # --------------------------------------------------------------
    # 3. 当前是否已经突破当前局部压力
    # --------------------------------------------------------------
    breakout_level = pressure * (
        1.0 + ENTRY_GATE_BREAKOUT_BUFFER_PCT / 100.0
    )

    breakout = current >= breakout_level
    result["breakout"] = bool(breakout)

    # --------------------------------------------------------------
    # 4. 当前有效站稳
    # --------------------------------------------------------------
    confirm_n = max(
        1,
        int(ENTRY_GATE_CONFIRM_BARS)
    )

    confirm_df = completed.tail(confirm_n)
    breakout_confirmed = False

    if len(confirm_df) >= confirm_n:
        closes = pd.to_numeric(
            confirm_df["close"],
            errors="coerce"
        ).dropna()

        if len(closes) >= confirm_n:
            breakout_confirmed = bool(
                (
                    closes
                    >= breakout_level
                ).all()
            )

    result["breakout_confirmed"] = breakout_confirmed

    # --------------------------------------------------------------
    # 5. V6.2：逐根计算“当时的前N根压力”，寻找当日突破事件
    # --------------------------------------------------------------
    # 关键：不能在收盘后直接用“最后30根最高价”倒推突破。
    # 必须逐根使用过去数据：
    #     prior_pressure[i] = i之前N根最高价
    # 这样不会产生未来函数。
    # --------------------------------------------------------------

    day_completed = completed.copy()

    if "datetime" in day_completed.columns:
        latest_date = day_completed["datetime"].dt.date.max()
        day_completed = day_completed[
            day_completed["datetime"].dt.date == latest_date
        ].copy().reset_index(drop=True)

    day_breakout_events = []

    if len(day_completed) >= 10:

        lb = min(
            int(ENTRY_GATE_LOOKBACK),
            len(day_completed) - 1
        )

        highs = pd.to_numeric(
            day_completed["high"],
            errors="coerce"
        )

        for i in range(lb, len(day_completed)):

            prior_window = highs.iloc[
                max(0, i - lb):i
            ].dropna()

            if prior_window.empty:
                continue

            prior_pressure = safe_float(
                prior_window.max(),
                np.nan
            )

            if (
                not np.isfinite(prior_pressure)
                or prior_pressure <= 0
            ):
                continue

            event_level = prior_pressure * (
                1.0
                +
                ENTRY_GATE_DAY_BREAKOUT_BUFFER_PCT
                / 100.0
            )

            bar_high = safe_float(
                day_completed["high"].iloc[i],
                np.nan
            )

            if (
                np.isfinite(bar_high)
                and
                bar_high >= event_level
            ):
                event_time = ""

                if "datetime" in day_completed.columns:
                    event_time = str(
                        day_completed["datetime"].iloc[i]
                    )

                day_breakout_events.append({
                    "index": i,
                    "pressure": prior_pressure,
                    "breakout_level": event_level,
                    "high": bar_high,
                    "time": event_time
                })

    # --------------------------------------------------------------
    # 6. V6.2：从“最后一个突破事件”判断是否失败
    # --------------------------------------------------------------
    # 选择当天最近一次有效突破事件。
    # 如果之后任何已完成K线收盘重新跌回该事件压力下方，
    # 且当前仍未重新站回压力，则视为“当日突破失败”。
    # --------------------------------------------------------------

    today_breakout = False
    today_breakout_failed = False
    latest_event = None

    if day_breakout_events:

        latest_event = day_breakout_events[-1]
        today_breakout = True

        event_idx = latest_event["index"]
        event_pressure = latest_event["pressure"]

        after_event = day_completed.iloc[
            event_idx + 1:
        ].copy()

        current_below_event_pressure = (
            current
            <
            event_pressure
            *
            (
                1.0
                -
                ENTRY_GATE_DAY_FAILURE_RETRACE_PCT
                / 100.0
            )
        )

        # 必须存在“突破以后”的后续K线，才能叫失败。
        if (
            not after_event.empty
            and
            current_below_event_pressure
        ):
            after_closes = pd.to_numeric(
                after_event["close"],
                errors="coerce"
            ).dropna()

            if not after_closes.empty:
                # 至少有一根后续K线收盘跌回压力下方。
                today_breakout_failed = bool(
                    (
                        after_closes
                        <
                        event_pressure
                        *
                        (
                            1.0
                            -
                            ENTRY_GATE_DAY_FAILURE_RETRACE_PCT
                            / 100.0
                        )
                    ).any()
                )

        result["today_breakout_pressure"] = event_pressure
        result["today_breakout_price"] = latest_event["high"]
        result["today_breakout_time"] = latest_event["time"]

    result["today_breakout"] = today_breakout
    result["today_breakout_failed"] = today_breakout_failed

    # --------------------------------------------------------------
    # 7. 兼容旧字段 breakout_failed
    # --------------------------------------------------------------
    # V6.2之后，breakout_failed优先代表“当日突破失败”。
    # 同时保留最近窗口检查作为辅助。
    # --------------------------------------------------------------

    failure_n = min(
        int(ENTRY_GATE_FAILURE_LOOKBACK),
        len(completed)
    )

    failure_df = completed.tail(
        failure_n
    )

    failure_break_level = pressure * (
        1.0 + ENTRY_GATE_FAILURE_BUFFER_PCT / 100.0
    )

    recent_high = safe_float(
        failure_df["high"].max(),
        pressure
    )

    recent_window_failed = bool(
        recent_high >= failure_break_level
        and
        current < pressure
    )

    breakout_failed = bool(
        today_breakout_failed
        or
        recent_window_failed
    )

    result["breakout_failed"] = breakout_failed

    # --------------------------------------------------------------
    # 8. V6.2：突破量能三态
    # --------------------------------------------------------------
    # 未突破：NOT_TRIGGERED
    # 已突破且有量：PASS
    # 已突破但量能不足：BLOCK
    # --------------------------------------------------------------

    volume_ok = None
    volume_state = "NOT_TRIGGERED"

    if today_breakout or breakout:

        vol_ratio = safe_float(
            dar_info.get("vol_ratio"),
            np.nan
        ) if isinstance(dar_info, dict) else np.nan

        if np.isfinite(vol_ratio):

            volume_ok = bool(
                vol_ratio >= ENTRY_GATE_BREAKOUT_VOL_MIN
            )

            volume_state = (
                "PASS"
                if volume_ok
                else "BLOCK"
            )

        else:

            volume_ok = False
            volume_state = "BLOCK"

    result["volume_ok"] = volume_ok
    result["volume_state"] = volume_state

    # --------------------------------------------------------------
    # 9. 个股日内相对强弱
    # --------------------------------------------------------------
    intraday_change = np.nan

    if "close" in d.columns and len(d) >= 2:

        # 用当天第一根1M收盘作为日内基准。
        first_day_close = safe_float(
            day_df["close"].iloc[0],
            np.nan
        )

        if (
            np.isfinite(first_day_close)
            and
            first_day_close > 0
        ):
            intraday_change = (
                current - first_day_close
            ) / first_day_close * 100.0

    result["intraday_change"] = intraday_change

    index_change = safe_float(
        market_info.get("idx_change"),
        0.0
    ) if isinstance(market_info, dict) else 0.0

    if np.isfinite(intraday_change):
        result["relative_strength"] = (
            intraday_change - index_change
        )

    # --------------------------------------------------------------
    # 10. 市场环境
    # --------------------------------------------------------------
    market_score = safe_float(
        market_info.get("score"),
        0.0
    ) if isinstance(market_info, dict) else 0.0

    market_ok = market_score > ENTRY_GATE_MARKET_SCORE_BLOCK
    result["market_ok"] = market_ok

    # --------------------------------------------------------------
    # 11. 最终闸门
    # --------------------------------------------------------------
    reasons = []

    # A：当日突破失败是最高优先级阻断理由。
    if (
        ENTRY_GATE_DAY_FAILURE_BLOCK
        and
        today_breakout_failed
    ):
        event_pressure = safe_float(
            result["today_breakout_pressure"],
            pressure
        )
        event_high = safe_float(
            result["today_breakout_price"],
            np.nan
        )

        if np.isfinite(event_high):
            reasons.append(
                f"今日突破失败：先突破{event_pressure:.2f}"
                f"并冲至{event_high:.2f}，随后重新跌回压力下方"
            )
        else:
            reasons.append(
                f"今日突破失败：突破{event_pressure:.2f}后重新跌回压力下方"
            )

    # B：刚突破但没有连续站稳。
    if (
        breakout
        and
        not breakout_confirmed
        and
        not today_breakout_failed
    ):
        reasons.append(
            f"刚突破压力{pressure:.2f}，"
            f"尚未连续{confirm_n}根已完成1M K线站稳"
        )

    # C：贴压力位，且还没有真正突破。
    if not breakout and distance >= 0:
        if distance < ENTRY_GATE_NEAR_PRESSURE_PCT:
            reasons.append(
                f"距离局部压力仅{distance:.2f}%，"
                f"不足{ENTRY_GATE_NEAR_PRESSURE_PCT:.2f}%"
            )

    # D：突破发生后量能不足。
    if (
        (today_breakout or breakout)
        and
        not volume_ok
    ):
        reasons.append(
            f"突破量能不足，"
            f"动态量能需≥{ENTRY_GATE_BREAKOUT_VOL_MIN:.2f}倍"
        )

    # E：大盘明显偏弱。
    if not market_ok:
        reasons.append(
            f"大盘环境偏弱，Market Score={market_score:.1f}，"
            f"低于闸门{ENTRY_GATE_MARKET_SCORE_BLOCK}"
        )

    # F：明显弱于指数，且不是已经确认的有效突破。
    relative_strength = result["relative_strength"]

    if (
        np.isfinite(relative_strength)
        and
        relative_strength < -1.0
        and
        not breakout_confirmed
    ):
        reasons.append(
            f"个股日内相对大盘弱{abs(relative_strength):.2f}%"
        )

    # V6.3 Hotfix：ENTRY_GATE 三态化
    # PASS=条件完整；WAIT=等待确认；BLOCK=真正硬阻断
    passed = len(reasons) == 0

    result["passed"] = passed
    result["reasons"] = reasons

    # 接近压力、量能不足、刚突破未确认 -> WAIT
    # 实际突破失败、市场硬弱、明显相对弱势 -> BLOCK
    hard_block = bool(today_breakout_failed)

    if not market_ok:
        hard_block = True

    if (
        np.isfinite(relative_strength)
        and relative_strength < -1.0
        and not breakout_confirmed
    ):
        hard_block = True

    if passed:
        result["valid"] = True
        result["state"] = "PASS"
        result["reason"] = (
            f"✓ ENTRY_GATE通过 | 局部压力={pressure:.2f} | "
            f"压力距离={distance:.2f}%"
        )
    else:
        result["valid"] = True
        result["state"] = "BLOCK" if hard_block else "WAIT"
        prefix = "✗ ENTRY_GATE阻断★★★★★：" if hard_block else "⏳ ENTRY_GATE等待："
        result["reason"] = prefix + "；".join(reasons)


    return result


# ======================================================================\n# ======================================================================
# Capital Behavior V2.0
# ----------------------------------------------------------------------
# 资金行为状态机 + 10/20/30根历史序列 + 转折检测
#
# 核心原则：
#   1. 只使用当前K线及其之前的数据，禁止未来函数。
#   2. 10/20/30根分别描述短、中、长三个行为窗口。
#   3. 不把“主力”当成可直接观测对象，只统计OHLCV可观测代理。
#   4. 重点识别：吸收 → 推升 → 追涨 → 派发，以及反向转折。
# ======================================================================

CAPITAL_BEHAVIOR_ENABLED = True
CAPITAL_BEHAVIOR_VERSION = "V2.0"
CAPITAL_BEHAVIOR_LOOKBACK = 60
CAPITAL_BEHAVIOR_BASELINE = 20
CAPITAL_BEHAVIOR_HIGH_VOL = 1.50
CAPITAL_BEHAVIOR_EXTREME_VOL = 2.50
CAPITAL_BEHAVIOR_ABSORB_DROP_PCT = 0.30
CAPITAL_BEHAVIOR_CHASE_RETURN_PCT = 0.35
CAPITAL_BEHAVIOR_DISTRIBUTION_NEAR_HIGH_PCT = 1.50
CAPITAL_BEHAVIOR_DISTRIBUTION_UPPER_WICK_PCT = 0.35
CAPITAL_BEHAVIOR_WINDOWS = (10, 20, 30)


def _cb_safe(v, default=0.0):
    try:
        x = float(v)
        return x if np.isfinite(x) else default
    except Exception:
        return default


def _cb_empty():
    return {
        "valid": False,
        "version": CAPITAL_BEHAVIOR_VERSION,
        "state": "数据不足",
        "previous_state": "数据不足",
        "transition": "NONE",
        "score": 0.0,
        "sell_pressure_ratio": np.nan,
        "buy_pressure_ratio": np.nan,
        "sell_percentile": np.nan,
        "chase_percentile": np.nan,
        "absorption_score": 0.0,
        "drive_efficiency": 0.0,
        "chase_score": 0.0,
        "distribution_risk": 0.0,
        "dump_risk": 0.0,
        "near_high": False,
        "window_10": {},
        "window_20": {},
        "window_30": {},
        "sequence": [],
        "transition_score": 0.0,
        "transition_label": "无",
        "reasons": []
    }


def _cb_window_features(ret, volume, sell_proxy, buy_proxy, close_loc,
                        upper_wick, lookback, end_idx):
    """计算截至end_idx的历史窗口；窗口内部不读取end_idx之后的数据。"""
    start = max(0, end_idx - lookback + 1)
    rr = ret.iloc[start:end_idx + 1]
    vv = volume.iloc[start:end_idx + 1]
    ss = sell_proxy.iloc[start:end_idx + 1]
    bb = buy_proxy.iloc[start:end_idx + 1]
    cl = close_loc.iloc[start:end_idx + 1]
    uw = upper_wick.iloc[start:end_idx + 1]

    if len(rr) == 0:
        return {}

    return {
        "bars": int(len(rr)),
        "return_pct": _cb_safe(rr.sum()),
        "mean_return_pct": _cb_safe(rr.mean()),
        "up_ratio": _cb_safe((rr > 0).mean()),
        "down_ratio": _cb_safe((rr < 0).mean()),
        "volume_mean": _cb_safe(vv.mean()),
        "sell_mean": _cb_safe(ss.mean()),
        "buy_mean": _cb_safe(bb.mean()),
        "sell_pressure": _cb_safe(ss.sum() / max(vv.sum(), 1e-9)),
        "buy_pressure": _cb_safe(bb.sum() / max(vv.sum(), 1e-9)),
        "close_loc_mean": _cb_safe(cl.mean()),
        "upper_wick_mean": _cb_safe(uw.mean()),
    }


def _cb_classify_window(f, current_price, prior_high):
    """将一个历史窗口归类为吸收/推升/追涨/派发/砸盘/中性。"""
    if not f or f.get("bars", 0) < 5:
        return "NEUTRAL"

    ret = f["return_pct"]
    up = f["up_ratio"]
    down = f["down_ratio"]
    sell = f["sell_pressure"]
    buy = f["buy_pressure"]
    loc = f["close_loc_mean"]
    wick = f["upper_wick_mean"]

    near_high = False
    if np.isfinite(prior_high) and prior_high > 0:
        near_high = current_price >= prior_high * (1 - CAPITAL_BEHAVIOR_DISTRIBUTION_NEAR_HIGH_PCT / 100.0)

    # 吸收：卖压占比不低，但窗口累计跌幅很小/价格重心没有明显下移。
    if sell > 0.18 and ret > -1.0 and down < 0.60 and loc >= 0.50:
        return "ABSORBING"

    # 砸盘：卖压占比高 + 明显负收益 + 收盘位置弱。
    if sell > 0.28 and ret < -1.20 and loc < 0.45:
        return "DUMPING"

    # 高位派发：价格仍强/接近高位，但上影和弱收盘明显。
    if near_high and ret >= -0.50 and wick >= 0.25 and loc < 0.58:
        return "DISTRIBUTING"

    # 追涨：上涨、上涨占比高、收盘靠近高位。
    if ret > 1.20 and up >= 0.55 and loc >= 0.62 and near_high:
        return "CHASING"

    # 推升：上涨效率良好，但尚未明显进入高位追涨。
    if ret > 0.70 and up >= 0.52 and loc >= 0.56 and buy > sell:
        return "DRIVING"

    return "NEUTRAL"


def _cb_sequence_transition(states):
    """根据10/20/30窗口状态及当前事件识别行为阶段切换。"""
    if not states:
        return "NONE", "无", 0.0

    # 由近到远：10M、20M、30M
    s10, s20, s30 = states
    transition = "NONE"
    label = "无"
    score = 0.0

    # 砸盘 → 吸收 → 推升：最重要的反转链。
    if s30 == "DUMPING" and s20 in {"DUMPING", "ABSORBING"} and s10 in {"ABSORBING", "DRIVING"}:
        transition, label, score = "DUMP_TO_ABSORB_TO_DRIVE", "砸盘→吸收→推升", 15.0
    elif s20 == "DUMPING" and s10 in {"ABSORBING", "DRIVING"}:
        transition, label, score = "DUMP_TO_REVERSAL", "砸盘→反转", 12.0
    elif s20 == "ABSORBING" and s10 == "DRIVING":
        transition, label, score = "ABSORB_TO_DRIVE", "吸收→推升", 12.0
    elif s30 == "ABSORBING" and s20 == "ABSORBING" and s10 == "DRIVING":
        transition, label, score = "ABSORB_TO_DRIVE", "吸收→推升", 13.0
    elif s20 == "DRIVING" and s10 == "CHASING":
        transition, label, score = "DRIVE_TO_CHASE", "推升→追涨", 5.0
    elif s20 == "CHASING" and s10 == "DISTRIBUTING":
        transition, label, score = "CHASE_TO_DISTRIBUTION", "追涨→派发", -12.0
    elif s20 == "DRIVING" and s10 == "DISTRIBUTING":
        transition, label, score = "DRIVE_TO_DISTRIBUTION", "推升→派发", -10.0
    elif s20 == "DISTRIBUTING" and s10 in {"DISTRIBUTING", "DUMPING"}:
        transition, label, score = "DISTRIBUTION_TO_DUMP", "派发→砸盘", -15.0

    return transition, label, score


def calculate_capital_behavior_layer(df):
    """
    Capital Behavior V2：资金行为状态机。

    所有历史窗口均截止到当前K线；基准均使用shift(1)，避免当前K线污染。
    """
    if not CAPITAL_BEHAVIOR_ENABLED or df is None or df.empty:
        return _cb_empty()

    required = ["open", "high", "low", "close", "volume"]
    if any(c not in df.columns for c in required):
        return _cb_empty()

    d = df.copy()
    for c in required:
        d[c] = pd.to_numeric(d[c], errors="coerce")
    d = d.dropna(subset=required).copy()

    min_len = max(35, CAPITAL_BEHAVIOR_BASELINE + 5)
    if len(d) < min_len:
        return _cb_empty()

    o = d["open"].astype(float)
    h = d["high"].astype(float)
    l = d["low"].astype(float)
    c = d["close"].astype(float)
    v = d["volume"].astype(float).clip(lower=0)

    rng = (h - l).clip(lower=1e-9)
    ret = (c / o.replace(0, np.nan) - 1.0) * 100.0
    ret = ret.replace([np.inf, -np.inf], np.nan).fillna(0.0)
    close_loc = ((c - l) / rng).clip(0, 1)
    upper_wick_pct = ((h - np.maximum(o, c)) / o.replace(0, np.nan) * 100.0).replace([np.inf, -np.inf], np.nan).fillna(0)
    lower_wick_pct = ((np.minimum(o, c) - l) / o.replace(0, np.nan) * 100.0).replace([np.inf, -np.inf], np.nan).fillna(0)

    sell_proxy = v * ((o - c).clip(lower=0) / rng)
    buy_proxy = v * ((c - o).clip(lower=0) / rng)

    prior_v = v.shift(1).rolling(CAPITAL_BEHAVIOR_BASELINE, min_periods=10).mean()
    prior_sell = sell_proxy.shift(1).rolling(CAPITAL_BEHAVIOR_BASELINE, min_periods=10).mean()
    prior_buy = buy_proxy.shift(1).rolling(CAPITAL_BEHAVIOR_BASELINE, min_periods=10).mean()

    vol_ratio = _cb_safe(v.iloc[-1] / prior_v.iloc[-1], 0.0) if prior_v.iloc[-1] > 0 else 0.0
    sell_ratio = _cb_safe(sell_proxy.iloc[-1] / prior_sell.iloc[-1], 0.0) if prior_sell.iloc[-1] > 0 else 0.0
    buy_ratio = _cb_safe(buy_proxy.iloc[-1] / prior_buy.iloc[-1], 0.0) if prior_buy.iloc[-1] > 0 else 0.0

    prior_high = h.shift(1).rolling(CAPITAL_BEHAVIOR_LOOKBACK, min_periods=20).max().iloc[-1]
    current = _cb_safe(c.iloc[-1])
    near_high = bool(np.isfinite(prior_high) and prior_high > 0 and current >= prior_high * (1 - CAPITAL_BEHAVIOR_DISTRIBUTION_NEAR_HIGH_PCT / 100.0))

    current_ret = _cb_safe(ret.iloc[-1])
    current_loc = _cb_safe(close_loc.iloc[-1])
    current_upper = _cb_safe(upper_wick_pct.iloc[-1])
    current_lower = _cb_safe(lower_wick_pct.iloc[-1])

    # 当前事件分数。
    absorption = 0.0
    if sell_ratio >= CAPITAL_BEHAVIOR_HIGH_VOL and current_ret <= 0:
        if abs(current_ret) <= CAPITAL_BEHAVIOR_ABSORB_DROP_PCT:
            absorption += 8.0
        if current_loc >= 0.55:
            absorption += 4.0
        if current_lower >= current_upper:
            absorption += 2.0
    absorption = min(absorption, 14.0)

    drive_efficiency = 0.0
    if current_ret > 0 and vol_ratio > 1.0:
        drive_efficiency = current_ret / max(vol_ratio, 0.1)

    chase = 0.0
    if current_ret >= CAPITAL_BEHAVIOR_CHASE_RETURN_PCT and vol_ratio >= 1.20:
        if current_loc >= 0.70:
            chase += 3.0
        if near_high:
            chase += 3.0
        if buy_ratio >= 1.20:
            chase += 2.0
    chase = min(chase, 8.0)

    distribution = 0.0
    if near_high and vol_ratio >= CAPITAL_BEHAVIOR_HIGH_VOL:
        if current_upper >= CAPITAL_BEHAVIOR_DISTRIBUTION_UPPER_WICK_PCT:
            distribution += 4.0
        if current_loc < 0.60:
            distribution += 3.0
        if chase >= 3.0:
            distribution += 3.0
    distribution = min(distribution, 12.0)

    dump = 0.0
    if sell_ratio >= CAPITAL_BEHAVIOR_HIGH_VOL and current_ret <= -0.30:
        dump += 5.0
        if current_loc < 0.40:
            dump += 4.0
        if vol_ratio >= CAPITAL_BEHAVIOR_EXTREME_VOL:
            dump += 3.0
    dump = min(dump, 12.0)

    # 历史分位：只在当前K线之前的历史样本上计算。
    sell_hist = sell_proxy.shift(1).rolling(CAPITAL_BEHAVIOR_LOOKBACK, min_periods=20)
    chase_hist = ((ret > CAPITAL_BEHAVIOR_CHASE_RETURN_PCT).astype(float) *
                  (v / v.shift(1).rolling(CAPITAL_BEHAVIOR_BASELINE, min_periods=10).mean().replace(0, np.nan))).shift(1)
    sell_hist_values = sell_hist.apply(lambda x: pd.Series(x).rank(pct=True).iloc[-1] * 100 if len(x.dropna()) else np.nan, raw=False)
    sell_percentile = _cb_safe(sell_hist_values.iloc[-1], np.nan) if np.isfinite(sell_hist_values.iloc[-1]) else np.nan

    ch_series = chase_hist.dropna()
    if len(ch_series) >= 10:
        rank = ch_series.rank(pct=True)
        chase_percentile = _cb_safe(rank.iloc[-1] * 100, np.nan)
    else:
        chase_percentile = np.nan

    # 10/20/30历史窗口：用当前及之前的K线，绝不看未来。
    windows = {}
    states = []
    for w in CAPITAL_BEHAVIOR_WINDOWS:
        f = _cb_window_features(ret, v, sell_proxy, buy_proxy, close_loc, upper_wick_pct, w, len(d) - 1)
        state = _cb_classify_window(f, current, prior_high)
        f["state"] = state
        windows[w] = f
        states.append(state)

    transition, transition_label, transition_score = _cb_sequence_transition(states)

    # 当前状态优先由转折事件决定，否则取10根窗口状态。
    state_map = {
        "DUMP_TO_ABSORB_TO_DRIVE": "REVERSING",
        "DUMP_TO_REVERSAL": "REVERSING",
        "ABSORB_TO_DRIVE": "DRIVING",
        "DRIVE_TO_CHASE": "CHASING",
        "CHASE_TO_DISTRIBUTION": "DISTRIBUTING",
        "DRIVE_TO_DISTRIBUTION": "DISTRIBUTING",
        "DISTRIBUTION_TO_DUMP": "DUMPING",
    }
    state = state_map.get(transition)
    if state is None:
        state = states[0] if states else "NEUTRAL"

    # 当前事件对状态进行细化。
    if dump >= 9:
        state = "DUMPING"
    elif distribution >= 9:
        state = "DISTRIBUTING"
    elif transition in {"DUMP_TO_ABSORB_TO_DRIVE", "DUMP_TO_REVERSAL"}:
        state = "REVERSING"
    elif absorption >= 10 and current_ret <= 0:
        state = "ABSORBING"
    elif chase >= 6:
        state = "CHASING"
    elif drive_efficiency > 0.20 and current_ret > 0:
        state = "DRIVING"

    # V2评分：行为状态 + 转折优先，仍限制在±15。
    score = 0.0
    score += transition_score
    if state == "REVERSING":
        score += 6.0
    elif state == "DRIVING":
        score += 5.0
    elif state == "ABSORBING":
        score += 4.0
    elif state == "CHASING":
        score += 1.0
    elif state == "DISTRIBUTING":
        score -= 8.0
    elif state == "DUMPING":
        score -= 10.0

    if distribution >= 9:
        score -= 3.0
    if dump >= 9:
        score -= 4.0

    score = max(-15.0, min(15.0, score))

    reasons = []
    if transition_label != "无":
        reasons.append(transition_label)
    if absorption >= 10:
        reasons.append("当前卖压被吸收")
    if drive_efficiency > 0.20:
        reasons.append("放量对价格形成有效推升")
    if chase >= 6:
        reasons.append("追涨增强")
    if distribution >= 9:
        reasons.append("高位派发风险")
    if dump >= 9:
        reasons.append("真实砸盘风险")
    if not reasons:
        reasons.append("行为未形成明显阶段切换")

    return {
        "valid": True,
        "version": CAPITAL_BEHAVIOR_VERSION,
        "state": state,
        "previous_state": states[1] if len(states) > 1 else "NEUTRAL",
        "transition": transition,
        "transition_label": transition_label,
        "transition_score": transition_score,
        "score": round(score, 2),
        "sell_pressure_ratio": round(sell_ratio, 2),
        "buy_pressure_ratio": round(buy_ratio, 2),
        "sell_percentile": round(sell_percentile, 2) if np.isfinite(sell_percentile) else np.nan,
        "chase_percentile": round(chase_percentile, 2) if np.isfinite(chase_percentile) else np.nan,
        "absorption_score": round(absorption, 2),
        "drive_efficiency": round(drive_efficiency, 3),
        "chase_score": round(chase, 2),
        "distribution_risk": round(distribution, 2),
        "dump_risk": round(dump, 2),
        "near_high": near_high,
        "vol_ratio": round(vol_ratio, 2),
        "window_10": windows.get(10, {}),
        "window_20": windows.get(20, {}),
        "window_30": windows.get(30, {}),
        "sequence": states,
        "reasons": reasons,
    }



# ======================================================================
# Capital Behavior V2.2：记忆状态 ≠ 当前确认
# ======================================================================

def _cb_make_3m_bars(df):
    """将1M K线压缩为3M，只保留收齐3根1M的完整3M K线。"""
    if df is None or df.empty or "datetime" not in df.columns:
        return pd.DataFrame()
    d = df.copy()
    d["datetime"] = pd.to_datetime(d["datetime"], errors="coerce")
    d = d.dropna(subset=["datetime"]).sort_values("datetime")
    required = ["open", "high", "low", "close", "volume"]
    if any(c not in d.columns for c in required):
        return pd.DataFrame()
    for col in required:
        d[col] = pd.to_numeric(d[col], errors="coerce")
    d = d.dropna(subset=required)
    if len(d) < 3:
        return pd.DataFrame()
    d = d.set_index("datetime")
    agg = d[required].resample("3min", label="left", closed="left").agg({
        "open": "first", "high": "max", "low": "min", "close": "last", "volume": "sum"
    }).dropna(subset=["open", "high", "low", "close"])
    counts = d["close"].resample("3min", label="left", closed="left").count()
    counts = counts.reindex(agg.index).fillna(0)
    agg = agg[counts >= 3].copy()
    return agg.reset_index()


def _cb_3m_compatible(candidate, state3):
    """判断已完成3M状态是否支持1M候选状态。"""
    if not CAPITAL_BEHAVIOR_3M_CONFIRM or not state3:
        return True
    groups = {
        "ABSORBING": {"ABSORBING", "REVERSING", "DRIVING", "NEUTRAL"},
        "REVERSING": {"REVERSING", "ABSORBING", "DRIVING"},
        "DRIVING": {"DRIVING", "CHASING", "REVERSING", "ABSORBING"},
        "CHASING": {"CHASING", "DRIVING", "DISTRIBUTING"},
        "DISTRIBUTING": {"DISTRIBUTING", "CHASING", "DUMPING"},
        "DUMPING": {"DUMPING", "DISTRIBUTING", "NEUTRAL"},
        "NEUTRAL": {"NEUTRAL", "ABSORBING", "DRIVING", "CHASING", "DISTRIBUTING", "DUMPING", "REVERSING"},
    }
    return state3 in groups.get(candidate, {state3})


def _cb_hysteresis_allows_change(current, candidate, cb):
    """当前状态退出采用滞回阈值，避免临界值来回翻转。"""
    if not CAPITAL_BEHAVIOR_HYSTERESIS or current in {None, "数据不足"} or candidate == current:
        return True
    absorption = _cb_safe(cb.get("absorption_score"), 0.0)
    drive = _cb_safe(cb.get("drive_efficiency"), 0.0)
    chase = _cb_safe(cb.get("chase_score"), 0.0)
    distribution = _cb_safe(cb.get("distribution_risk"), 0.0)
    dump = _cb_safe(cb.get("dump_risk"), 0.0)
    if current == "ABSORBING" and candidate not in {"DUMPING", "DISTRIBUTING"}:
        return absorption < CB_EXIT_ABSORPTION
    if current == "DRIVING" and candidate not in {"DUMPING", "DISTRIBUTING"}:
        return drive < CB_EXIT_DRIVE_EFFICIENCY
    if current == "CHASING" and candidate not in {"DUMPING", "DISTRIBUTING"}:
        return chase < CB_EXIT_CHASE
    if current == "DISTRIBUTING" and candidate != "DUMPING":
        return distribution < CB_EXIT_DISTRIBUTION
    if current == "DUMPING":
        return dump < CB_EXIT_DUMP
    if current == "REVERSING":
        return candidate not in {"ABSORBING", "DRIVING"}
    return True


def _cb_current_confirmation_score(raw, stable_state, pending_state, pending_count, state3):
    """
    Capital Behavior V2.2：
    “记忆状态”与“当前确认”彻底分离。

    stable_state：
        只回答“最近一段时间资金行为处于什么阶段”。

    current confirmation：
        只回答“现在这一刻是否真的再次确认该阶段”。

    注意：
        历史稳定状态本身绝不直接给分。
        必须有当前行为 + 连续确认 +（必要时）3M确认，才能获得正向资金行为分。
    """
    raw_state = raw.get("state", "NEUTRAL")
    absorption = _cb_safe(raw.get("absorption_score"), 0.0)
    drive = _cb_safe(raw.get("drive_efficiency"), 0.0)
    chase = _cb_safe(raw.get("chase_score"), 0.0)
    distribution = _cb_safe(raw.get("distribution_risk"), 0.0)
    dump = _cb_safe(raw.get("dump_risk"), 0.0)

    # --------------------------------------------------------------
    # 当前1M是否具备“真实行为”
    # --------------------------------------------------------------
    # 注意：这里的 current_behavior 不等于 current_confirmed。
    # current_behavior 只代表“这一根1M确实出现了对应行为”；
    # current_confirmed 必须再经过连续3根 + 已完成3M同向确认。
    current_behavior = False
    strong_confirm = False

    if raw_state == "ABSORBING":
        current_behavior = absorption >= CB_ENTER_ABSORPTION
        strong_confirm = absorption >= 12.0

    elif raw_state == "DRIVING":
        current_behavior = drive >= CB_ENTER_DRIVE_EFFICIENCY
        strong_confirm = drive >= 0.30

    elif raw_state == "CHASING":
        current_behavior = chase >= CB_ENTER_CHASE
        strong_confirm = chase >= 7.0

    elif raw_state == "DISTRIBUTING":
        current_behavior = distribution >= CB_ENTER_DISTRIBUTION
        strong_confirm = distribution >= 11.0

    elif raw_state == "DUMPING":
        current_behavior = dump >= CB_ENTER_DUMP
        strong_confirm = dump >= 11.0

    elif raw_state == "REVERSING":
        current_behavior = (
            absorption >= CB_ENTER_ABSORPTION
            or drive >= CB_ENTER_DRIVE_EFFICIENCY
        )
        strong_confirm = absorption >= 12.0 or drive >= 0.30

    # --------------------------------------------------------------
    # 连续确认：只统计“当前行为真实成立”的1M K线。
    # 同一根K线不会重复累计（上层已经保证 pending_count 的去重）。
    # --------------------------------------------------------------
    if current_behavior and pending_state == raw_state:
        debounce = min(
            int(pending_count),
            CAPITAL_BEHAVIOR_CONFIRM_BARS
        )
    elif current_behavior and stable_state == raw_state:
        # 记忆状态与当前状态一致，也不能凭“记忆”直接变成确认。
        # 最多表示当前重新出现第1根有效行为。
        debounce = 1
    else:
        debounce = 0

    # --------------------------------------------------------------
    # 3M确认：只有“已完成3M同向”才能把1M连续行为升级为当前确认。
    # 若3M数据不足，则不阻塞3/3连续确认；一旦3M有效，则必须同向。
    # --------------------------------------------------------------
    three_m_ok = _cb_3m_compatible(raw_state, state3)
    three_m_available = bool(state3) and state3 != "数据不足"
    three_m_confirmed = bool(
        current_behavior
        and debounce >= CAPITAL_BEHAVIOR_CONFIRM_BARS
        and (
            not three_m_available
            or state3 == raw_state
        )
    )

    # 真正的“当前确认”：
    # 1) 当前这一根确实出现该行为
    # 2) 连续3根1M确认
    # 3) 若3M数据可用，则必须3M同向
    current_confirm = bool(three_m_confirmed)

    # 正向资金行为评分
    positive_score = 0.0
    if raw_state in {"ABSORBING", "DRIVING", "REVERSING"} and current_confirm:
        if debounce >= 3:
            positive_score = 4.0
        elif debounce == 2:
            positive_score = 2.0
        elif debounce == 1:
            positive_score = 1.0

        # 强吸收/强推升 + 已完成3M同向确认，才给额外权重。
        if strong_confirm and three_m_confirmed:
            positive_score = max(positive_score, 6.0)

        # 完整“砸盘→吸收→推升”路径且当前再次确认，才允许进入强资金评分。
        transition = raw.get("transition")
        if (
            transition == "DUMP_TO_ABSORB_TO_DRIVE"
            and current_confirm
            and three_m_confirmed
        ):
            positive_score = 10.0

        elif (
            transition == "DUMP_TO_REVERSAL"
            and current_confirm
            and three_m_confirmed
        ):
            positive_score = 8.0

    # 风险行为仍然实时生效：
    # 历史DISTRIBUTING/DUMPING不能因为“记忆状态”直接造成当前负分，
    # 但只要当前1M真实确认，就立即施加风险分。
    negative_score = 0.0
    if raw_state == "DISTRIBUTING" and current_confirm:
        negative_score = -8.0 if not strong_confirm else -12.0
    elif raw_state == "DUMPING" and current_confirm:
        negative_score = -10.0 if not strong_confirm else -15.0

    # 当前不是确认状态：资金行为评分严格归零。
    score = positive_score if positive_score > 0 else negative_score

    if not current_confirm:
        score = 0.0

    if not current_confirm:
        if debounce >= 2:
            strength = "PENDING"
        elif debounce == 1:
            strength = "WATCH"
        else:
            strength = "NONE"
    elif strong_confirm and three_m_confirmed:
        strength = "STRONG"
    else:
        strength = "CONFIRMED"

    return {
        "current_behavior": bool(current_behavior),
        "current_confirmed": bool(current_confirm),
        "current_confirm_count": int(debounce),
        "current_confirm_required": CAPITAL_BEHAVIOR_CONFIRM_BARS,
        "current_confirm_3m": bool(three_m_confirmed),
        "current_3m_state": state3 or "数据不足",
        "current_strength": strength,
        "current_score": round(score, 2),
        "historical_memory_state": stable_state or "NEUTRAL",
        "current_raw_state": raw_state,
        "three_m_compatible": bool(three_m_ok),
    }


def calculate_capital_behavior_stable(df, code="", name=""):
    """
    Capital Behavior V2.2：记忆状态 ≠ 当前确认。

    四个概念严格分离：

    1. historical_memory_state
       最近一段时间经过防抖/滞回后的“记忆状态”，只描述历史阶段。

    2. current_raw_state
       当前1M K线即时识别出的行为。

    3. current confirmation
       当前行为必须满足：
           当前行为阈值
           + 连续3根1M确认
           + 已完成3M同向确认（可用时）
       才能成为“当前确认”。

    4. score
       资金行为评分只由“当前确认”产生。
       历史记忆状态本身绝不直接加分。

    这样可以避免：
        “历史吸收” ≠ “现在仍在吸收”
        “历史推升” ≠ “现在仍在推升”
        “历史派发” ≠ “现在正在派发”
    """
    raw = calculate_capital_behavior_layer(df)
    if not raw.get("valid"):
        return raw

    key = str(code or name or "UNKNOWN")
    bar_key = None

    try:
        if "datetime" in df.columns:
            dd = pd.to_datetime(
                df["datetime"],
                errors="coerce"
            ).dropna()
            if not dd.empty:
                bar_key = str(dd.iloc[-1])
    except Exception:
        pass

    # --------------------------------------------------------------
    # 已完成3M K线
    # --------------------------------------------------------------
    df3 = _cb_make_3m_bars(df)

    cb3 = (
        calculate_capital_behavior_layer(df3)
        if len(df3) >= CAPITAL_BEHAVIOR_3M_MIN_BARS
        else _cb_empty()
    )

    state3 = (
        cb3.get("state", "")
        if cb3.get("valid")
        else ""
    )

    raw_state = raw.get("state", "NEUTRAL")

    rec = CB_STABILITY_STATE.get(key, {})

    stable_state = rec.get("stable_state")
    pending_state = rec.get("pending_state")
    pending_count = int(
        rec.get("pending_count", 0)
    )
    stable_bars = int(
        rec.get("stable_bars", 0)
    )
    last_bar = rec.get("last_bar")
    state_changed = False

    # --------------------------------------------------------------
    # ① 历史记忆状态：继续使用 V2.1 的防抖 + 滞回
    # --------------------------------------------------------------
    if not stable_state:
        stable_state = raw_state
        stable_bars = 1
        pending_state = None
        pending_count = 0
        state_changed = True

    elif raw_state == stable_state:
        stable_bars += (
            1 if last_bar != bar_key else 0
        )
        pending_state = None
        pending_count = 0

    elif last_bar == bar_key:
        # 同一根1M K线：不重复累计确认。
        pass

    else:
        hysteresis_ok = _cb_hysteresis_allows_change(
            stable_state,
            raw_state,
            raw
        )

        extreme_risk = (
            raw_state == "DUMPING"
            and _cb_safe(raw.get("dump_risk"), 0.0) >= 11.0
        )

        compatible_3m = _cb_3m_compatible(
            raw_state,
            state3
        )

        can_count = (
            hysteresis_ok
            and
            (
                compatible_3m
                or extreme_risk
            )
        )

        if can_count:
            if pending_state == raw_state:
                pending_count += 1
            else:
                pending_state = raw_state
                pending_count = 1
        else:
            pending_state = None
            pending_count = 0

        if (
            pending_state == raw_state
            and
            pending_count >= CAPITAL_BEHAVIOR_CONFIRM_BARS
        ):
            stable_state = raw_state
            stable_bars = 1
            pending_state = None
            pending_count = 0
            state_changed = True

    # --------------------------------------------------------------
    # ② V2.2：当前确认
    # --------------------------------------------------------------
    confirmation = _cb_current_confirmation_score(
        raw=raw,
        stable_state=stable_state,
        pending_state=pending_state,
        pending_count=pending_count,
        state3=state3
    )

    # --------------------------------------------------------------
    # ③ 持久化：只保存“历史记忆状态”，不保存当前评分
    # --------------------------------------------------------------
    CB_STABILITY_STATE[key] = {
        "stable_state": stable_state,
        "pending_state": pending_state,
        "pending_count": pending_count,
        "stable_bars": stable_bars,
        "last_bar": bar_key,
        "last_state_change": (
            datetime.datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )
            if state_changed
            else rec.get("last_state_change", "")
        ),
    }

    # --------------------------------------------------------------
    # ④ 返回结果
    # --------------------------------------------------------------
    result = dict(raw)

    result.update({
        "version": "V2.2",

        # 当前原始行为
        "raw_state": raw_state,
        "current_raw_state": raw_state,

        # 历史记忆状态
        "state": stable_state,
        "stable_state": stable_state,
        "historical_memory_state": stable_state,
        "stable_bars": int(stable_bars),

        # 当前确认
        "current_behavior": confirmation.get("current_behavior", False),
        "current_confirmed": confirmation["current_confirmed"],
        "current_confirm_count": confirmation["current_confirm_count"],
        "current_confirm_required": confirmation["current_confirm_required"],
        "current_confirm_3m": confirmation["current_confirm_3m"],
        "current_3m_state": confirmation["current_3m_state"],
        "current_strength": confirmation["current_strength"],
        "current_score": confirmation["current_score"],

        # 对外统一资金评分：
        # 只允许当前确认产生分数
        "score": confirmation["current_score"],

        # 历史状态切换信息
        "pending_state": pending_state or "无",
        "confirm_count": int(pending_count),
        "confirm_required": CAPITAL_BEHAVIOR_CONFIRM_BARS,
        "state_changed": bool(state_changed),

        # 3M
        "confirm_3m_state": state3 or "数据不足",
        "confirm_3m_valid": bool(cb3.get("valid")),
        "confirm_3m_score": _cb_safe(
            cb3.get("score"),
            0.0
        ),

        # 路径
        "path_text": (
            " → ".join(raw.get("sequence", []))
            if raw.get("sequence")
            else "无"
        ),

        "stability_text": (
            f"记忆:{stable_state}"
            if not pending_state
            else
            f"记忆:{stable_state} | "
            f"等待:{pending_state} "
            f"{pending_count}/{CAPITAL_BEHAVIOR_CONFIRM_BARS}"
        ),

        # V2.2明确字段
        "memory_vs_current": (
            "CONFIRMED"
            if confirmation["current_confirmed"]
            else "MEMORY_ONLY"
        ),
    })

    return result


# ======================================================================
# 持仓风控
# ======================================================================

def check_position_risk(
    status,
    cost,
    current,
    highest
):
    """T+1持仓风险检查：固定止损 -3%，盈利超过5%后高点回撤2%移动止盈。"""
    if str(status).strip() not in [
        "是", "持有", "True", "1"
    ]:
        return None

    current = safe_float(current)
    cost = safe_float(cost, current)
    highest = safe_float(highest, current)

    if (
        not np.isfinite(current)
        or current <= 0
        or not np.isfinite(cost)
        or cost <= 0
    ):
        return None

    if (current - cost) / cost <= -0.03:
        return "🔴触发固定止损线(-3%)，建议清仓"

    if (
        np.isfinite(highest)
        and highest > cost * 1.05
        and (highest - current) / highest >= 0.02
    ):
        return "🟡触发移动止盈线(高点回撤2%)，锁定利润"

    return None


# ======================================================================
# Excel
# ======================================================================

class ExcelManager:
    """StockPulse Excel股票池读写管理。

    V7.2修复：补回主扫描依赖的ExcelManager，避免启动后在
    perform_scan_round()处出现 NameError。
    """

    @classmethod
    def load_stock_pool(cls, path=EXCEL_PATH):
        if not os.path.exists(path):
            logger.warning(f"Excel不存在，自动创建默认股票池: {path}")
            d = pd.DataFrame([
                {
                    "代码": "600519",
                    "名称": "贵州茅台",
                    "持仓状态": "否",
                    "买入价": 0,
                    "盘口卖买比": 1.2,
                }
            ])
            d.to_excel(path, index=False)
            return d

        d = pd.read_excel(path)
        if d is None or d.empty:
            logger.warning(f"Excel股票池为空: {path}")
            return pd.DataFrame(columns=["代码", "名称", "持仓状态", "买入价", "盘口卖买比"])

        if "代码" not in d.columns:
            raise ValueError(f"Excel缺少【代码】列: {path}")

        d["代码"] = (
            d["代码"]
            .astype(str)
            .str.replace(".0", "", regex=False)
            .str.strip()
            .str.zfill(6)
        )

        if "名称" not in d.columns:
            d["名称"] = d["代码"]
        if "持仓状态" not in d.columns:
            d["持仓状态"] = "否"
        if "买入价" not in d.columns:
            d["买入价"] = 0
        if "盘口卖买比" not in d.columns:
            d["盘口卖买比"] = 1.2

        logger.info(f"Excel股票池加载成功: {len(d)} 只股票")
        return d

    @classmethod
    def write_signal(cls, data, path=EXCEL_PATH):
        try:
            if not os.path.exists(path):
                logger.warning(f"Excel不存在，跳过信号回写: {path}")
                return

            d = pd.read_excel(path)
            if "代码" not in d.columns:
                logger.error("Excel缺少【代码】列，无法回写信号")
                return

            code = str(data.get("代码", "")).replace(".0", "").strip().zfill(6)
            codes = d["代码"].astype(str).str.replace(".0", "", regex=False).str.strip().str.zfill(6)
            mask = codes == code
            if not mask.any():
                logger.warning(f"Excel中未找到股票 {code}")
                return

            for col in ["交易信号", "信号时间", "综合评分"]:
                if col not in d.columns:
                    d[col] = ""

            d.loc[mask, "交易信号"] = data.get("交易信号", "")
            d.loc[mask, "信号时间"] = data.get("信号时间", "")
            d.loc[mask, "综合评分"] = data.get("综合评分", 0)

            d.to_excel(path, index=False)
            logger.info(f"成功将信号回写至 {path}: {code}")

        except Exception as e:
            logger.exception(f"Excel写入失败: {e}")


# ======================================================================
# V7.6 五星时间确认状态机
# ======================================================================
def update_v76_five_star_confirmation(code, raw_candidate, now=None):
    """
    五星确认不再按“扫描次数”计数，而按“跨时间窗口的有效确认”计数。

    状态：
      IDLE -> PENDING_CONFIRM -> CONFIRMED_BUY
      任一关键条件失效 -> INVALIDATED/IDLE

    注意：这只是信号确认，不执行交易。A股股票本身仍按T+1处理。
    """
    now = now or datetime.datetime.now()
    st = V76_FIVE_STAR_STATE.setdefault(code, {
        "state": "IDLE",
        "count": 0,
        "first_time": None,
        "last_confirm_time": None,
        "last_seen_time": None,
        "invalidations": 0,
    })

    st["last_seen_time"] = now

    # 非五星：如果处于等待期，立即失效；避免信号断掉后继续累计。
    if not raw_candidate:
        if st.get("state") in {"PENDING_CONFIRM", "CONFIRMED_BUY"}:
            st["invalidations"] = int(st.get("invalidations", 0)) + 1
        st.update({
            "state": "IDLE",
            "count": 0,
            "first_time": None,
            "last_confirm_time": None,
        })
        return {**st, "ready": False, "reason": "当前五星条件失效，确认链重置"}

    first_time = st.get("first_time")
    last_time = st.get("last_confirm_time")

    if first_time is None:
        st.update({
            "state": "PENDING_CONFIRM",
            "count": 1,
            "first_time": now,
            "last_confirm_time": now,
        })
        return {**st, "ready": False, "reason": "首次五星触发，进入时间确认等待"}

    # 等待超时，重新以当前信号作为第一确认点。
    if (now - first_time).total_seconds() > V76_FIVE_STAR_PENDING_EXPIRE_MIN * 60:
        st.update({
            "state": "PENDING_CONFIRM",
            "count": 1,
            "first_time": now,
            "last_confirm_time": now,
        })
        return {**st, "ready": False, "reason": "五星确认等待超时，重新开始确认"}

    # 同一确认窗口内的30秒轮询不重复计数。
    if last_time is not None:
        elapsed_min = (now - last_time).total_seconds() / 60.0
        if elapsed_min < V76_FIVE_STAR_CONFIRM_INTERVAL_MIN:
            ready = st.get("state") == "CONFIRMED_BUY"
            return {**st, "ready": ready, "reason": f"确认间隔不足，等待跨越{V76_FIVE_STAR_CONFIRM_INTERVAL_MIN}分钟"}

    st["count"] = min(V71_CONFIRM_SCANS, int(st.get("count", 0)) + 1)
    st["last_confirm_time"] = now

    if st["count"] >= V71_CONFIRM_SCANS:
        st["state"] = "CONFIRMED_BUY"
        return {**st, "ready": True, "reason": "五星已完成跨时间连续确认"}

    st["state"] = "PENDING_CONFIRM"
    return {**st, "ready": False, "reason": "五星仍在等待下一次时间确认"}


# ======================================================================
# V7.7 独立最终交易权限
# ======================================================================
def calculate_v77_trade_permission(stars, raw_five_star, stable_five_star, v71_gate, sector_gate, trend_info, entry_gate, confirmation):
    decision=v71_gate.get("decision_state","BLOCK")
    if decision=="BLOCK": return {"state":"BLOCK","reason":"最终交易闸门阻断"}
    if raw_five_star and not stable_five_star: return {"state":"WAIT","reason":confirmation.get("reason","等待五星时间确认")}
    if stars==5 and stable_five_star:
        if v71_gate.get("special"): return {"state":"HUMAN_CONFIRM","reason":"月K长期风险高，仅允许人工确认"}
        return {"state":"HUMAN_CONFIRM","reason":"五星稳定确认完成，等待人工确认"}
    if stars in (3,4): return {"state":"OBSERVE","reason":"观察级信号，不自动开仓"}
    return {"state":"BLOCK","reason":"未达到交易星级"}


# ======================================================================
# 单轮扫描
# ======================================================================

def perform_scan_round(
    tdx_pool,
    market_analyzer
):

    global ALERTED_STOCKS

    stock_pool = (
        ExcelManager
        .load_stock_pool(
            EXCEL_PATH
        )
    )

    # ----------------------------------------------------------
    # 整轮只计算一次月K
    # ----------------------------------------------------------

    month_regime = (
        market_analyzer
        .get_market_regime()
    )

    if month_regime.get(
        "valid",
        False
    ):

        logger.info(
            "Market Regime: "
            f"{month_regime['state']} "
            f"| DIF="
            f"{month_regime['dif']:.4f} "
            f"DEA="
            f"{month_regime['dea']:.4f} "
            f"HIST="
            f"{month_regime['hist']:.4f} "
            f"| bars="
            f"{month_regime['bars']} "
            f"| "
            +
            (
                "禁止五星"
                if
                month_regime[
                    "risk_block_5star"
                ]
                else
                "允许五星"
            )
        )

    else:

        logger.warning(
            "Market Regime: UNKNOWN "
            "| 月K数据不足 "
            "| 五星禁止"
        )

    # ----------------------------------------------------------
    # V7.5.2：整轮只计算一次市场快照
    # ----------------------------------------------------------
    market_info = market_analyzer.analyze(stock_pool=stock_pool)
    if market_info.get("valid"):
        idx = market_info.get("indices", {})
        sh = idx.get("SH", {}); sz = idx.get("SZ", {}); cy = idx.get("CY", {})
        breadth = market_info.get("breadth", {})
        logger.info(
            "V7.6.0 MARKET SNAPSHOT | "
            f"上证={safe_float(sh.get('change'), np.nan):+.2f}% | "
            f"深证={safe_float(sz.get('change'), np.nan):+.2f}% | "
            f"创业板={safe_float(cy.get('change'), np.nan):+.2f}% | "
            f"全市场↑={breadth.get('up', 0)} ↓={breadth.get('down', 0)} →/停={breadth.get('flat', 0)} / {breadth.get('members', 0)} | "
            f"↑占比={safe_float(breadth.get('up_ratio'), 0)*100:.1f}% | "
            f"BreadthState={market_info.get('breadth_state', 'UNKNOWN_BREADTH')} | "
            f"广度源={breadth.get('source', 'NONE')} | "
            f"MarketState={market_info.get('state5', 'E')} | Score={market_info.get('score', 0):+.1f} | IndustrySource={market_info.get('sector_source', 'NONE')}"
        )

    # ----------------------------------------------------------
    # 股票循环
    # ----------------------------------------------------------

    for _, row in stock_pool.iterrows():

        try:

            code = (
                str(
                    row.get(
                        "代码",
                        ""
                    )
                )
                .replace(
                    ".0",
                    ""
                )
                .strip()
                .zfill(6)
            )

            if not code:

                continue

            name = str(
                row.get(
                    "名称",
                    code
                )
            )

            market = get_market_code(
                code
            )

            # --------------------------------------------------
            # 多周期
            # --------------------------------------------------

            df_1m = (
                tdx_pool
                .get_bars(
                    CATEGORY_1M,
                    market,
                    code,
                    0,
                    240
                )
            )

            df_15m = (
                tdx_pool
                .get_bars(
                    CATEGORY_15M,
                    market,
                    code,
                    0,
                    100
                )
            )

            df_30m = (
                tdx_pool
                .get_bars(
                    CATEGORY_30M,
                    market,
                    code,
                    0,
                    100
                )
            )

            # --------------------------------------------------
            # 指标
            # --------------------------------------------------

            macd_1m = calculate_macd(
                df_1m
            )

            macd_15m = calculate_macd(
                df_15m
            )

            macd_30m = calculate_macd(
                df_30m
            )

            vwap_info = calculate_vwap(
                df_1m
            )

            sell_buy_ratio = safe_float(
                row.get(
                    "盘口卖买比",
                    1.2
                ),
                1.2
            )

            dar_info = (
                calculate_dar_and_volume(
                    df_1m,
                    sell_buy_ratio
                )
            )

            # --------------------------------------------------
            # 资金行为 V2.2：记忆状态 ≠ 当前确认
            # --------------------------------------------------
            capital_behavior = calculate_capital_behavior_stable(
                df_1m,
                code=code,
                name=name
            )

            # --------------------------------------------------
            # V7.6.0：市场快照只在本轮开始时计算一次。
            # 个股分析严禁再次调用 market_analyzer.analyze()，
            # 所有股票共享同一个 immutable-by-convention 快照。
            # --------------------------------------------------
            # market_info 已在本轮股票循环之前生成。

            # --------------------------------------------------
            # V7.5：板块/题材第二层闸门
            # 必须在 raw_five_star_candidate 之前计算。
            # 之前版本遗漏了这一变量的赋值，导致所有股票
            # 在进入 V7.1 总闸门前直接 NameError。
            # --------------------------------------------------
            sector_gate = calculate_sector_gate(
                code=code,
                row=row,
                market_info=market_info,
                stock_df=df_1m
            )

            # --------------------------------------------------
            # 当前价格
            # --------------------------------------------------

            if (
                df_1m is not None
                and
                not df_1m.empty
                and
                "close" in df_1m.columns
            ):

                current_price = safe_float(
                    df_1m[
                        "close"
                    ].iloc[-1]
                )

            else:

                current_price = np.nan

            # --------------------------------------------------
            # 个股趋势分层
            # --------------------------------------------------
            trend_info = calculate_stock_trend_layer(
                tdx_pool, market, code
            )

            # --------------------------------------------------
            # ENTRY_GATE V6.3
            # --------------------------------------------------
            # V6.2负责“位置/突破”，趋势层负责“大周期方向”。
            entry_gate = calculate_entry_gate(
                df_1m,
                market_info,
                dar_info,
                current_price
            )

            # V7.1：保存“纯位置/突破”结果。后面趋势层会修改 entry_gate，
            # 但 raw 五星候选必须能独立于 V7 市场总闸门计算。
            entry_gate_base = dict(entry_gate)
            entry_gate_base["reasons"] = list(entry_gate.get("reasons", []))

            # --------------------------------------------------
            # ENTRY_GATE V6.3：趋势层并入总闸门
            # --------------------------------------------------
            # V6.2 = 位置/突破闸门
            # V6.3 = V6.2 + 个股大周期趋势闸门
            # BLOCK：禁止★★★★★主动开仓
            # WAIT ：允许继续观察，但不形成五星主动开仓
            # PASS ：进入原有五星硬条件判断
            if TREND_LAYER_ENABLED:
                trend_gate = trend_info.get("gate", "WAIT")
                if trend_gate == "BLOCK":
                    entry_gate["passed"] = False
                    entry_gate["state"] = "BLOCK"
                    entry_gate.setdefault("reasons", []).append(
                        f"个股趋势{trend_info.get('grade', 'UNKNOWN')}级："
                        f"{trend_info.get('reason', '趋势偏弱')}"
                    )
                elif trend_gate == "WAIT":
                    entry_gate["passed"] = False
                    entry_gate["state"] = "WAIT"
                    entry_gate.setdefault("reasons", []).append(
                        f"个股趋势C级观察：{trend_info.get('reason', '趋势未明确')}"
                    )

            # --------------------------------------------------
            # 持仓风险
            # --------------------------------------------------

            risk_message = (
                check_position_risk(
                    row.get(
                        "持仓状态"
                    ),
                    row.get(
                        "买入价"
                    ),
                    current_price,
                    row.get(
                        "最高价",
                        current_price
                    )
                )
            )

            # ==================================================
            # 评分
            # ==================================================

            score_1m = (
                macd_1m["score"]
            )

            score_15m = (
                macd_15m["score"]
            )

            score_30m = (
                macd_30m["score"]
            )

            score_vwap = (
                vwap_info["score"]
            )

            score_dar = (
                dar_info["dar_score"]
            )

            # 资金行为 V2.2：只有“当前确认”才允许进入评分；历史记忆状态不直接加分
            score_capital = safe_float(
                capital_behavior.get("score"),
                0.0
            )

            score_market = (
                market_info["score"]
            )

            # --------------------------------------------------
            # 月K默认不直接参与评分
            # --------------------------------------------------

            score_month = 0

            if MONTH_SCORE_ENABLED:

                if (
                    month_regime.get(
                        "state"
                    )
                    ==
                    "BULLISH"
                ):

                    score_month = 10

                elif (
                    month_regime.get(
                        "state"
                    )
                    ==
                    "BEARISH"
                ):

                    score_month = -10

            # --------------------------------------------------
            # V4.5 个股趋势层评分
            # --------------------------------------------------
            score_trend = safe_float(
                trend_info.get("score"), 0
            )

            # --------------------------------------------------
            # 最终评分
            #
            # 中性基准50
            #
            # 原程序的问题：
            #
            # 多个模块直接相加后可能超过100，
            # 再直接截断100。
            #
            # 结果：
            #
            # 90分、120分、150分
            # 全部变成100。
            #
            # 当前因子权重已经重新压缩：
            #
            # 1M      ±12
            # 15M      ±6
            # 30M      ±6
            # VWAP     ±8
            # DAR     -10~15
            # 市场    -15~15
            #
            # 使100分真正更难出现。
            # --------------------------------------------------

            raw_score = (
                50
                +
                score_1m
                +
                score_15m
                +
                score_30m
                +
                score_vwap
                +
                score_dar
                +
                score_capital
                +
                score_market
                +
                score_month
                +
                score_trend
            )

            final_score = int(
                max(
                    0,
                    min(
                        100,
                        round(
                            raw_score
                        )
                    )
                )
            )

            # ==================================================
            # StockPulse V7.3 PRO：稳定资金行为 → 原始五星 → 连续确认 → 市场总闸门
            # ==================================================

            # ① 原始五星候选：故意不读取 V7 市场总闸门结果。
            #    月K BEARISH 也不在这里直接抹掉，交给 V7.1 决策。
            raw_five_star_candidate = bool(
                final_score >= 85
                and macd_1m["valid"]
                and macd_15m["valid"]
                and macd_30m["valid"]
                and vwap_info["valid"]
                and market_info["valid"]
                and not dar_info["overheat"]
                and capital_behavior.get("distribution_risk", 0.0) < 10.0
                and capital_behavior.get("dump_risk", 0.0) < 9.0
                and capital_behavior.get("raw_state", "NEUTRAL") not in {"DISTRIBUTING", "DUMPING"}
                and (
                    capital_behavior.get("current_confirmed", False)
                    or capital_behavior.get("raw_state", "NEUTRAL") == "NEUTRAL"
                )
                and capital_behavior.get("transition") not in {"CHASE_TO_DISTRIBUTION", "DRIVE_TO_DISTRIBUTION", "DISTRIBUTION_TO_DUMP"}
                and market_info.get("pressure_distance", 999) >= 0.50
                and entry_gate_base.get("passed", False)
                and trend_info.get("grade") in {"A", "B"}
                and sector_gate.get("state") == "STRONG"
            )

            # ② V7.6：时间确认。30秒轮询不重复计数，必须跨越确认间隔。
            confirmation = update_v76_five_star_confirmation(
                code,
                raw_five_star_candidate,
                datetime.datetime.now()
            )
            V71_STABILITY_COUNT[code] = int(confirmation.get("count", 0))
            stable_five_star = bool(confirmation.get("ready", False))

            # ③ 市场总闸门：最后才决定今天是否放行。
            v71_gate = calculate_v71_market_gate(
                market_info=market_info,
                month_regime=month_regime,
                trend_info=trend_info,
                entry_gate=entry_gate,
                dar_info=dar_info,
                raw_five_star=raw_five_star_candidate,
                stable_five_star=stable_five_star,
                final_score=final_score,
                macd_1m=macd_1m,
                macd_15m=macd_15m,
                macd_30m=macd_30m,
                vwap_info=vwap_info,
                sector_gate=sector_gate
            )

            # ④ 最终五星：原始条件 + 连续确认 + V7.1放行。
            final_five_star = bool(
                raw_five_star_candidate
                and stable_five_star
                and v71_gate.get("passed", False)
            )

            stars = 0

            if final_five_star:
                stars = 5
            elif (
                final_score >= 75
                and macd_1m["valid"]
                and vwap_info["valid"]
            ):
                stars = 4
            elif final_score >= 65:
                stars = 3

            trade_permission = calculate_v77_trade_permission(
                stars, raw_five_star_candidate, stable_five_star, v71_gate,
                sector_gate, trend_info, entry_gate, confirmation
            )

            # 交易建议
            if stars == 5:
                if v71_gate.get("special", False):
                    advice = "🟠触发★★★★★逆势反转候选 | 仅人工确认，不自动开仓"
                else:
                    advice = "🟢触发★★★★★强买入候选 | 连续确认通过 | 人工确认"
            elif stars == 4:
                if trend_info.get("gate") == "BLOCK":
                    advice = "🟡触发A级观察 (★★★★) | 趋势闸门阻断主动开仓"
                elif not entry_gate.get("passed", False):
                    advice = "🟡触发A级观察 (★★★★) | ENTRY_GATE要求观察"
                elif not stable_five_star and raw_five_star_candidate:
                    advice = f"🟡触发A级观察 (★★★★) | 五星时间确认 {V71_STABILITY_COUNT.get(code, 0)}/{V71_CONFIRM_SCANS} | {confirmation.get('reason', '')}"
                elif not v71_gate.get("passed", False):
                    advice = "🟡触发A级观察 (★★★★) | V7.1市场总闸门要求等待"
                else:
                    advice = "🟢触发A级推荐 (★★★★)"
            elif stars == 3:
                advice = "🟡触发B级推荐 (★★★)"
            else:
                advice = "⚪观望 (未达到星级条件)"

            timestamp = (
                datetime.datetime.now()
                .strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
            )

            # ==================================================
            # 输出
            # ==================================================

            print(
                "\n"
                +
                "=" * 78
            )

            print(
                f"【{name} ({code})】"
                f"诊股报告 - {timestamp}"
            )

            print(
                "-" * 78
            )

            # --------------------------------------------------
            # Market Regime
            # --------------------------------------------------

            print(
                "【Market Regime】"
            )

            if month_regime.get(
                "valid",
                False
            ):

                print(
                    f"上证月K："
                    f"{month_regime['text']}"
                )

                print(
                    f"月K状态："
                    f"{month_regime['state']}"
                )

                print(
                    f"月K数据源："
                    f"{month_regime['source']} "
                    f"| K线数量="
                    f"{month_regime['bars']}"
                )

                print(
                    "月K MACD："
                    f"DIF="
                    f"{month_regime['dif']:.4f} "
                    f"DEA="
                    f"{month_regime['dea']:.4f} "
                    f"HIST="
                    f"{month_regime['hist']:.4f} "
                    f"前HIST="
                    f"{month_regime['prev_hist']:.4f}"
                )

                print(
                    f"月K结构："
                    f"{month_regime['structure']}"
                )

                print(
                    "月K风险过滤："
                    +
                    (
                        "⚠五星禁止触发"
                        if
                        month_regime[
                            "risk_block_5star"
                        ]
                        else
                        "✓允许五星"
                    )
                )

            else:

                print(
                    "上证月K：数据不足"
                )

                print(
                    "月K状态：UNKNOWN"
                )

                print(
                    "月K风险过滤："
                    "⚠五星禁止触发"
                )

            # --------------------------------------------------
            # 个股趋势分层 V4.5
            # --------------------------------------------------
            print(
                "【个股趋势分层】 "
                f"{trend_info.get('grade', 'UNKNOWN')}级 "
                f"| 趋势分={trend_info.get('score', 0):.0f} "
                f"| 趋势闸门={trend_info.get('gate', 'WAIT')}"
            )

            if trend_info.get("valid", False):
                print(
                    "   月K MA20："
                    f"{trend_info['monthly']['states']['MA20']} "
                    f"| 价位={trend_info['monthly']['price_states']['MA20']} "
                    f"| MA20={trend_info['monthly']['values']['MA20']:.2f}"
                )
                print(
                    "   周K MA20/60："
                    f"{trend_info['weekly']['states']['MA20']}/"
                    f"{trend_info['weekly']['states']['MA60']} "
                    f"| 价位="
                    f"{trend_info['weekly']['price_states']['MA20']}/"
                    f"{trend_info['weekly']['price_states']['MA60']}"
                )
                print(
                    "   日K MA20/60："
                    f"{trend_info['daily']['states']['MA20']}/"
                    f"{trend_info['daily']['states']['MA60']} "
                    f"| 价位="
                    f"{trend_info['daily']['price_states']['MA20']}/"
                    f"{trend_info['daily']['price_states']['MA60']}"
                )
                print(
                    "   趋势结论："
                    f"{trend_info.get('reason', '')}"
                )
            else:
                print(
                    "   趋势数据：不足，趋势层不允许★★★★★"
                )

            # --------------------------------------------------
            # 市场
            # --------------------------------------------------

            if market_info["valid"]:

                print(
                    f"1. 市场环境: "
                    f"{market_info['state']} "
                    f"(压力距离: "
                    f"{market_info['pressure_distance']:.2f}%)"
                )

            else:

                print(
                    "1. 市场环境: "
                    "指数数据不足"
                )

            # --------------------------------------------------
            # StockPulse V7.1 市场总闸门
            # --------------------------------------------------
            print(
                "7. StockPulse V7.8.1 市场环境/交易闸门: "
                f"Market={v71_gate.get('state', 'E')} "
                f"| 市场状态={market_info.get('state5', 'E')} "
                f"| 市场评分={market_info.get('score', 0):.1f} "
                f"| 上证当日涨跌={market_info.get('idx_day_change', np.nan):.2f}% "
                f"| 压力距离={market_info.get('pressure_distance', np.nan):.2f}% "
                f"| 月K长期风险={v71_gate.get('long_term_risk', 'UNKNOWN')} "
                f"| 个股趋势={trend_info.get('grade', 'UNKNOWN')} "
                f"| Decision={v71_gate.get('decision_state', 'BLOCK')}"
            )
            print(
                "   V7.1："
                f"raw五星={'是' if raw_five_star_candidate else '否'} "
                f"| 连续确认={V71_STABILITY_COUNT.get(code, 0)}/{V71_CONFIRM_SCANS} "
                f"| 稳定五星={'是' if stable_five_star else '否'} "
                f"| 最终五星={'是' if final_five_star else '否'} "
                f"| 逆势人工={'是' if v71_gate.get('special') else '否'} "
                f"| 时间确认状态={confirmation.get('state', 'IDLE')} "
                f"| {confirmation.get('reason', '')}"
            )
            if v71_gate.get("reasons"):
                for reason in v71_gate["reasons"]:
                    print("   ⚠ V7.1：" + reason)
            else:
                print("   ✓ V7.8.1：市场环境与个股闸门通过")

            print(f"   V7.8.1最终交易权限：{trade_permission.get('state','BLOCK')} | {trade_permission.get('reason','')}")

            breadth = market_info.get("breadth", {})
            if breadth.get("valid"):
                print(
                    "   全市场广度："
                    f"↑={breadth.get('up', 0)} ↓={breadth.get('down', 0)} "
                    f"| →/停={breadth.get('flat', 0)} / {breadth.get('members', 0)} "
                    f"| 上涨占比={safe_float(breadth.get('up_ratio'), 0)*100:.1f}% "
                    f"| BreadthState={market_info.get('breadth_state', 'UNKNOWN_BREADTH')}"
                )

            print(
                "   行业第二闸门 V7.8.1："
                f"{sector_gate.get('state', 'UNKNOWN')} "
                f"| {sector_gate.get('sector', '未分类')} "
                f"| 平均涨跌={safe_float(sector_gate.get('mean_change'), np.nan):.2f}% "
                f"| 上涨占比={safe_float(sector_gate.get('up_ratio'), np.nan)*100:.1f}% "
                f"| 行为={sector_gate.get('behavior_state', 'UNKNOWN')} "
                f"| 30M={safe_float(sector_gate.get('behavior_trend_30m'), np.nan):.2f}% "
                f"| 60M={safe_float(sector_gate.get('behavior_trend_60m'), np.nan):.2f}% "
                f"| 加速={safe_float(sector_gate.get('behavior_accel'), np.nan):.3f}%/10M "
                f"| 退潮={safe_float(sector_gate.get('behavior_retreat'), np.nan):.2f}% "
                f"| 分化STD={safe_float(sector_gate.get('behavior_dispersion'), np.nan):.2f}% "
                f"| 分化价差={safe_float(sector_gate.get('behavior_spread'), np.nan):.2f}% "
                f"| 成员={sector_gate.get('members', 0)} "
                f"| 样本={sector_gate.get('behavior_samples', 0)} "
                f"| 来源={sector_gate.get('source', 'NONE')}"
            )

            # --------------------------------------------------
            # ENTRY_GATE
            # --------------------------------------------------
            print(
                "6. ENTRY_GATE V6.3: "
                f"{entry_gate['state']} "
                f"| 当前局部压力="
                f"{entry_gate['pressure']:.2f} "
                f"| 压力距离="
                f"{entry_gate['distance_to_pressure']:.2f}%"
            )

            print(
                "   当前突破="
                f"{'是' if entry_gate['breakout'] else '否'} "
                f"| 当前有效站稳="
                f"{'是' if entry_gate['breakout_confirmed'] else '否'} "
                f"| 今日曾突破="
                f"{'是' if entry_gate['today_breakout'] else '否'} "
                f"| 今日突破失败="
                f"{'是' if entry_gate['today_breakout_failed'] else '否'}"
            )

            if entry_gate.get("today_breakout", False):
                event_pressure = safe_float(
                    entry_gate.get("today_breakout_pressure"),
                    np.nan
                )
                event_price = safe_float(
                    entry_gate.get("today_breakout_price"),
                    np.nan
                )

                if np.isfinite(event_pressure) and np.isfinite(event_price):
                    print(
                        "   今日突破事件："
                        f"压力={event_pressure:.2f} "
                        f"→最高={event_price:.2f} "
                        f"| 时间={entry_gate.get('today_breakout_time', '')}"
                    )

            volume_state_text = {
                "NOT_TRIGGERED": "未触发",
                "PASS": "通过",
                "BLOCK": "不足"
            }.get(
                entry_gate.get("volume_state", "NOT_TRIGGERED"),
                "未触发"
            )

            print(
                "   突破量能="
                f"{volume_state_text}"
            )

            if entry_gate["reasons"]:
                for gate_reason in entry_gate["reasons"]:
                    print(
                        "   ⚠ ENTRY_GATE："
                        + gate_reason
                    )
            else:
                print(
                    "   ✓ ENTRY_GATE："
                    "位置与突破结构通过"
                )

            # --------------------------------------------------
            # VWAP
            # --------------------------------------------------

            if vwap_info["valid"]:

                print(
                    f"2. VWAP均价: "
                    f"{vwap_info['vwap']:.4f} "
                    f"(偏离: "
                    f"{vwap_info['distance']:.2f}%)"
                )

            else:

                print(
                    "2. VWAP均价: "
                    f"异常/熔断 "
                    f"({vwap_info['reason']})"
                )

            # --------------------------------------------------
            # MACD
            # --------------------------------------------------

            print(
                "3. 多周期MACD: "
                f"1M[{macd_1m['state']}] "
                f"15M[{macd_15m['state']}] "
                f"30M[{macd_30m['state']}]"
            )

            # --------------------------------------------------
            # DAR
            # --------------------------------------------------

            vol_ratio = dar_info[
                "vol_ratio"
            ]

            if np.isfinite(
                vol_ratio
            ):

                vol_text = (
                    f"{vol_ratio:.2f}倍"
                )

            else:

                vol_text = "数据不足"

            print(
                "4. DAR微观形态: "
                f"{dar_info['dar_state']} "
                f"| 量能比: "
                f"{vol_text} "
                f"| 量价过热: "
                f"{'是' if dar_info['overheat'] else '否'}"
            )

            # --------------------------------------------------
            # --------------------------------------------------
            # Capital Behavior V2：资金行为状态机
            # --------------------------------------------------
            cb = capital_behavior
            if cb.get("valid"):
                def _cbfmt(x, suffix=""):
                    try:
                        return f"{float(x):.2f}{suffix}" if np.isfinite(float(x)) else "N/A"
                    except Exception:
                        return "N/A"

                print(
                    "4A. 资金行为 V2.2：记忆状态 ≠ 当前确认"
                )
                print(
                    "   历史记忆状态="
                    f"{cb.get('historical_memory_state', cb.get('stable_state', '数据不足'))} "
                    f"| 当前原始="
                    f"{cb.get('current_raw_state', cb.get('raw_state', '数据不足'))}"
                )
                print(
                    "   当前确认="
                    f"{'是' if cb.get('current_confirmed') else '否'} "
                    f"| 当前确认强度={cb.get('current_strength', 'NONE')} "
                    f"| 当前确认={cb.get('current_confirm_count', 0)}/{cb.get('current_confirm_required', CAPITAL_BEHAVIOR_CONFIRM_BARS)} "
                    f"| 当前资金分={safe_float(cb.get('current_score'), 0.0):+.1f}"
                )
                print(
                    "   3M确认="
                    f"{cb.get('current_3m_state', cb.get('confirm_3m_state', '数据不足'))} "
                    f"| 3M同向确认={'是' if cb.get('current_confirm_3m') else '否'} "
                    f"| 记忆防抖={cb.get('stability_text', 'N/A')}"
                )
                print(
                    "   10/20/30根状态="
                    f"{cb.get('window_10', {}).get('state', 'NA')}/"
                    f"{cb.get('window_20', {}).get('state', 'NA')}/"
                    f"{cb.get('window_30', {}).get('state', 'NA')} "
                    f"| 卖压分位={_cbfmt(cb.get('sell_percentile'), '%')} "
                    f"| 追涨分位={_cbfmt(cb.get('chase_percentile'), '%')}"
                )
                print(
                    "   吸收=" + _cbfmt(cb.get('absorption_score')) +
                    " | 推升效率=" + _cbfmt(cb.get('drive_efficiency')) +
                    " | 追涨=" + _cbfmt(cb.get('chase_score')) +
                    " | 派发=" + _cbfmt(cb.get('distribution_risk')) +
                    " | 砸盘=" + _cbfmt(cb.get('dump_risk')) +
                    " | 高位=" + ("是" if cb.get('near_high') else "否")
                )
                for cb_reason in cb.get("reasons", [])[:3]:
                    print("   → 资金行为：" + cb_reason)
            else:
                print("4A. 资金行为 V2: 数据不足")

            # --------------------------------------------------
            # 风控
            # --------------------------------------------------

            print(
                "5. 持仓风控: "
                +
                (
                    risk_message
                    if risk_message
                    else
                    "安全范围 / 未持仓"
                )
            )

            print(
                "-" * 78
            )

            # --------------------------------------------------
            # 最终评分
            # --------------------------------------------------

            print(
                f"【系统评分】: "
                f"{final_score} 分"
            )

            print(
                "【星级触发】: "
                +
                (
                    "★" * stars
                    if stars
                    else
                    "未触发"
                )
            )

            print(
                f"【交易建议】: "
                f"{advice}"
            )

            # --------------------------------------------------
            # 评分拆解
            # --------------------------------------------------

            print(
                "【评分拆解】: "
                f"基础50 "
                f"+ 1M={score_1m} "
                f"+ 15M={score_15m} "
                f"+ 30M={score_30m} "
                f"+ VWAP={score_vwap} "
                f"+ DAR={score_dar} "
                f"+ 资金行为V2={score_capital:+.1f} "
                f"+ 指数={score_market} "
                f"+ 月K={score_month} "
                f"+ 趋势层={score_trend:+.0f}"
            )

            print(
                "=" * 78
            )

            # ==================================================
            # 警报
            # ==================================================

            last_alert = (
                ALERTED_STOCKS.get(
                    code
                )
            )

            # --------------------------------------------------
            # ★★★★★
            # --------------------------------------------------

            if stars == 5:

                if last_alert != advice:

                    ALERTED_STOCKS[
                        code
                    ] = advice

                    speak_text_async(
                        f"注意！"
                        f"{name}"
                        f"触发五星级强买入信号"
                    )

                    if ask_human_confirmation_cmd(
                        code,
                        name
                    ):

                        ExcelManager.write_signal(
                            {
                                "代码":
                                    code,

                                "交易信号":
                                    advice,

                                "信号时间":
                                    timestamp,

                                "综合评分":
                                    final_score
                            }
                        )

            # --------------------------------------------------
            # ★★★★
            # --------------------------------------------------

            elif stars == 4:

                if last_alert != advice:

                    ALERTED_STOCKS[
                        code
                    ] = advice

                    speak_text_async(
                        f"提示，"
                        f"{name}"
                        f"触发四星级推荐"
                    )

                    ExcelManager.write_signal(
                        {
                            "代码":
                                code,

                            "交易信号":
                                advice,

                            "信号时间":
                                timestamp,

                            "综合评分":
                                final_score
                        }
                    )

            # --------------------------------------------------
            # 低于四星
            # --------------------------------------------------

            elif stars < 4:

                ALERTED_STOCKS.pop(
                    code,
                    None
                )

        except Exception as e:

            logger.exception(
                f"扫描股票 "
                f"{row.get('代码', '')} "
                f"异常: {e}"
            )


# ======================================================================
# 主监控
# ======================================================================

def start_continuous_monitoring(
    interval_seconds=INTERVAL_SECONDS,
    force_run=False
):

    tdx = TDXServerPool()

    if not tdx.connect():

        return

    analyzer = MarketAnalyzer(
        tdx
    )

    focus_terminal_window()

    logger.info(
        "=" * 60
    )

    logger.info(
        "StockPulse V7.8 Industry Behavior PRO Capital Behavior V2.2 + "
        "市场总闸门 + 五星连续确认 Stable FINAL"
    )
    logger.info(
        "V7.8 Industry Behavior：行业趋势 + 加速/减速 + 退潮 + 内部分化：V7.6核心逻辑保留 + TDX真实行业强度 + 独立最终交易权限"
    )
    logger.info(
        "三指数：999999上证 + 399001深证 + 399006创业板 | quote实时优先 + get_index_bars备用"
    )
    logger.info(
        "市场广度：TDX原生880005全市场涨跌家数 | Quote实时优先 + 1M统计指数K线备用 | 不拼接999999/399001 | 不枚举证券全集"
    )
    logger.info(
        "行业闸门：仅接受tdxhy.cfg/block_hy.dat | 禁止block.dat/block_gn.dat冒充行业 | 无可靠数据=UNKNOWN且不阻断"
    )

    logger.info(
        "999999 月K："
        "get_index_bars(category=6)"
    )

    logger.info(
        "999999 日K："
        "get_index_bars(category=4)"
    )

    logger.info(
        "999999 / 399001："
        "禁止使用get_security_bars"
    )

    logger.info(
        "指数压力："
        "前20个已完成日K高点"
    )

    logger.info(
        "指数突破历史压力："
        "压力距离=0"
    )

    logger.info(
        "月K："
        "排除当前未完成月K"
    )

    logger.info(
        "月K："
        "BEARISH禁止普通★★★★★；极强反转仅允许人工确认"
    )

    logger.info(
        "DAR："
        "当前量不计入自身量能均值"
    )

    logger.info(
        "趋势分层：月K MA20 + 周K MA20/60 + 日K MA20/60 "
        "→ A/B/C/D/E → 评分调整 + ENTRY_GATE V6.3 + V7.1总闸门"
    )

    logger.info(
        "评分："
        "重新压缩权重，降低100分饱和"
    )

    logger.info(
        "Capital Behavior V2.2："
        f"1M原始发现 + 3M已完成K确认 + 连续{CAPITAL_BEHAVIOR_CONFIRM_BARS}根防抖 + 状态滞回"
    )
    logger.info(
        "Capital Behavior V2.2退出阈值："
        f"吸收<{CB_EXIT_ABSORPTION:.1f} | 推升<{CB_EXIT_DRIVE_EFFICIENCY:.2f} | "
        f"追涨<{CB_EXIT_CHASE:.1f} | 派发<{CB_EXIT_DISTRIBUTION:.1f} | 砸盘<{CB_EXIT_DUMP:.1f}"
    )
    logger.info(
        f"轮询间隔: "
        f"{interval_seconds} 秒"
    )

    logger.info(
        "ENTRY_GATE V6.3: "
        f"{'开启' if ENTRY_GATE_ENABLED else '关闭'} | "
        f"局部压力回看={ENTRY_GATE_LOOKBACK}根1M | "
        f"贴压阈值={ENTRY_GATE_NEAR_PRESSURE_PCT:.2f}% | "
        f"突破确认={ENTRY_GATE_CONFIRM_BARS}根 | "
        f"当日突破缓冲={ENTRY_GATE_DAY_BREAKOUT_BUFFER_PCT:.2f}% | "
        f"失败回撤={ENTRY_GATE_DAY_FAILURE_RETRACE_PCT:.2f}%"
    )

    logger.info(
        "V7.6 PRO：市场五级状态 + 广度独立状态 + 五星时间确认="
        f"{V71_CONFIRM_SCANS}次 / 间隔{V76_FIVE_STAR_CONFIRM_INTERVAL_MIN}分钟 | "
        "14:00加强检查但不强制等待 | 月K长期风险与当日环境分离 | 板块/题材第二闸门开启"
    )

    logger.info(
        "V7.5阈值："
        f"市场BLOCK={V71_MARKET_SCORE_BLOCK} | 市场WAIT={V71_MARKET_SCORE_WAIT} | "
        f"指数跌幅BLOCK={V71_INDEX_CHANGE_BLOCK:.1f}% | "
        f"指数压力BLOCK={V71_INDEX_PRESSURE_BLOCK:.2f}%"
    )

    logger.info(
        "=" * 60
    )

    try:

        while True:

            in_trading, status = (
                is_trading_time()
            )

            if (
                in_trading
                or
                force_run
            ):

                logger.info(
                    f">>> 开始新一轮扫描 "
                    f"({datetime.datetime.now():%H:%M:%S}) <<<"
                )

                perform_scan_round(
                    tdx,
                    analyzer
                )

                time.sleep(
                    interval_seconds
                )

            else:

                logger.info(
                    f"当前非交易时间: "
                    f"[{status}]，"
                    f"等待30秒后重试..."
                )

                time.sleep(30)

    except KeyboardInterrupt:

        logger.info(
            "已手动停止监控流程。"
        )

    finally:

        try:

            tdx.api.disconnect()

        except Exception:

            pass


# ======================================================================
# MAIN
# ======================================================================

if __name__ == "__main__":

    start_continuous_monitoring(
        interval_seconds=30,
        force_run=False
    )
