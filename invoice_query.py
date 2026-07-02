# -*- coding: utf-8 -*-
"""
数金数据查询工具
模块：
  1. 发票查询        （单个 / 批量Excel）
  2. 不动产查询      （申请人直接写库链路：申请提交 / 查询申请记录 / Excel补录 /
                      编辑资产结果 / 处理回调）
  3. 公积金收入评分  （单个 / 批量Excel / 历史记录）
所有接口共用 Authorization（Bearer Token），自动补 “Bearer ” 前缀。

不动产采用“申请人直接写库”链路（/search/assets/application）：
  ① 申请提交：仅提交申请人 {persons:[{name,cardNum}]}，服务端直接写库并内部生成
     reqOrderNo（MANUAL-...），初始状态“进行中”；
  ② 查询申请记录：GET /search/assets/ 拿到 reqOrderNo / 状态 / 已补录资产；
  ③ Excel补录：上传 Excel 为申请记录补录不动产资产结果，状态更新为“成功”；
  ④ 编辑资产结果：按 assetId 修改单条资产明细；
  ⑤ 处理回调：对请求单号数组调用 /callback/handle 触发回调。
"""
import os
import re
import sys
import json
import time
import base64
import hashlib
import threading
import traceback
from datetime import datetime
from urllib.parse import quote

import requests
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

BASE_URL = "https://api.dlszjr.com"

# 发票
INV_SUBMIT_PATH = "/api/invoices/query"
INV_QUERY_PATH = "/api/invoices/query/{id}"
# 公积金评分
GJJ_SUBMIT_PATH = "/api/gjj-income-score/query"
GJJ_QUERY_PATH = "/api/gjj-income-score/query/{id}"
GJJ_HISTORY_PATH = "/api/gjj-income-score/history"
# 不动产（申请人直接写库链路 /search/assets/application）
RE_APPLICATION_PATH = "/search/assets/application"       # ① 申请提交（直接写库）POST
RE_RECORDS_PATH = "/search/assets/"                      # ② 查询申请记录 GET
RE_UPLOAD_PATH = "/search/assets/application/upload"     # ③ Excel补录资产结果 POST(multipart)
RE_RESULT_UPDATE_PATH = "/search/assets/result/update"   # ④ 编辑单条资产 POST
RE_CALLBACK_HANDLE_PATH = "/callback/handle"             # ⑤ 处理回调 POST(请求单号数组)

# ---------------- 探知数据（上海勃池）----------------
# 与上面三个接口认证方式不同：使用 apiKey + apiSecret 签名（SHA-1），form 表单提交。
TANZHI_BASE_URL = "https://api.tanzhishuju.com/api/gateway"
TANZHI_VERSION = "1.0.0"

# 每个产品：key/显示名/method/数据字典sheet名(无则None)/身份证映射字段/姓名是否必填/姓名是否使用/手机号是否使用
TANZHI_PRODUCTS = [
    # 名下车辆状况分析按 usertype（关系类型）拆成两个产品
    {"key": "vehicle_analysis_etc", "name": "名下车辆状况分析(ETC开户人)", "method": "api.vehicle.analysis",
     "sheet": None, "id_field": "identityNo", "name_req": 0, "name_use": 1, "mobile_use": 0,
     "fixed": {"usertype": "1"}},   # usertype=1 ETC开户人
    {"key": "vehicle_analysis_owner", "name": "名下车辆状况分析(车辆所有人)", "method": "api.vehicle.analysis",
     "sheet": None, "id_field": "identityNo", "name_req": 0, "name_use": 1, "mobile_use": 0,
     "fixed": {"usertype": "2"}},   # usertype=2 车辆所有人
    {"key": "car_totalvalue", "name": "名下车辆总价值", "method": "api.car.totalvalue",
     "sheet": None, "id_field": "idnum", "name_req": 0, "name_use": 0, "mobile_use": 0},
    # 车牌类接口（合并产品内部调用的三个接口，单独也开放为查询产品）：只需车牌[/VIN]，与
    # 身份证类产品输入不同，input_mode="plate" 时复用“姓名”框填VIN、“身份证号”框填车牌。
    {"key": "car_fiveinfo", "name": "车五项信息查询", "method": "api.car.fiveinfo",
     "sheet": None, "input_mode": "plate", "vin_field": None},
    {"key": "car_valueinfo", "name": "车辆价值查询", "method": "api.car.valueinfo",
     "sheet": None, "input_mode": "plate", "vin_field": "vin"},      # vin 小写，与车牌均必填
    {"key": "business_validate", "name": "商业保险有效性判定", "method": "api.business.validate",
     "sheet": None, "input_mode": "plate", "vin_field": "vin"},      # 实测服务端要求小写vin
     # ↑ 文档表格写的是大写VIN，但实测传VIN会返回"vin不能为空"，服务端实际按小写vin校验，
     #   与车辆价值查询一致；以实测为准，均传小写与车牌均必填
    {"key": "multiple_lending", "name": "多头借贷指数", "method": "api.multiple.lending",
     "sheet": "多头借贷指数", "id_field": "identityNo", "name_req": 1, "name_use": 1, "mobile_use": 1},
    {"key": "payment_index", "name": "支付行为指数", "method": "api.payment.index",
     "sheet": "支付行为指数", "id_field": "identityNo", "name_req": 1, "name_use": 1, "mobile_use": 1},
    {"key": "panoramic_index", "name": "全景指数", "method": "api.panoramic.index",
     "sheet": "全景指数", "id_field": "identityNo", "name_req": 1, "name_use": 1, "mobile_use": 1},
    {"key": "spend_transaction", "name": "消费交易特征", "method": "api.spend.transaction",
     "sheet": "消费交易特征", "id_field": "identityNo", "name_req": 1, "name_use": 1, "mobile_use": 1},
    {"key": "specific_risk", "name": "特定风险名单", "method": "api.specific.risk",
     "sheet": "特定风险名单", "id_field": "identityNo", "name_req": 1, "name_use": 1, "mobile_use": 1},
    {"key": "credit_index", "name": "信用指数", "method": "api.credit.index",
     "sheet": "信用指数", "id_field": "identityNo", "name_req": 1, "name_use": 1, "mobile_use": 1},
    # 合并/串联产品：身份证→名下车辆→逐车 车估值 + 商业险有效性
    {"key": "vehicle_combo", "name": "名下车辆估值+商业险", "method": None,
     "sheet": None, "id_field": "identityNo", "name_req": 0, "name_use": 1, "mobile_use": 0,
     "combo": True},
]
TANZHI_PRODUCT_BY_NAME = {p["name"]: p for p in TANZHI_PRODUCTS}

# 串联产品的方法名
TZ_M_VEHICLE_ANALYSIS = "api.vehicle.analysis"
TZ_M_CAR_FIVEINFO = "api.car.fiveinfo"
TZ_M_CAR_VALUEINFO = "api.car.valueinfo"
TZ_M_BUSINESS_VALIDATE = "api.business.validate"

# 通用状态码字典：结果表/导出里但凡字段名是 code / xxx_code，都用这个把裸数字翻成中文
TANZHI_STATUS_CODES = {
    "0000": "查得（正常，计费）",
    "1012": "服务器异常",
    "1013": "参数格式错误（必填参数缺失或格式不对，如缺 VIN/车牌）",
    "1015": "账户余额不足",
    "8000": "数据通道不可用",
    "8001": "数据通道返回格式异常",
    "8002": "查询成功，未查得数据（不计费）",
    "8003": "内部处理错误",
    "SKIP": "本工具未获取到该车 VIN，判断必填参数不全，主动跳过此接口（未产生调用/计费）",
    "异常": "请求过程中发生网络/程序异常",
}


def tanzhi_code_note(key, val):
    """字段名以 code 结尾（如 code / xxx_code）时，把状态码值翻译成中文说明。"""
    last = re.sub(r"\[\d+\]", "", str(key)).split(".")[-1]
    if last == "code" or last.endswith("_code"):
        sval = str(val)
        meaning = TANZHI_STATUS_CODES.get(sval, "")
        if meaning:
            return "%s：%s" % (sval, meaning)
        if sval not in ("", "None"):
            return "%s：未在文档状态码表中列出" % sval
    return ""


# 合并产品聚合字段的中文注释（喂给通用 flatten+annotate 渲染/导出）
TANZHI_COMBO_DICT = {
    "vehicleCount": {"cn": "车辆数", "desc": "合并ETC开户人/车辆所有人并按车牌去重后的总数", "note": ""},
    "vehicles.source": {"cn": "数据来源",
                        "desc": "该车牌来自哪个关系类型：ETC开户人(usertype=1) / 车辆所有人(usertype=2)；两者都有则都列出",
                        "note": ""},
    "vehicles.plateNum": {"cn": "车牌号", "desc": "名下车辆状况分析返回的车牌", "note": ""},
    "vehicles.plateColor": {"cn": "车牌颜色",
                            "desc": "0蓝 1黄 2黑 3白 4渐变绿 5黄绿渐变 6蓝白渐变 7临时牌照 9未确定 11绿 12红",
                            "note": ""},
    "vehicles.vehicleType": {"cn": "车辆类型",
                             "desc": "0未确定 1-4型客车 11-16型货车 21-26专项作业车", "note": ""},
    "vehicles.vin": {"cn": "VIN码(车架号)", "desc": "由“车五项信息查询”按车牌反查得到", "note": ""},
    "vehicles.fiveinfo_code": {"cn": "车五项查询状态码", "desc": "", "note": ""},
    "vehicles.brand": {"cn": "车辆品牌", "desc": "", "note": ""},
    "vehicles.clxh": {"cn": "车辆型号", "desc": "", "note": ""},
    "vehicles.fdjh": {"cn": "发动机唯一标识号", "desc": "", "note": ""},
    "vehicles.clfl": {"cn": "车辆类别", "desc": "如轿车、卡车等", "note": ""},
    "vehicles.usage": {"cn": "使用性质", "desc": "营运/非营运等", "note": ""},
    "vehicles.registerDate": {"cn": "首次上牌日期", "desc": "", "note": ""},
    "vehicles.vehicleValue": {"cn": "车辆估值(元)",
                              "desc": "车辆价值查询返回，四舍五入到万位；空=未查到或已跳过", "note": ""},
    "vehicles.vehicleValue_code": {"cn": "车估值接口状态码", "desc": "", "note": ""},
    "vehicles.businessInsuranceValid": {"cn": "商业保险有效性",
                                        "desc": "1=有效；0=无效；空=未查到或已跳过", "note": ""},
    "vehicles.businessInsurance_code": {"cn": "商业险接口状态码", "desc": "", "note": ""},
}

# 车辆两个接口数据字典未在标注表里，这里补充（best-effort），并作为全局兜底注释
TANZHI_EXTRA_DICT = {
    "vehicleCount": {"cn": "车辆数", "desc": "", "note": ""},
    "list.plateNum": {"cn": "匿名化车牌号", "desc": "", "note": ""},
    "list.plateColor": {"cn": "车牌颜色",
                        "desc": "0蓝 1黄 2黑 3白 4渐变绿 5黄绿渐变 6蓝白渐变 7临时牌照 9未确定 11绿 12红",
                        "note": ""},
    "list.vehicleType": {"cn": "车辆类型",
                         "desc": "0未确定 1-4型客车 11-16型货车 21-26专项作业车", "note": ""},
    "result": {"cn": "查询结果", "desc": "1：查询成功(计费)；-1：查询失败(不计费)", "note": ""},
    "resultdata.vehicleCountInterval": {"cn": "车辆数量区间",
                                        "desc": "2：10辆及以上；1：10辆以内", "note": ""},
    "resultdata.vehicleValue": {"cn": "车辆估值/总价值(元)",
                                "desc": "四舍五入到万；名下车辆总价值时10辆以上不返回具体值", "note": ""},
    "resultdata.vin": {"cn": "VIN码(车架号)", "desc": "", "note": ""},
    "resultdata.brand": {"cn": "车辆品牌", "desc": "", "note": ""},
    "resultdata.clxh": {"cn": "车辆型号", "desc": "", "note": ""},
    "resultdata.fdjh": {"cn": "发动机唯一标识号", "desc": "", "note": ""},
    "resultdata.clfl": {"cn": "车辆类别", "desc": "如轿车、卡车等", "note": ""},
    "resultdata.usage": {"cn": "使用性质", "desc": "营运/非营运等", "note": ""},
    "resultdata.registerDate": {"cn": "首次上牌日期", "desc": "", "note": ""},
    "resultdata": {"cn": "商业保险有效性",
                   "desc": "仅商业保险有效性判定接口：1=有效；0=无效", "note": ""},
    "code": {"cn": "状态码", "desc": "", "note": ""},
    "msg": {"cn": "状态描述", "desc": "", "note": ""},
    "token": {"cn": "流程唯一标识", "desc": "", "note": ""},
}

POLL_INTERVAL = 3       # 轮询间隔（秒）
POLL_MAX_TIMES = 40     # 单条最大轮询次数（约2分钟）
REQUEST_TIMEOUT = 30    # 普通HTTP请求超时
UPLOAD_TIMEOUT = 180    # 带文件上传的请求超时

CONFIG_FILE = os.path.join(
    os.path.expanduser("~"), ".invoice_query_tool_config.json"
)

HEADER_FILL = "4F81BD"

# ---------------- 发票字段中文映射 ----------------
FIELD_LABELS = {
    "id": "查询ID", "success": "成功标志", "nsrsbh": "纳税人识别号",
    "djxh": "登记序号", "xzqhDm": "行政区划代码", "ssqq": "属期起", "ssqz": "属期止",
    "dlhyDm": "大类行业代码", "nsrmc": "纳税人名称", "dlhymc": "大类行业名称",
    "zzsnsrlxId": "增值税纳税人类型", "djzclxmc": "登记注册类型名称", "minKprq": "最小开票日期",
    "nsrsh": "纳税人识别号", "kprqMinMonth": "实际经营月份数", "nsrztmc": "纳税人状态",
    "scjydz": "生产经营地址", "djzclxDm": "登记注册类型代码", "nsrlx": "纳税人类型",
    "nsrxydjId": "纳税信用等级", "scjyddhhm": "生产经营地电话号码",
    "dykptsLp": "当月开票天数(蓝票)", "dykptsQb": "当月开票天数(全部)", "fpje": "负票金额",
    "fpsl": "负票数量", "hpje": "红票金额", "hpsl": "红票数量", "yxhpje": "有效红票金额",
    "yxhpsl": "有效红票数量", "kpje": "开票金额", "kpqj": "开票期间", "kpsl": "开票数量",
    "kpyf": "开票月份", "lxwjyts": "当前连续无交易记录天数", "zddzje": "单次最高开票金额",
    "zjkpsj": "最近一笔开票时间", "wsje": "开票不含税金额", "hpwsje": "红票不含税金额",
    "yxhpwsje": "有效红票不含税金额", "fpwsje": "负票不含税金额",
    "gfDlhyDm": "购方_大类行业代码", "gfDlhymc": "购方_大类行业名称", "xfhyId": "销方行业代码",
    "xfCount": "销方数量", "xfmc": "销方名称", "xfDjxh": "销方登记序号", "xfnsrdzdah": "销方电子档案号",
    "zcjeCnt": "开票次数", "zcjeCntRate": "开票次数占比", "zcjeSum": "开票金额",
    "zcjeSumRank": "排名", "zcjeSumRate": "开票金额占比", "wsjeSum": "开票不含税金额",
    "gfCount": "购方数量", "gfDjxh": "购方登记序号", "gfhyId": "购方行业代码", "gfmc": "购方名称",
    "xfDlhyDm": "销方_大类行业代码", "xfDlhymc": "销方_大类行业名称",
    "gfdsswjgId": "购方主管税务机关", "gfsCount": "购方主管税务机关数量",
    "hwbmCount": "货物名称数量", "jeSum": "商品总金额", "jeSumRank": "交易金额占比排名",
    "seSum": "商品税额", "slSum": "商品数量", "slv": "税率", "spbm": "商品编码", "spmc": "商品名称",
}

SECTION_DEFS = [
    ("qyxxList", "企业信息", [
        "nsrsh", "nsrmc", "djxh", "nsrztmc", "nsrlx", "zzsnsrlxId",
        "djzclxDm", "djzclxmc", "dlhyDm", "dlhymc",
        "minKprq", "kprqMinMonth", "nsrxydjId", "scjydz", "scjyddhhm",
    ]),
    ("kphzxxList", "开票汇总", [
        "djxh", "kpqj", "kpyf", "dlhyDm", "dlhymc", "kpje", "kpsl", "wsje",
        "fpje", "fpsl", "fpwsje", "hpje", "hpsl", "hpwsje",
        "yxhpje", "yxhpsl", "yxhpwsje",
        "dykptsLp", "dykptsQb", "lxwjyts", "zddzje", "zjkpsj",
    ]),
    ("syKphzxxList", "上游开票分类", [
        "djxh", "kpqj", "kpyf", "xfDjxh", "xfmc", "xfhyId", "gfDlhyDm", "gfDlhymc",
        "xfCount", "zcjeCnt", "zcjeCntRate",
        "zcjeSum", "zcjeSumRate", "zcjeSumRank", "wsjeSum", "xfnsrdzdah",
    ]),
    ("xyKphzxxList", "下游开票分类", [
        "djxh", "kpqj", "kpyf", "gfDjxh", "gfmc", "gfhyId", "xfDlhyDm", "xfDlhymc",
        "gfCount", "zcjeCnt", "zcjeCntRate",
        "zcjeSum", "zcjeSumRate", "zcjeSumRank", "wsjeSum", "xfnsrdzdah",
    ]),
    ("khxsdqxxList", "客户销售地区", [
        "djxh", "kpqj", "kpyf", "gfdsswjgId", "gfsCount", "zcjeCnt", "zcjeCntRate",
        "zcjeSum", "zcjeSumRate", "zcjeSumRank", "wsjeSum", "xfnsrdzdah",
    ]),
    ("spxsxxList", "商品销售", [
        "djxh", "kpyf", "spbm", "spmc", "slv", "hwbmCount", "slSum", "jeSum", "seSum",
        "jeSumRank", "zcjeCntRate", "zcjeSumRate", "xfnsrdzdah",
    ]),
]

# ---------------- 公积金评分字段 ----------------
GJJ_DATA_FIELDS = [
    ("status", "处理状态"), ("score", "评分结果"), ("scoreRange", "评分区间"),
    ("jfzt", "缴存状态编码"), ("jfztText", "缴存状态说明"), ("date", "缴存月份"),
    ("success", "第三方成功"), ("thirdPartyCode", "第三方返回码"),
    ("thirdPartyMessage", "第三方返回说明"), ("seqNo", "第三方流水号"),
    ("free", "计费标识原值"), ("charged", "是否计费"), ("fileIndex", "文件索引"),
    ("elapsedMs", "耗时(ms)"), ("id", "受理ID"),
    ("createdDate", "创建时间"), ("modifiedDate", "更新时间"),
]

# ---------------- 不动产字段（直接写库链路） ----------------
# 提交（直接写库）返回的单条结果 data[]
RE_SUBMIT_FIELDS = [
    ("name", "姓名"), ("cardNum", "身份证号"),
    ("isSuccess", "是否成功"), ("code", "单条响应码"), ("msg", "写入结果说明"),
]
# 查询申请记录 GET /search/assets/ 的 data[]
RE_RECORD_FIELDS = [
    ("reqOrderNo", "请求单号"), ("name", "姓名"), ("cardNum", "身份证号"),
    ("status", "申请状态"), ("callbacked", "是否已回调"),
    ("callbackId", "回调配置ID"), ("id", "申请记录ID"),
]
# 资产明细字段（assets[] / 编辑资产）
RE_ASSET_FIELDS = [
    ("certNo", "房产证号"), ("unitNo", "房产单元号"), ("ownership", "权利人"),
    ("rightsType", "权利类型"), ("useTo", "用途"), ("houseArea", "房屋面积"),
    ("location", "房屋坐落"), ("isSealUp", "是否查封"), ("isMortgaged", "是否抵押"),
]
# 处理回调 /callback/handle 的 data[]
RE_CALLBACK_RESULT_FIELDS = [
    ("reqOrderNo", "请求单号"), ("success", "是否成功"),
    ("errorMessage", "失败原因"), ("callbackId", "回调配置ID"),
    ("processTime", "处理时间"),
]
# Excel补录模板列（前3列必填，须与申请记录一致；其余为资产结果列，best-effort）
RE_UPLOAD_TEMPLATE_COLS = (
    ["Name", "CardNum", "ReqOrderNo"] + [k for k, _ in RE_ASSET_FIELDS]
)


# ---------------- 配置文件 ----------------
def load_config():
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _resource_path(rel):
    """兼容 PyInstaller onefile：优先从解压目录(_MEIPASS)取，否则取脚本同目录。"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def load_tanzhi_dict():
    """加载探知数据字段标注字典：{sheet: {code: {cn,desc,note}}}。"""
    try:
        with open(_resource_path("tanzhi_field_dict.json"), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def normalize_auth(value):
    """未带常见scheme时默认按 Bearer 拼接。"""
    s = (value or "").strip()
    if not s:
        return ""
    if s.lower().startswith(("bearer ", "basic ", "token ", "jwt ")):
        return s
    return "Bearer " + s


# ---------------- 网络客户端 ----------------
class BaseClient:
    def __init__(self, authorization):
        self.authorization = normalize_auth(authorization)

    def _headers(self, content_type="application/json"):
        h = {"Content-Type": content_type, "Accept": "application/json"}
        if self.authorization:
            h["Authorization"] = self.authorization
        return h

    def _post(self, path, payload, timeout=REQUEST_TIMEOUT):
        r = requests.post(
            BASE_URL + path, headers=self._headers(),
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            timeout=timeout,
        )
        r.raise_for_status()
        return r.json()

    def _get(self, path, timeout=REQUEST_TIMEOUT):
        r = requests.get(BASE_URL + path, headers=self._headers(), timeout=timeout)
        r.raise_for_status()
        return r.json()


class InvoiceClient(BaseClient):
    def submit(self, sh):
        return self._post(INV_SUBMIT_PATH, [{"sh": sh}])

    def query(self, query_id):
        return self._get(INV_QUERY_PATH.format(id=query_id))

    def query_one(self, sh, on_status=None):
        info = {"sh": sh, "status": "", "code": "", "message": "", "id": ""}
        try:
            submit_resp = self.submit(sh)
        except Exception as e:
            info["message"] = f"提交失败: {e}"
            return False, info, None
        if not isinstance(submit_resp, dict):
            info["message"] = f"提交返回非预期: {submit_resp}"
            return False, info, None
        if str(submit_resp.get("status")) != "1":
            info["message"] = submit_resp.get("message") or "提交失败"
            info["code"] = submit_resp.get("code", "")
            return False, info, None
        data_arr = submit_resp.get("data") or []
        if not data_arr:
            info["message"] = "提交返回data为空"
            return False, info, None
        first = data_arr[0]
        query_id = first.get("id")
        info["id"] = query_id or ""
        if not query_id:
            info["message"] = first.get("msg") or "未拿到查询ID"
            info["code"] = first.get("code", "")
            return False, info, None
        last_msg = ""
        for i in range(POLL_MAX_TIMES):
            if on_status:
                on_status(f"正在查询 {sh} (尝试 {i+1}/{POLL_MAX_TIMES})...")
            try:
                resp = self.query(query_id)
            except Exception as e:
                last_msg = f"查询失败: {e}"
                time.sleep(POLL_INTERVAL)
                continue
            code = str(resp.get("code", ""))
            info["code"] = code
            info["message"] = resp.get("message") or ""
            data = resp.get("data") or {}
            data_status = (data.get("status") or "") if isinstance(data, dict) else ""
            if code == "200000":
                info["status"] = "成功"
                return True, info, resp
            elif code == "202000" or data_status == "Processing":
                last_msg = "处理中"
                time.sleep(POLL_INTERVAL)
                continue
            elif code == "204000":
                info["status"] = "无数据"
                info["message"] = info["message"] or "查到查询记录但没有明细数据"
                return False, info, resp
            elif code == "500000":
                info["status"] = "失败"
                info["message"] = info["message"] or "查询失败"
                return False, info, resp
            else:
                last_msg = f"未知响应 code={code}"
                time.sleep(POLL_INTERVAL)
                continue
        info["status"] = "超时"
        info["message"] = last_msg or "轮询超时"
        return False, info, None


class GjjClient(BaseClient):
    def submit(self, name, idcard, mobile=""):
        payload = {"name": name, "idCard": idcard, "mobile": mobile or ""}
        return self._post(GJJ_SUBMIT_PATH, payload)

    def query(self, query_id):
        return self._get(GJJ_QUERY_PATH.format(id=query_id))

    def history(self, page_index=1, page_size=10):
        path = f"{GJJ_HISTORY_PATH}?pageIndex={page_index}&pageSize={page_size}"
        return self._get(path)

    def query_one(self, name, idcard, mobile="", on_status=None):
        info = {"name": name, "idCard": idcard, "mobile": mobile,
                "status": "", "code": "", "message": "", "id": ""}
        try:
            sub = self.submit(name, idcard, mobile)
        except Exception as e:
            info["message"] = f"提交失败: {e}"
            return False, info, None
        if not isinstance(sub, dict):
            info["message"] = f"提交返回非预期: {sub}"
            return False, info, None
        info["code"] = str(sub.get("code", ""))
        if str(sub.get("status")) != "1":
            info["message"] = sub.get("message") or "提交失败"
            return False, info, None
        data = sub.get("data") or {}
        qid = data.get("id") if isinstance(data, dict) else None
        info["id"] = qid or ""
        if not qid:
            info["message"] = sub.get("message") or "未拿到受理ID"
            return False, info, None
        last_msg = ""
        for i in range(POLL_MAX_TIMES):
            if on_status:
                on_status(f"正在查询 {name} (尝试 {i+1}/{POLL_MAX_TIMES})...")
            try:
                resp = self.query(qid)
            except Exception as e:
                last_msg = f"查询失败: {e}"
                time.sleep(POLL_INTERVAL)
                continue
            d = resp.get("data") or {}
            st = (d.get("status") or "") if isinstance(d, dict) else ""
            info["code"] = str(resp.get("code", ""))
            info["message"] = resp.get("message") or ""
            if st in ("Pending", "Processing"):
                last_msg = "处理中"
                time.sleep(POLL_INTERVAL)
                continue
            elif st == "Succeeded":
                info["status"] = "成功"
                return True, info, resp
            elif st == "NoData":
                info["status"] = "无数据"
                info["message"] = (d.get("thirdPartyMessage")
                                   or info["message"] or "第三方无匹配结果")
                return False, info, resp
            elif st == "Failed":
                info["status"] = "失败"
                info["message"] = (d.get("thirdPartyMessage")
                                   or info["message"] or "查询失败")
                return False, info, resp
            else:
                last_msg = f"未知状态 {st}"
                time.sleep(POLL_INTERVAL)
                continue
        info["status"] = "超时"
        info["message"] = last_msg or "轮询超时"
        return False, info, None


class RealEstateClient(BaseClient):
    """不动产·申请人直接写库链路客户端。鉴权：Bearer（与发票/公积金共用 token）。"""

    def submit_application(self, persons):
        """① 申请提交（直接写库）：body = {"persons":[{name,cardNum}]}"""
        return self._post(RE_APPLICATION_PATH, {"persons": persons})

    def query_records(self, page_num=1, page_size=10, search_text=""):
        """② 查询申请记录，返回 PagedDto。"""
        path = f"{RE_RECORDS_PATH}?pageNum={int(page_num)}&pageSize={int(page_size)}"
        if search_text:
            path += "&searchText=" + quote(search_text)
        return self._get(path)

    def upload_excel(self, file_path, req_order_no=""):
        """③ Excel补录资产结果（multipart/form-data）。"""
        headers = {"Accept": "application/json"}
        if self.authorization:
            headers["Authorization"] = self.authorization
        with open(file_path, "rb") as fh:
            files = {"file": (os.path.basename(file_path), fh,
                              "application/octet-stream")}
            data = {}
            if req_order_no:
                data["reqOrderNo"] = req_order_no
            r = requests.post(BASE_URL + RE_UPLOAD_PATH, headers=headers,
                              files=files, data=data, timeout=UPLOAD_TIMEOUT)
        r.raise_for_status()
        try:
            return r.json()
        except Exception:
            return {"_raw_text": r.text}

    def update_asset(self, payload):
        """④ 编辑单条资产结果：payload 含 assetId(必填) 及可选字段。"""
        return self._post(RE_RESULT_UPDATE_PATH, payload)

    def handle_callback(self, order_list):
        """⑤ 处理回调：请求体是请求单号数组 ["MANUAL-...."]。"""
        return self._post(RE_CALLBACK_HANDLE_PATH, order_list)


# ---------------- 探知数据 客户端 / 解析 ----------------
def tanzhi_sign(params, api_secret):
    """对除 sign 外的参数按 key 字典升序，拼成 k=v&k=v...，末尾直接追加 apisecret，SHA-1 小写hex。"""
    items = sorted((k, v) for k, v in params.items()
                   if k != "sign" and v is not None and v != "")
    raw = "&".join("%s=%s" % (k, v) for k, v in items) + (api_secret or "")
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


class TanzhiClient:
    def __init__(self, api_key, api_secret):
        self.api_key = (api_key or "").strip()
        self.api_secret = (api_secret or "").strip()

    def call(self, method, biz_params, timeout=REQUEST_TIMEOUT):
        params = {"apiKey": self.api_key, "version": TANZHI_VERSION, "method": method}
        for k, v in biz_params.items():
            if v is not None and str(v).strip() != "":
                params[k] = str(v).strip()
        params["sign"] = tanzhi_sign(params, self.api_secret)
        r = requests.post(
            TANZHI_BASE_URL, data=params, timeout=timeout,
            headers={"Content-Type": "application/x-www-form-urlencoded"})
        r.raise_for_status()
        try:
            return r.json()
        except Exception:
            return {"code": "", "msg": "返回非JSON", "raw": r.text}

    def query_one(self, product, name, idnum, mobile=""):
        """按产品定义组装业务参数并发起查询，返回 (info, resp)。
        input_mode=="plate" 的车牌类产品：复用 idnum 参数承载车牌号、name 参数承载 VIN。"""
        biz = {}
        if product.get("input_mode") == "plate":
            biz["plateno"] = idnum
            vf = product.get("vin_field")
            if vf and name:
                biz[vf] = name
        else:
            if product.get("name_use"):
                biz["name"] = name
            biz[product["id_field"]] = idnum
            if product.get("mobile_use"):
                biz["mobile"] = mobile
            biz.update(product.get("fixed", {}))   # 固定附加入参（如车辆状况分析 usertype=2）
        info = {"product": product["name"], "name": name, "idnum": idnum,
                "mobile": mobile, "code": "", "msg": "", "token": ""}
        try:
            resp = self.call(product["method"], biz)
        except Exception as e:
            info["msg"] = "请求失败: %s" % e
            return info, None
        if isinstance(resp, dict):
            info["code"] = str(resp.get("code", ""))
            info["msg"] = str(resp.get("msg", ""))
            info["token"] = str(resp.get("token", ""))
        return info, resp

    def query_chain(self, product, name, idnum, mobile=""):
        """串联查询：身份证→名下车辆状况分析→车五项信息查询(取VIN)→车估值+商业险有效性。
        step1 同时查 usertype=1(ETC开户人) 与 usertype=2(车辆所有人)，按车牌合并去重，
        并在每辆车上标注来源；然后按车牌查车五项信息拿到 VIN，再用 VIN+车牌去查
        车估值与商业险有效性均要求 VIN+车牌同时提供（均用小写vin，商业险文档虽写大写VIN，
        但实测服务端按小写vin校验，以实测为准）。
        若某车牌查不到 VIN，跳过后两步并标记 SKIP。
        返回 (info, synthetic_resp)，走通用 flatten/annotate/导出。"""
        info = {"product": product["name"], "name": name, "idnum": idnum,
                "mobile": "", "code": "", "msg": "", "token": ""}
        raw = {"analysis": {}, "fiveinfo": [], "valueinfo": [], "validate": []}
        # step1 两个关系类型都查，按车牌合并
        sources = [("1", "ETC开户人"), ("2", "车辆所有人")]
        merged, order, src_map = {}, [], {}
        analysis_codes, token = [], ""
        any_request_ok = False
        for ut, label in sources:
            try:
                r1 = self.call(TZ_M_VEHICLE_ANALYSIS,
                               {"name": name, "identityNo": idnum, "usertype": ut})
                any_request_ok = True
            except Exception as e:
                r1 = {"code": "异常", "msg": str(e)}
            raw["analysis"][label] = r1
            if isinstance(r1, dict):
                analysis_codes.append(str(r1.get("code", "")))
                if not token and r1.get("token"):
                    token = str(r1.get("token"))
            d1 = tanzhi_extract_data(r1) or {}
            vlist = d1.get("list") if isinstance(d1, dict) else None
            if isinstance(vlist, list):
                for v in vlist:
                    if not isinstance(v, dict):
                        continue
                    plate = str(v.get("plateNum", "") or "").strip()
                    key = plate or ("__noplate_%d" % len(order))
                    if key not in merged:
                        merged[key] = {"source": "", "plateNum": plate,
                                       "vehicleType": v.get("vehicleType", ""),
                                       "plateColor": v.get("plateColor", ""),
                                       "vin": "", "fiveinfo_code": "",
                                       "brand": "", "clxh": "", "fdjh": "", "clfl": "",
                                       "usage": "", "registerDate": "",
                                       "vehicleValue": "", "vehicleValue_code": "",
                                       "businessInsuranceValid": "", "businessInsurance_code": ""}
                        order.append(key)
                        src_map[key] = []
                    if label not in src_map[key]:
                        src_map[key].append(label)
        if not any_request_ok:
            info["msg"] = "请求失败: 名下车辆状况分析调用异常"
            return info, None
        # 整理来源标注 + 逐车（去重后）查 车估值 / 商业险
        vehicles = []
        for key in order:
            item = merged[key]
            item["source"] = "、".join(src_map[key])
            plate = item["plateNum"]
            if plate:
                # step2 车五项信息查询（仅传车牌）→ 拿 VIN + 顺带的品牌/型号等信息
                try:
                    r_five = self.call(TZ_M_CAR_FIVEINFO, {"plateno": plate})
                except Exception as e:
                    r_five = {"code": "异常", "msg": str(e)}
                raw["fiveinfo"].append(r_five)
                item["fiveinfo_code"] = str(r_five.get("code", "")) if isinstance(r_five, dict) else ""
                d_five = tanzhi_extract_data(r_five)
                vin = ""
                if isinstance(d_five, dict):
                    rd_five = d_five.get("resultdata")
                    if isinstance(rd_five, dict):
                        vin = str(rd_five.get("vin", "") or "").strip()
                        item["vin"] = vin
                        item["brand"] = rd_five.get("brand", "")
                        item["clxh"] = rd_five.get("clxh", "")
                        item["fdjh"] = rd_five.get("fdjh", "")
                        item["clfl"] = rd_five.get("clfl", "")
                        item["usage"] = rd_five.get("usage", "")
                        item["registerDate"] = rd_five.get("registerDate", "")
                if vin:
                    # step3 车辆价值查询（vin 小写 + plateno 均必填）
                    try:
                        r2 = self.call(TZ_M_CAR_VALUEINFO, {"vin": vin, "plateno": plate})
                    except Exception as e:
                        r2 = {"code": "异常", "msg": str(e)}
                    raw["valueinfo"].append(r2)
                    item["vehicleValue_code"] = str(r2.get("code", "")) if isinstance(r2, dict) else ""
                    d2 = tanzhi_extract_data(r2)
                    if isinstance(d2, dict):
                        rd2 = d2.get("resultdata")
                        if isinstance(rd2, dict):
                            item["vehicleValue"] = rd2.get("vehicleValue", "")
                    # step4 商业保险有效性判定（实测服务端要求小写vin + plateno 均必填；
                    # 文档表格写的是大写VIN，但实测传VIN会返回"vin不能为空"，以实测为准）
                    try:
                        r3 = self.call(TZ_M_BUSINESS_VALIDATE, {"vin": vin, "plateno": plate})
                    except Exception as e:
                        r3 = {"code": "异常", "msg": str(e)}
                    raw["validate"].append(r3)
                    item["businessInsurance_code"] = str(r3.get("code", "")) if isinstance(r3, dict) else ""
                    d3 = tanzhi_extract_data(r3)
                    if isinstance(d3, dict) and d3.get("resultdata") is not None:
                        item["businessInsuranceValid"] = d3.get("resultdata")
                else:
                    # 未取到 VIN：车估值/商业险必然报参数错误，主动跳过，避免无意义调用/计费
                    item["vehicleValue_code"] = "SKIP"
                    item["businessInsurance_code"] = "SKIP"
            vehicles.append(item)
        code = "0000" if "0000" in analysis_codes else (analysis_codes[0] if analysis_codes else "")
        info["code"] = code
        info["token"] = token
        info["msg"] = ("成功（%d辆，已合并ETC开户人/车辆所有人）" % len(vehicles)
                       if code == "0000" else
                       "；".join(TANZHI_STATUS_CODES.get(c, c) for c in analysis_codes))
        synthetic = {
            "code": code, "msg": info["msg"], "token": token,
            "data": {"vehicleCount": len(vehicles), "vehicles": vehicles},
            "_raw": raw,
        }
        return info, synthetic


def tanzhi_extract_data(resp):
    """取出 data 对象；若 data 是 JSON 字符串则解析。"""
    if not isinstance(resp, dict):
        return None
    data = resp.get("data")
    if isinstance(data, str):
        try:
            data = json.loads(data)
        except Exception:
            pass
    return data


def tanzhi_flatten(obj, prefix=""):
    """把嵌套 dict/list 拍平为 [(dotted_key, value)]，保持顺序。"""
    out = []
    if isinstance(obj, dict):
        for k, v in obj.items():
            np = "%s.%s" % (prefix, k) if prefix else str(k)
            out.extend(tanzhi_flatten(v, np))
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            out.extend(tanzhi_flatten(v, "%s[%d]" % (prefix, i)))
    else:
        out.append((prefix, obj))
    return out


def tanzhi_lookup(field_dict, product):
    """合并某产品的字段字典 + 兜底字典。"""
    d = dict(TANZHI_EXTRA_DICT)
    if product.get("combo"):
        d.update(TANZHI_COMBO_DICT)
    sheet = product.get("sheet")
    if sheet and isinstance(field_dict, dict):
        d.update(field_dict.get(sheet, {}))
    return d


def tanzhi_annotate(lookup, key):
    """对拍平后的 key 查注释：去掉[下标]，再从完整到末段逐级匹配。"""
    k = re.sub(r"\[\d+\]", "", key)
    if k in lookup:
        return lookup[k]
    parts = k.split(".")
    for i in range(1, len(parts)):
        cand = ".".join(parts[i:])
        if cand in lookup:
            return lookup[cand]
    return None


# ---------------- Excel 公共方法 ----------------
def _autosize_columns(ws, max_width=60):
    for col_cells in ws.columns:
        try:
            length = max(
                (len(str(c.value)) if c.value is not None else 0) for c in col_cells
            )
        except ValueError:
            length = 10
        col_letter = col_cells[0].column_letter
        ws.column_dimensions[col_letter].width = min(max(length + 2, 10), max_width)


def _align_left_all(ws):
    """把整张表所有已用单元格统一设为左对齐，避免文本/数字混排时对齐方式各异显得凌乱。"""
    align = Alignment(horizontal="left", vertical="center")
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = align


def _first_array_index(key):
    """从形如 vehicles[0].plateNum 的拍平键里取出数组下标(字符串)，非数组字段返回 None。"""
    m = re.search(r"\[(\d+)\]", key)
    return m.group(1) if m else None


def _style_header_row(ws, row_idx=1):
    fill = PatternFill("solid", fgColor=HEADER_FILL)
    for cell in ws[row_idx]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center")


# ---------------- 发票 导出/导入 ----------------
def export_results_to_excel(results, output_path):
    """results: List[Tuple[info_dict, raw_response_dict_or_None]]"""
    wb = Workbook()
    summary = wb.active
    summary.title = "查询汇总"
    summary.append(["纳税人识别号", "状态", "Code", "消息", "查询ID"])
    _style_header_row(summary)
    for info, _ in results:
        summary.append([
            info.get("sh", ""), info.get("status", ""), info.get("code", ""),
            info.get("message", ""), info.get("id", ""),
        ])
    _autosize_columns(summary)

    section_rows = {key: [] for key, _, _ in SECTION_DEFS}
    for info, resp in results:
        sh = info.get("sh", "")
        if not resp:
            continue
        data = resp.get("data") or {}
        if not isinstance(data, dict):
            continue
        for key, _name, _fields in SECTION_DEFS:
            arr = data.get(key) or []
            if not isinstance(arr, list):
                continue
            for item in arr:
                if not isinstance(item, dict):
                    continue
                row = dict(item)
                row["_sh"] = sh
                section_rows[key].append(row)

    for key, name, fields in SECTION_DEFS:
        rows = section_rows[key]
        if not rows:
            continue
        ws = wb.create_sheet(name[:30])
        ext_fields = ["_sh"] + fields
        ws.append(["纳税人识别号"] + [FIELD_LABELS.get(f, f) for f in fields])
        _style_header_row(ws)
        for row in rows:
            ws.append([row.get(f, "") for f in ext_fields])
        _autosize_columns(ws)
    wb.save(output_path)


def read_taxnos_from_excel(path):
    wb = load_workbook(path, data_only=True)
    out = []
    for ws in wb.worksheets:
        if ws.max_row == 0:
            continue
        first_row = [(c.value if c.value is not None else "") for c in ws[1]]
        col_idx = None
        for i, v in enumerate(first_row):
            sv = str(v).strip()
            if sv in ("纳税人识别号", "识别号", "sh", "SH", "税号", "纳税人识别号(必填)"):
                col_idx = i
                break
        start_row = 2 if col_idx is not None else 1
        if col_idx is None:
            col_idx = 0
        for row in ws.iter_rows(min_row=start_row, values_only=True):
            if not row:
                continue
            v = row[col_idx] if col_idx < len(row) else None
            if v is None:
                continue
            s = str(v).strip()
            if s and s not in out:
                out.append(s)
        if out:
            break
    return out


# ---------------- 公积金 导出/导入 ----------------
def export_gjj_to_excel(results, output_path):
    """results: List[Tuple[info_dict, raw_response_dict_or_None]]"""
    wb = Workbook()
    ws = wb.active
    ws.title = "公积金评分结果"
    headers = ["身份证号", "姓名", "手机号", "查询状态", "Code", "消息"] + \
        [lbl for _, lbl in GJJ_DATA_FIELDS]
    ws.append(headers)
    _style_header_row(ws)
    for info, resp in results:
        d = (resp.get("data") or {}) if resp else {}
        if not isinstance(d, dict):
            d = {}
        base = [
            info.get("idCard", ""), info.get("name", ""), info.get("mobile", ""),
            info.get("status", ""), info.get("code", ""), info.get("message", ""),
        ]
        rest = [d.get(k, "") for k, _ in GJJ_DATA_FIELDS]
        ws.append(base + rest)
    _autosize_columns(ws)
    wb.save(output_path)


def read_persons_from_excel(path):
    """读取 姓名/身份证号/手机号 三列，返回 [{'name','idCard','mobile'}]"""
    wb = load_workbook(path, data_only=True)
    out = []
    name_keys = ("姓名", "name", "名字", "查询人姓名")
    card_keys = ("身份证号", "身份证", "idcard", "cardnum", "证件号", "查询人身份证号")
    mob_keys = ("手机号", "手机", "mobile", "电话", "联系电话")
    for ws in wb.worksheets:
        if ws.max_row == 0:
            continue
        header = [str(c.value).strip().lower() if c.value is not None else "" for c in ws[1]]

        def find(keys):
            for i, h in enumerate(header):
                if h in keys:
                    return i
            return None
        i_name = find(name_keys)
        i_card = find(card_keys)
        i_mob = find(mob_keys)
        if i_name is None and i_card is None:
            i_name, i_card, i_mob = 0, 1, 2
            start = 1
        else:
            start = 2
        for row in ws.iter_rows(min_row=start, values_only=True):
            if not row:
                continue

            def cell(idx):
                if idx is None or idx >= len(row) or row[idx] is None:
                    return ""
                return str(row[idx]).strip()
            name = cell(i_name)
            card = cell(i_card)
            mob = cell(i_mob)
            if not name and not card:
                continue
            out.append({"name": name, "idCard": card, "mobile": mob})
        if out:
            break
    return out


# ---------------- 不动产 导入（申请人）----------------
def read_re_persons_from_excel(path):
    """读取 姓名/身份证号 两列，返回 [{'name','cardNum'}]（用于① 申请提交批量导入）。"""
    wb = load_workbook(path, data_only=True)
    out = []
    seen = set()
    name_keys = ("姓名", "name", "名字", "申请人姓名", "申请人")
    card_keys = ("身份证号", "身份证", "idcard", "cardnum", "证件号",
                 "申请人身份证号", "身份证号码")
    for ws in wb.worksheets:
        if ws.max_row == 0:
            continue
        header = [str(c.value).strip().lower() if c.value is not None else "" for c in ws[1]]

        def find(keys):
            for i, h in enumerate(header):
                if h in keys:
                    return i
            return None
        i_name = find(name_keys)
        i_card = find(card_keys)
        if i_name is None and i_card is None:
            i_name, i_card = 0, 1
            start = 1
        else:
            start = 2
        for row in ws.iter_rows(min_row=start, values_only=True):
            if not row:
                continue

            def cell(idx):
                if idx is None or idx >= len(row) or row[idx] is None:
                    return ""
                return str(row[idx]).strip()
            name = cell(i_name)
            card = cell(i_card)
            if not name or not card:
                continue
            key = (name, card)
            if key in seen:
                continue
            seen.add(key)
            out.append({"name": name, "cardNum": card})
        if out:
            break
    return out


# ---------------- 不动产 导出（申请记录） ----------------
def _re_extract_records(resp):
    """从 GET /search/assets/ 返回里取出申请记录列表 data[]。"""
    if not isinstance(resp, dict):
        return []
    data = resp.get("data")
    if isinstance(data, list):
        return [r for r in data if isinstance(r, dict)]
    # 兼容 PagedDto 嵌套：{"data":{"data":[...]}}
    if isinstance(data, dict) and isinstance(data.get("data"), list):
        return [r for r in data["data"] if isinstance(r, dict)]
    return []


def export_realestate_records_to_excel(records, output_path):
    """records: GET /search/assets/ 返回的申请记录 data[] 列表。
    生成两张表：申请记录 + 资产明细。"""
    if not isinstance(records, list):
        records = []
    wb = Workbook()
    ws = wb.active
    ws.title = "申请记录"
    ws.append([lbl for _, lbl in RE_RECORD_FIELDS] + ["资产条数"])
    _style_header_row(ws)
    asset_rows = []
    for rec in records:
        if not isinstance(rec, dict):
            continue
        assets = rec.get("assets") or []
        if not isinstance(assets, list):
            assets = []
        ws.append([rec.get(k, "") for k, _ in RE_RECORD_FIELDS] + [len(assets)])
        for a in assets:
            if not isinstance(a, dict):
                continue
            ar = dict(a)
            ar["_req"] = rec.get("reqOrderNo", "")
            ar["_card"] = rec.get("cardNum", "")
            asset_rows.append(ar)
    _autosize_columns(ws)

    if asset_rows:
        ws2 = wb.create_sheet("资产明细")
        ws2.append(["请求单号", "身份证号"] +
                   [lbl for _, lbl in RE_ASSET_FIELDS] + ["资产ID"])
        _style_header_row(ws2)
        for r in asset_rows:
            ws2.append([r.get("_req", ""), r.get("_card", "")] +
                       [r.get(k, "") for k, _ in RE_ASSET_FIELDS] +
                       [r.get("assetId", "") or r.get("id", "")])
        _autosize_columns(ws2)
    wb.save(output_path)


def generate_re_upload_template(output_path):
    """生成 Excel补录模板：前3列(Name/CardNum/ReqOrderNo)必填，须与申请记录一致；
    其余为资产结果列（best-effort，具体补录列以数金实际校验为准）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "补录模板"
    ws.append(list(RE_UPLOAD_TEMPLATE_COLS))
    _style_header_row(ws)
    _autosize_columns(ws)
    wb.save(output_path)


# ---------------- 探知数据 导入/导出 ----------------
def generate_tanzhi_template(output_path, combo=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "导入"
    if combo:
        # 串联产品：车牌自动从名下车辆取，只需姓名/身份证号
        ws.append(["姓名", "身份证号"])
        ws.append(["张三", "11010119991010XXXX"])
    else:
        ws.append(["姓名", "身份证号", "手机号"])
        ws.append(["张三", "11010119991010XXXX", "13800000000"])
        ws.column_dimensions["C"].width = 16
    for c in ws[1]:
        c.font = Font(bold=True)
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 24
    wb.save(output_path)


def generate_tanzhi_plate_template(output_path, need_vin=False):
    """车牌类产品的导入模板：车牌号[，VIN]。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "导入"
    if need_vin:
        ws.append(["车牌号", "VIN"])
        ws.append(["沪ADG8226", "L1NSPGHB5LA004492"])
        ws.column_dimensions["B"].width = 24
    else:
        ws.append(["车牌号"])
        ws.append(["沪ADG8226"])
    for c in ws[1]:
        c.font = Font(bold=True)
    ws.column_dimensions["A"].width = 16
    wb.save(output_path)


def read_tanzhi_plate_rows(path):
    """读取 车牌号[/VIN] 两列，返回 [{'name':VIN,'idCard':车牌,'mobile':''}]
    ——刻意复用批量查询消费方（_tz_run_batch）已有的 name/idCard/mobile 字段形状。"""
    wb = load_workbook(path, data_only=True)
    out = []
    plate_keys = ("车牌号", "车牌", "plateno", "platenum")
    vin_keys = ("vin", "车架号")
    for ws in wb.worksheets:
        if ws.max_row == 0:
            continue
        header = [str(c.value).strip().lower() if c.value is not None else "" for c in ws[1]]

        def find(keys):
            for i, h in enumerate(header):
                if h in keys:
                    return i
            return None
        i_plate = find(plate_keys)
        i_vin = find(vin_keys)
        if i_plate is None:
            i_plate, i_vin = 0, 1
            start = 1
        else:
            start = 2
        for row in ws.iter_rows(min_row=start, values_only=True):
            if not row:
                continue

            def cell(idx):
                if idx is None or idx >= len(row) or row[idx] is None:
                    return ""
                return str(row[idx]).strip()
            plate = cell(i_plate)
            vin = cell(i_vin)
            if not plate:
                continue
            out.append({"name": vin, "idCard": plate, "mobile": ""})
        if out:
            break
    return out


def export_tanzhi_single(info, resp, field_dict, output_path):
    """单条查询：纵向表 字段|中文名|值|说明。"""
    product = TANZHI_PRODUCT_BY_NAME.get(info.get("product", ""), {})
    lookup = tanzhi_lookup(field_dict, product)
    wb = Workbook()
    ws = wb.active
    ws.title = "查询结果"
    ws.append(["产品", info.get("product", "")])
    ws.append(["姓名", info.get("name", "")])
    ws.append(["身份证号", info.get("idnum", "")])
    ws.append(["手机号", info.get("mobile", "")])
    ws.append(["code", info.get("code", "")])
    ws.append(["msg", info.get("msg", "")])
    ws.append(["token", info.get("token", "")])
    ws.append([])
    hdr = ws.max_row + 1
    ws.append(["字段", "中文名", "值", "说明"])
    for c in ws[hdr]:
        c.font = Font(bold=True)
    data = tanzhi_extract_data(resp)
    if data is not None:
        last_idx = None
        for key, val in tanzhi_flatten(data):
            idx = _first_array_index(key)
            if idx is not None and last_idx is not None and idx != last_idx:
                ws.append([])   # 换到下一辆车/下一条列表项时空一行分隔
            last_idx = idx
            ann = tanzhi_annotate(lookup, key) or {}
            desc = ann.get("desc", "")
            note = ann.get("note", "")
            code_note = tanzhi_code_note(key, val)
            expl = "；".join(x for x in (code_note, desc, note) if x)
            ws.append([key, ann.get("cn", ""), val, expl])
    _autosize_columns(ws)
    _align_left_all(ws)
    wb.save(output_path)


def export_tanzhi_batch(results, field_dict, output_path):
    """批量结果按产品分sheet导出；每个sheet前两行为(字段码 / 中文名注释)双表头。
    results: [(info, resp)]"""
    wb = Workbook()
    wb.remove(wb.active)
    # 按产品分组（保持出现顺序）
    groups = {}
    order = []
    for info, resp in results:
        p = info.get("product", "其他")
        if p not in groups:
            groups[p] = []
            order.append(p)
        groups[p].append((info, resp))
    for pname in order:
        product = TANZHI_PRODUCT_BY_NAME.get(pname, {})
        lookup = tanzhi_lookup(field_dict, product)
        rows = groups[pname]
        # 收集所有字段码（去下标，保持首次出现顺序）
        field_codes = []
        seen = set()
        flat_cache = []
        for info, resp in rows:
            data = tanzhi_extract_data(resp)
            pairs = tanzhi_flatten(data) if data is not None else []
            flat_cache.append(pairs)
            for key, _ in pairs:
                code = re.sub(r"\[\d+\]", "", key)
                if code not in seen:
                    seen.add(code)
                    field_codes.append(code)
        title = (pname or "结果")[:28]
        ws = wb.create_sheet(title=title)
        base = ["姓名", "身份证号", "手机号", "code", "msg", "token"]
        ws.append(base + field_codes)
        ann_row = [""] * len(base) + [
            (tanzhi_annotate(lookup, c) or {}).get("cn", "") for c in field_codes]
        ws.append(ann_row)
        for c in ws[1]:
            c.font = Font(bold=True)
        for c in ws[2]:
            c.font = Font(italic=True, color="808080")
        for (info, resp), pairs in zip(rows, flat_cache):
            valmap = {}
            for key, val in pairs:
                code = re.sub(r"\[\d+\]", "", key)
                # 同名(列表展开)用分号合并
                if code in valmap and valmap[code] != "":
                    valmap[code] = "%s；%s" % (valmap[code], val)
                else:
                    valmap[code] = val
            line = [info.get("name", ""), info.get("idnum", ""), info.get("mobile", ""),
                    info.get("code", ""), info.get("msg", ""), info.get("token", "")]
            line += [valmap.get(c, "") for c in field_codes]
            ws.append(line)
        ws.freeze_panes = "G3"
        # 状态码图例：批量表没有逐格"说明"列，在数据下方补一段中文对照，避免用户翻文档
        ws.append([])
        ws.append(["状态码说明（适用于本表所有 code / xxx_code 列）"])
        for code, meaning in TANZHI_STATUS_CODES.items():
            ws.append([code, meaning])
        _align_left_all(ws)
    if not wb.sheetnames:
        wb.create_sheet("结果")
    wb.save(output_path)


# ---------------- GUI ----------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("数金数据查询工具")
        root.geometry("1040x740")

        self.cfg = load_config()

        # 各模块结果
        self.inv_results = []           # [(info, resp)]
        self.gjj_results = []           # [(info, resp)]
        self.re_persons = []            # 申请提交·申请人列表 [{name,cardNum}]
        self.re_records = []            # 最近一次查询到的申请记录列表
        self.tz_results = []            # 探知数据批量结果 [(info, resp)]
        self.tz_last = None             # 探知数据·单个查询最近结果 (info, resp)
        self.tanzhi_dict = load_tanzhi_dict()   # 字段标注字典

        self.inv_cancel = threading.Event()
        self.gjj_cancel = threading.Event()
        self.tz_cancel = threading.Event()

        self._build_ui()

    def _build_ui(self):
        # 顶部 Authorization
        top = ttk.LabelFrame(
            self.root,
            text="Authorization（所有接口共用，自动补 Bearer 前缀）",
        )
        top.pack(fill="x", padx=10, pady=8)
        self.auth_var = tk.StringVar(value=self.cfg.get("authorization", ""))
        self.auth_entry = ttk.Entry(top, textvariable=self.auth_var, show="*")
        self.auth_entry.pack(side="left", fill="x", expand=True, padx=(8, 4), pady=6)
        self.show_auth_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="显示", variable=self.show_auth_var,
                        command=self._toggle_auth_show).pack(side="left", padx=4)
        ttk.Button(top, text="保存", command=self._save_auth).pack(side="left", padx=(2, 8))

        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True, padx=10, pady=4)
        self._build_invoice_module(nb)
        self._build_realestate_module(nb)
        self._build_gjj_module(nb)
        self._build_tanzhi_module(nb)

        self.status_var = tk.StringVar(value="就绪")
        bar = ttk.Frame(self.root)
        bar.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Label(bar, textvariable=self.status_var, anchor="w").pack(
            side="left", fill="x", expand=True)

    # ----- 公共 -----
    def _make_tree(self, parent, columns, height=10, horizontal=False):
        """在一个容器框里创建带垂直（可选水平）滚动条的 Treeview。
        返回 (wrap_frame, tree)；调用方负责 pack/grid wrap_frame。"""
        wrap = ttk.Frame(parent)
        tree = ttk.Treeview(wrap, columns=columns, show="headings", height=height)
        vs = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vs.set)
        if horizontal:
            hs = ttk.Scrollbar(wrap, orient="horizontal", command=tree.xview)
            tree.configure(xscrollcommand=hs.set)
            hs.pack(side="bottom", fill="x")
        vs.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)
        return wrap, tree

    def _attach_tree_tooltip(self, tree, wraplength=600):
        """鼠标悬停在 Treeview 单元格上时，浮窗显示该单元格的完整文本。"""
        state = {"win": None, "key": None}

        def hide(_e=None):
            if state["win"] is not None:
                try:
                    state["win"].destroy()
                except Exception:
                    pass
                state["win"] = None
                state["key"] = None

        def on_motion(event):
            row = tree.identify_row(event.y)
            col = tree.identify_column(event.x)
            if not row or not col:
                hide()
                return
            try:
                idx = int(col[1:]) - 1
                vals = tree.item(row, "values")
                text = str(vals[idx]) if 0 <= idx < len(vals) else ""
            except Exception:
                text = ""
            if not text:
                hide()
                return
            key = (row, col)
            if key == state["key"] and state["win"] is not None:
                return
            hide()
            state["key"] = key
            win = tk.Toplevel(tree)
            win.wm_overrideredirect(True)
            try:
                win.attributes("-topmost", True)
            except Exception:
                pass
            tk.Label(win, text=text, justify="left", background="#ffffe0",
                     relief="solid", borderwidth=1, wraplength=wraplength,
                     font=("", 9)).pack(ipadx=3, ipady=2)
            win.wm_geometry("+%d+%d" % (event.x_root + 14, event.y_root + 16))
            state["win"] = win

        tree.bind("<Motion>", on_motion)
        tree.bind("<Leave>", hide)

    def _toggle_auth_show(self):
        self.auth_entry.config(show="" if self.show_auth_var.get() else "*")

    def _save_auth(self):
        self.cfg["authorization"] = self.auth_var.get().strip()
        save_config(self.cfg)
        self.status_var.set("Authorization 已保存")

    def _auth_or_warn(self):
        auth = self.auth_var.get().strip()
        if not auth:
            messagebox.showwarning("提示", "请先填写 Authorization")
            return None
        return auth

    def _set_status(self, s):
        self.root.after(0, self.status_var.set, s)

    def _ask_save_path(self, initialfile):
        return filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            initialfile=initialfile)

    # =========================================================
    #  模块一：发票查询
    # =========================================================
    def _build_invoice_module(self, parent_nb):
        outer = ttk.Frame(parent_nb)
        parent_nb.add(outer, text="发票查询")
        nb = ttk.Notebook(outer)
        nb.pack(fill="both", expand=True)
        self._build_inv_single(nb)
        self._build_inv_batch(nb)

    def _build_inv_single(self, nb):
        frm = ttk.Frame(nb)
        nb.add(frm, text="单个查询")
        bar = ttk.Frame(frm)
        bar.pack(fill="x", padx=8, pady=6)
        ttk.Label(bar, text="纳税人识别号：").pack(side="left")
        self.inv_sh_var = tk.StringVar()
        ttk.Entry(bar, textvariable=self.inv_sh_var, width=30).pack(side="left", padx=4)
        self.inv_single_btn = ttk.Button(bar, text="查询", command=self._inv_run_single)
        self.inv_single_btn.pack(side="left", padx=4)
        ttk.Button(bar, text="导出本次结果到Excel",
                   command=self._inv_export_single).pack(side="left", padx=4)
        self.inv_output = scrolledtext.ScrolledText(frm, wrap="word", font=("Consolas", 10))
        self.inv_output.pack(fill="both", expand=True, padx=8, pady=4)
        self._inv_last = None

    def _build_inv_batch(self, nb):
        frm = ttk.Frame(nb)
        nb.add(frm, text="批量查询（Excel导入）")
        top = ttk.Frame(frm)
        top.pack(fill="x", padx=8, pady=6)
        ttk.Label(top, text="Excel文件：").pack(side="left")
        self.inv_file_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.inv_file_var).pack(
            side="left", fill="x", expand=True, padx=4)
        ttk.Button(top, text="选择...", command=self._inv_choose_excel).pack(side="left", padx=2)
        ttk.Button(top, text="生成导入模板", command=self._inv_gen_template).pack(side="left", padx=2)

        bar = ttk.Frame(frm)
        bar.pack(fill="x", padx=8, pady=2)
        self.inv_batch_btn = ttk.Button(bar, text="开始批量查询", command=self._inv_run_batch)
        self.inv_batch_btn.pack(side="left", padx=2)
        self.inv_cancel_btn = ttk.Button(bar, text="取消", command=self.inv_cancel.set,
                                         state="disabled")
        self.inv_cancel_btn.pack(side="left", padx=2)
        ttk.Button(bar, text="导出全部结果到Excel", command=self._inv_export_batch).pack(side="left", padx=2)
        ttk.Button(bar, text="清空结果", command=self._inv_clear).pack(side="left", padx=2)

        prog = ttk.Frame(frm)
        prog.pack(fill="x", padx=8, pady=4)
        self.inv_progress = ttk.Progressbar(prog, mode="determinate")
        self.inv_progress.pack(fill="x", expand=True, side="left")
        self.inv_progress_label = ttk.Label(prog, text="0/0", width=10)
        self.inv_progress_label.pack(side="left", padx=4)

        cols = ("sh", "status", "code", "message", "id")
        names = ("纳税人识别号", "状态", "Code", "消息", "查询ID")
        wrap, self.inv_tree = self._make_tree(frm, cols, height=10, horizontal=True)
        for c, n in zip(cols, names):
            self.inv_tree.heading(c, text=n)
            self.inv_tree.column(c, width=140 if c != "message" else 320, anchor="w")
        wrap.pack(fill="both", expand=True, padx=8, pady=4)

    def _inv_choose_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx;*.xlsm"), ("All", "*.*")])
        if p:
            self.inv_file_var.set(p)

    def _inv_gen_template(self):
        path = self._ask_save_path("纳税人识别号_导入模板.xlsx")
        if not path:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "导入"
        ws.append(["纳税人识别号"])
        ws["A1"].font = Font(bold=True)
        ws.append(["91350802MA34XXBL8G"])
        ws.column_dimensions["A"].width = 28
        wb.save(path)
        messagebox.showinfo("完成", f"模板已生成:\n{path}")

    def _inv_run_single(self):
        sh = self.inv_sh_var.get().strip()
        if not sh:
            messagebox.showwarning("提示", "请输入纳税人识别号")
            return
        auth = self._auth_or_warn()
        if not auth:
            return
        client = InvoiceClient(auth)
        self.inv_single_btn.config(state="disabled")
        self.inv_output.delete("1.0", "end")
        self.inv_output.insert("end", f"正在查询 {sh} ...\n")
        self.status_var.set("查询中...")

        def task():
            try:
                ok, info, resp = client.query_one(sh, on_status=self._set_status)
                self.root.after(0, self._inv_on_single_done, ok, info, resp)
            except Exception as e:
                self.root.after(0, self._on_error, e, traceback.format_exc(),
                                self.inv_single_btn)
        threading.Thread(target=task, daemon=True).start()

    def _inv_on_single_done(self, ok, info, resp):
        self.inv_single_btn.config(state="normal")
        self.status_var.set("完成" if ok else f"失败：{info.get('message','')}")
        self._inv_last = (info, resp)
        self.inv_results.append((info, resp))
        self._inv_append_row(info)
        self.inv_output.delete("1.0", "end")
        self.inv_output.insert(
            "end",
            f"纳税人识别号: {info.get('sh')}\n状态: {info.get('status')}\n"
            f"Code: {info.get('code')}\n消息: {info.get('message')}\n"
            f"查询ID: {info.get('id')}\n{'-'*60}\n")
        if resp:
            try:
                self.inv_output.insert("end", json.dumps(resp, ensure_ascii=False, indent=2))
            except Exception:
                self.inv_output.insert("end", str(resp))

    def _inv_export_single(self):
        if not self._inv_last:
            messagebox.showinfo("提示", "请先执行一次单个查询")
            return
        info, resp = self._inv_last
        path = self._ask_save_path(
            f"发票查询_{info.get('sh','')}_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        if not path:
            return
        try:
            export_results_to_excel([(info, resp)], path)
            self.status_var.set(f"已导出: {path}")
            messagebox.showinfo("完成", f"已导出到:\n{path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def _inv_run_batch(self):
        path = self.inv_file_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请选择有效的Excel文件")
            return
        auth = self._auth_or_warn()
        if not auth:
            return
        client = InvoiceClient(auth)
        try:
            sh_list = read_taxnos_from_excel(path)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return
        if not sh_list:
            messagebox.showinfo("提示", "未在Excel中读取到纳税人识别号")
            return
        if not messagebox.askyesno("确认", f"共读取到 {len(sh_list)} 个纳税人识别号，是否开始查询？"):
            return
        self.inv_cancel.clear()
        self.inv_batch_btn.config(state="disabled")
        self.inv_cancel_btn.config(state="normal")
        self.inv_progress["maximum"] = len(sh_list)
        self.inv_progress["value"] = 0
        self.inv_progress_label.config(text=f"0/{len(sh_list)}")

        def task():
            done = 0
            for sh in sh_list:
                if self.inv_cancel.is_set():
                    break
                try:
                    ok, info, resp = client.query_one(sh, on_status=self._set_status)
                except Exception as e:
                    info = {"sh": sh, "status": "异常", "code": "", "message": str(e), "id": ""}
                    resp = None
                done += 1
                self.root.after(0, self._inv_on_batch_one, info, resp, done, len(sh_list))
            self.root.after(0, self._inv_on_batch_done)
        threading.Thread(target=task, daemon=True).start()

    def _inv_on_batch_one(self, info, resp, done, total):
        self.inv_results.append((info, resp))
        self._inv_append_row(info)
        self.inv_progress["value"] = done
        self.inv_progress_label.config(text=f"{done}/{total}")

    def _inv_on_batch_done(self):
        self.inv_batch_btn.config(state="normal")
        self.inv_cancel_btn.config(state="disabled")
        self.status_var.set("批量查询结束")
        messagebox.showinfo("完成", "批量查询已结束，可点击导出按钮保存结果")

    def _inv_append_row(self, info):
        self.inv_tree.insert("", "end", values=(
            info.get("sh", ""), info.get("status", ""), info.get("code", ""),
            info.get("message", ""), info.get("id", "")))

    def _inv_export_batch(self):
        if not self.inv_results:
            messagebox.showinfo("提示", "暂无结果可导出")
            return
        path = self._ask_save_path(f"发票查询结果_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        if not path:
            return
        try:
            export_results_to_excel(self.inv_results, path)
            self.status_var.set(f"已导出: {path}")
            messagebox.showinfo("完成", f"已导出到:\n{path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def _inv_clear(self):
        if not self.inv_results:
            return
        if not messagebox.askyesno("确认", "确认清空发票查询结果？"):
            return
        self.inv_results.clear()
        for i in self.inv_tree.get_children():
            self.inv_tree.delete(i)
        self.inv_progress["value"] = 0
        self.inv_progress_label.config(text="0/0")
        self.status_var.set("已清空")

    def _on_error(self, e, tb, btn=None):
        if btn is not None:
            btn.config(state="normal")
        self.status_var.set("出错")
        messagebox.showerror("错误", f"{e}\n\n{tb}")

    # =========================================================
    #  模块二：公积金收入评分
    # =========================================================
    def _build_gjj_module(self, parent_nb):
        outer = ttk.Frame(parent_nb)
        parent_nb.add(outer, text="公积金收入评分")
        nb = ttk.Notebook(outer)
        nb.pack(fill="both", expand=True)
        self._build_gjj_single(nb)
        self._build_gjj_batch(nb)
        self._build_gjj_history(nb)

    def _build_gjj_single(self, nb):
        frm = ttk.Frame(nb)
        nb.add(frm, text="单个查询")
        bar = ttk.Frame(frm)
        bar.pack(fill="x", padx=8, pady=6)
        ttk.Label(bar, text="姓名：").pack(side="left")
        self.gjj_name_var = tk.StringVar()
        ttk.Entry(bar, textvariable=self.gjj_name_var, width=12).pack(side="left", padx=4)
        ttk.Label(bar, text="身份证号：").pack(side="left")
        self.gjj_card_var = tk.StringVar()
        ttk.Entry(bar, textvariable=self.gjj_card_var, width=24).pack(side="left", padx=4)
        ttk.Label(bar, text="手机号(可选)：").pack(side="left")
        self.gjj_mobile_var = tk.StringVar()
        ttk.Entry(bar, textvariable=self.gjj_mobile_var, width=15).pack(side="left", padx=4)
        self.gjj_single_btn = ttk.Button(bar, text="查询", command=self._gjj_run_single)
        self.gjj_single_btn.pack(side="left", padx=4)
        ttk.Button(bar, text="导出本次结果到Excel",
                   command=self._gjj_export_single).pack(side="left", padx=4)
        self.gjj_output = scrolledtext.ScrolledText(frm, wrap="word", font=("Consolas", 10))
        self.gjj_output.pack(fill="both", expand=True, padx=8, pady=4)
        self._gjj_last = None

    def _build_gjj_batch(self, nb):
        frm = ttk.Frame(nb)
        nb.add(frm, text="批量查询（Excel导入）")
        top = ttk.Frame(frm)
        top.pack(fill="x", padx=8, pady=6)
        ttk.Label(top, text="Excel文件：").pack(side="left")
        self.gjj_file_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.gjj_file_var).pack(
            side="left", fill="x", expand=True, padx=4)
        ttk.Button(top, text="选择...", command=self._gjj_choose_excel).pack(side="left", padx=2)
        ttk.Button(top, text="生成导入模板", command=self._gjj_gen_template).pack(side="left", padx=2)

        bar = ttk.Frame(frm)
        bar.pack(fill="x", padx=8, pady=2)
        self.gjj_batch_btn = ttk.Button(bar, text="开始批量查询", command=self._gjj_run_batch)
        self.gjj_batch_btn.pack(side="left", padx=2)
        self.gjj_cancel_btn = ttk.Button(bar, text="取消", command=self.gjj_cancel.set,
                                         state="disabled")
        self.gjj_cancel_btn.pack(side="left", padx=2)
        ttk.Button(bar, text="导出全部结果到Excel", command=self._gjj_export_batch).pack(side="left", padx=2)
        ttk.Button(bar, text="清空结果", command=self._gjj_clear).pack(side="left", padx=2)

        prog = ttk.Frame(frm)
        prog.pack(fill="x", padx=8, pady=4)
        self.gjj_progress = ttk.Progressbar(prog, mode="determinate")
        self.gjj_progress.pack(fill="x", expand=True, side="left")
        self.gjj_progress_label = ttk.Label(prog, text="0/0", width=10)
        self.gjj_progress_label.pack(side="left", padx=4)

        cols = ("idCard", "name", "status", "score", "scoreRange", "message")
        names = ("身份证号", "姓名", "状态", "评分", "评分区间", "消息")
        wrap, self.gjj_tree = self._make_tree(frm, cols, height=10, horizontal=True)
        for c, n in zip(cols, names):
            self.gjj_tree.heading(c, text=n)
            self.gjj_tree.column(c, width=150 if c in ("idCard", "message") else 110, anchor="w")
        wrap.pack(fill="both", expand=True, padx=8, pady=4)

    def _build_gjj_history(self, nb):
        frm = ttk.Frame(nb)
        nb.add(frm, text="历史记录")
        bar = ttk.Frame(frm)
        bar.pack(fill="x", padx=8, pady=6)
        ttk.Label(bar, text="页码：").pack(side="left")
        self.gjj_hist_page_var = tk.StringVar(value="1")
        ttk.Entry(bar, textvariable=self.gjj_hist_page_var, width=6).pack(side="left", padx=4)
        ttk.Label(bar, text="每页：").pack(side="left")
        self.gjj_hist_size_var = tk.StringVar(value="10")
        ttk.Entry(bar, textvariable=self.gjj_hist_size_var, width=6).pack(side="left", padx=4)
        ttk.Button(bar, text="查询历史", command=self._gjj_run_history).pack(side="left", padx=4)
        self.gjj_hist_info = ttk.Label(bar, text="")
        self.gjj_hist_info.pack(side="left", padx=8)

        cols = ("idCard", "name", "status", "score", "scoreRange", "jfztText", "createdDate")
        names = ("身份证号", "姓名", "状态", "评分", "评分区间", "缴存状态", "创建时间")
        wrap, self.gjj_hist_tree = self._make_tree(frm, cols, height=14, horizontal=True)
        for c, n in zip(cols, names):
            self.gjj_hist_tree.heading(c, text=n)
            self.gjj_hist_tree.column(c, width=150 if c in ("idCard", "createdDate") else 100, anchor="w")
        wrap.pack(fill="both", expand=True, padx=8, pady=4)

    def _gjj_choose_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx;*.xlsm"), ("All", "*.*")])
        if p:
            self.gjj_file_var.set(p)

    def _gjj_gen_template(self):
        path = self._ask_save_path("公积金评分_导入模板.xlsx")
        if not path:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "导入"
        ws.append(["姓名", "身份证号", "手机号"])
        for c in ws[1]:
            c.font = Font(bold=True)
        ws.append(["张三", "210202199001018888", "13800138000"])
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 24
        ws.column_dimensions["C"].width = 15
        wb.save(path)
        messagebox.showinfo("完成", f"模板已生成:\n{path}")

    def _gjj_run_single(self):
        name = self.gjj_name_var.get().strip()
        card = self.gjj_card_var.get().strip()
        mobile = self.gjj_mobile_var.get().strip()
        if not name or not card:
            messagebox.showwarning("提示", "请输入姓名和身份证号")
            return
        auth = self._auth_or_warn()
        if not auth:
            return
        client = GjjClient(auth)
        self.gjj_single_btn.config(state="disabled")
        self.gjj_output.delete("1.0", "end")
        self.gjj_output.insert("end", f"正在查询 {name} ...\n")
        self.status_var.set("查询中...")

        def task():
            try:
                ok, info, resp = client.query_one(name, card, mobile, on_status=self._set_status)
                self.root.after(0, self._gjj_on_single_done, ok, info, resp)
            except Exception as e:
                self.root.after(0, self._on_error, e, traceback.format_exc(),
                                self.gjj_single_btn)
        threading.Thread(target=task, daemon=True).start()

    def _gjj_on_single_done(self, ok, info, resp):
        self.gjj_single_btn.config(state="normal")
        self.status_var.set("完成" if ok else f"失败：{info.get('message','')}")
        self._gjj_last = (info, resp)
        self.gjj_results.append((info, resp))
        self._gjj_append_row(info, resp)
        d = (resp.get("data") or {}) if resp else {}
        self.gjj_output.delete("1.0", "end")
        self.gjj_output.insert(
            "end",
            f"姓名: {info.get('name')}\n身份证号: {info.get('idCard')}\n"
            f"状态: {info.get('status')}\n"
            f"评分: {d.get('score','') if isinstance(d, dict) else ''}  "
            f"区间: {d.get('scoreRange','') if isinstance(d, dict) else ''}\n"
            f"消息: {info.get('message')}\n受理ID: {info.get('id')}\n{'-'*60}\n")
        if resp:
            try:
                self.gjj_output.insert("end", json.dumps(resp, ensure_ascii=False, indent=2))
            except Exception:
                self.gjj_output.insert("end", str(resp))

    def _gjj_export_single(self):
        if not self._gjj_last:
            messagebox.showinfo("提示", "请先执行一次单个查询")
            return
        info, resp = self._gjj_last
        path = self._ask_save_path(
            f"公积金评分_{info.get('idCard','')}_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        if not path:
            return
        try:
            export_gjj_to_excel([(info, resp)], path)
            self.status_var.set(f"已导出: {path}")
            messagebox.showinfo("完成", f"已导出到:\n{path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def _gjj_run_batch(self):
        path = self.gjj_file_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请选择有效的Excel文件")
            return
        auth = self._auth_or_warn()
        if not auth:
            return
        client = GjjClient(auth)
        try:
            persons = read_persons_from_excel(path)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return
        if not persons:
            messagebox.showinfo("提示", "未在Excel中读取到姓名/身份证号")
            return
        if not messagebox.askyesno("确认", f"共读取到 {len(persons)} 条记录，是否开始查询？"):
            return
        self.gjj_cancel.clear()
        self.gjj_batch_btn.config(state="disabled")
        self.gjj_cancel_btn.config(state="normal")
        self.gjj_progress["maximum"] = len(persons)
        self.gjj_progress["value"] = 0
        self.gjj_progress_label.config(text=f"0/{len(persons)}")

        def task():
            done = 0
            for p in persons:
                if self.gjj_cancel.is_set():
                    break
                try:
                    ok, info, resp = client.query_one(
                        p["name"], p["idCard"], p.get("mobile", ""), on_status=self._set_status)
                except Exception as e:
                    info = {"name": p.get("name", ""), "idCard": p.get("idCard", ""),
                            "mobile": p.get("mobile", ""), "status": "异常",
                            "code": "", "message": str(e), "id": ""}
                    resp = None
                done += 1
                self.root.after(0, self._gjj_on_batch_one, info, resp, done, len(persons))
            self.root.after(0, self._gjj_on_batch_done)
        threading.Thread(target=task, daemon=True).start()

    def _gjj_on_batch_one(self, info, resp, done, total):
        self.gjj_results.append((info, resp))
        self._gjj_append_row(info, resp)
        self.gjj_progress["value"] = done
        self.gjj_progress_label.config(text=f"{done}/{total}")

    def _gjj_on_batch_done(self):
        self.gjj_batch_btn.config(state="normal")
        self.gjj_cancel_btn.config(state="disabled")
        self.status_var.set("批量查询结束")
        messagebox.showinfo("完成", "批量查询已结束，可点击导出按钮保存结果")

    def _gjj_append_row(self, info, resp):
        d = (resp.get("data") or {}) if resp else {}
        if not isinstance(d, dict):
            d = {}
        self.gjj_tree.insert("", "end", values=(
            info.get("idCard", ""), info.get("name", ""), info.get("status", ""),
            d.get("score", ""), d.get("scoreRange", ""), info.get("message", "")))

    def _gjj_export_batch(self):
        if not self.gjj_results:
            messagebox.showinfo("提示", "暂无结果可导出")
            return
        path = self._ask_save_path(f"公积金评分结果_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        if not path:
            return
        try:
            export_gjj_to_excel(self.gjj_results, path)
            self.status_var.set(f"已导出: {path}")
            messagebox.showinfo("完成", f"已导出到:\n{path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def _gjj_clear(self):
        if not self.gjj_results:
            return
        if not messagebox.askyesno("确认", "确认清空公积金查询结果？"):
            return
        self.gjj_results.clear()
        for i in self.gjj_tree.get_children():
            self.gjj_tree.delete(i)
        self.gjj_progress["value"] = 0
        self.gjj_progress_label.config(text="0/0")
        self.status_var.set("已清空")

    def _gjj_run_history(self):
        auth = self._auth_or_warn()
        if not auth:
            return
        try:
            page = int(self.gjj_hist_page_var.get().strip() or "1")
            size = int(self.gjj_hist_size_var.get().strip() or "10")
        except ValueError:
            messagebox.showwarning("提示", "页码/每页必须为数字")
            return
        client = GjjClient(auth)
        self.status_var.set("查询历史中...")

        def task():
            try:
                resp = client.history(page, size)
                self.root.after(0, self._gjj_on_history, resp)
            except Exception as e:
                self.root.after(0, self._on_error, e, traceback.format_exc())
        threading.Thread(target=task, daemon=True).start()

    def _gjj_on_history(self, resp):
        for i in self.gjj_hist_tree.get_children():
            self.gjj_hist_tree.delete(i)
        data = resp.get("data") or {}
        if not isinstance(data, dict):
            self.status_var.set("历史返回异常")
            messagebox.showwarning("提示", str(resp.get("message") or resp))
            return
        rows = data.get("data") or []
        self.gjj_hist_info.config(
            text=f"共 {data.get('totalCount', 0)} 条 / {data.get('totalPage', 0)} 页 "
                 f"(当前第 {data.get('currentPage', 0)} 页)")
        for r in rows:
            if not isinstance(r, dict):
                continue
            self.gjj_hist_tree.insert("", "end", values=(
                r.get("idCard", ""), r.get("name", ""), r.get("status", ""),
                r.get("score", ""), r.get("scoreRange", ""),
                r.get("jfztText", ""), r.get("createdDate", "")))
        self.status_var.set("历史查询完成")

    # =========================================================
    #  模块三：不动产查询（申请人直接写库链路）
    # =========================================================
    def _build_realestate_module(self, parent_nb):
        outer = ttk.Frame(parent_nb)
        parent_nb.add(outer, text="不动产查询")

        # 顶部开关：补录/编辑属于“管理端写入”功能，纯查询用不到，默认隐藏
        adminbar = ttk.Frame(outer)
        adminbar.pack(fill="x", padx=4, pady=(4, 0))
        self.re_show_admin_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            adminbar,
            text="显示“③ Excel补录 / ④ 编辑资产 / ⑤ 处理回调”（管理端写入功能，纯查询无需勾选）",
            variable=self.re_show_admin_var, command=self._re_toggle_admin,
        ).pack(side="left")

        nb = ttk.Notebook(outer)
        nb.pack(fill="both", expand=True)
        self.re_nb = nb
        self._build_re_apply(nb)        # ① 申请提交（直接写库）
        self._build_re_records(nb)      # ② 查询申请记录
        self._build_re_upload(nb)       # ③ Excel补录资产结果（默认隐藏）
        self._build_re_edit(nb)         # ④ 编辑资产结果（默认隐藏）
        self._build_re_callback(nb)     # ⑤ 处理回调
        # 默认隐藏 ③④
        self._re_toggle_admin()

    def _re_toggle_admin(self):
        """根据开关显示/隐藏 ③Excel补录、④编辑资产、⑤处理回调 三个子页（保持 ①②③④⑤ 顺序）。"""
        nb = self.re_nb
        if self.re_show_admin_var.get():
            # 显示：插回到 ② 之后（索引2、3、4）
            try:
                nb.insert(2, self.re_tab_upload, text="③ Excel补录")
                nb.insert(3, self.re_tab_edit, text="④ 编辑资产")
                nb.insert(4, self.re_tab_callback, text="⑤ 处理回调")
            except Exception:
                nb.add(self.re_tab_upload, text="③ Excel补录")
                nb.add(self.re_tab_edit, text="④ 编辑资产")
                nb.add(self.re_tab_callback, text="⑤ 处理回调")
        else:
            for frm in (getattr(self, "re_tab_upload", None),
                        getattr(self, "re_tab_edit", None),
                        getattr(self, "re_tab_callback", None)):
                if frm is not None:
                    try:
                        nb.forget(frm)
                    except Exception:
                        pass

    # ----- ① 申请提交（直接写库） -----
    def _build_re_apply(self, nb):
        frm = ttk.Frame(nb)
        nb.add(frm, text="① 申请提交")

        ttk.Label(
            frm,
            text="提交申请人（仅姓名+身份证号）直接写入申请记录表，服务端内部生成请求单号；"
                 "提交后到“② 查询申请记录”取回 reqOrderNo。"
                 "手动添加或 Excel 导入都会进入下方“待提交列表”，核对无误后点“提交申请”一次性提交。",
            foreground="#666", wraplength=940, justify="left",
        ).pack(fill="x", padx=8, pady=(8, 4))

        # 方式一：手动逐条添加
        man = ttk.LabelFrame(frm, text="方式一：手动添加")
        man.pack(fill="x", padx=8, pady=4)
        pbar = ttk.Frame(man)
        pbar.pack(fill="x", padx=8, pady=6)
        ttk.Label(pbar, text="姓名：").pack(side="left")
        self.re_name_var = tk.StringVar()
        ttk.Entry(pbar, textvariable=self.re_name_var, width=14).pack(side="left", padx=4)
        ttk.Label(pbar, text="身份证号：").pack(side="left")
        self.re_card_var = tk.StringVar()
        ttk.Entry(pbar, textvariable=self.re_card_var, width=24).pack(side="left", padx=4)
        ttk.Button(pbar, text="添加到列表", command=self._re_add_person).pack(side="left", padx=4)
        ttk.Button(pbar, text="移除选中", command=self._re_remove_person).pack(side="left", padx=4)

        # 方式二：Excel 批量导入（仅导入到列表，不直接提交）
        imp = ttk.LabelFrame(frm, text="方式二：Excel批量导入（两列——姓名、身份证号）")
        imp.pack(fill="x", padx=8, pady=4)
        ibar = ttk.Frame(imp)
        ibar.pack(fill="x", padx=8, pady=6)
        ttk.Label(ibar, text="Excel文件：").pack(side="left")
        self.re_apply_file_var = tk.StringVar()
        ttk.Entry(ibar, textvariable=self.re_apply_file_var).pack(
            side="left", fill="x", expand=True, padx=4)
        ttk.Button(ibar, text="选择...", command=self._re_apply_choose_excel).pack(side="left", padx=2)
        ttk.Button(ibar, text="生成导入模板", command=self._re_apply_gen_template).pack(side="left", padx=2)
        ttk.Button(ibar, text="导入到列表", command=self._re_import_to_list).pack(side="left", padx=2)

        # 待提交列表（手动/导入都进这里，提交前可核对）
        ttk.Label(frm, text="待提交申请人（核对无误后点“提交申请”）：",
                  foreground="#666").pack(anchor="w", padx=8, pady=(6, 0))
        treewrap, self.re_person_tree = self._make_tree(
            frm, ("name", "cardNum"), height=6)
        self.re_person_tree.heading("name", text="姓名")
        self.re_person_tree.heading("cardNum", text="身份证号")
        self.re_person_tree.column("name", width=180, anchor="w")
        self.re_person_tree.column("cardNum", width=300, anchor="w")
        treewrap.pack(fill="both", expand=True, padx=8, pady=2)

        # 提交
        sbar = ttk.Frame(frm)
        sbar.pack(fill="x", padx=8, pady=6)
        self.re_apply_btn = ttk.Button(sbar, text="提交申请", command=self._re_run_apply)
        self.re_apply_btn.pack(side="left", padx=2)
        ttk.Button(sbar, text="清空列表", command=self._re_clear_apply).pack(side="left", padx=2)
        self.re_apply_count_var = tk.StringVar(value="待提交：0 人")
        ttk.Label(sbar, textvariable=self.re_apply_count_var,
                  foreground="#666").pack(side="left", padx=8)

        self.re_apply_output = scrolledtext.ScrolledText(
            frm, wrap="word", font=("Consolas", 10), height=7)
        self.re_apply_output.pack(fill="x", padx=8, pady=4)

    def _re_apply_choose_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx;*.xlsm"), ("All", "*.*")])
        if p:
            self.re_apply_file_var.set(p)

    def _re_apply_gen_template(self):
        path = self._ask_save_path("不动产申请人_导入模板.xlsx")
        if not path:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "申请人"
        ws.append(["姓名", "身份证号"])
        _style_header_row(ws)
        ws.append(["张三", "350802199001011234"])
        ws.column_dimensions["A"].width = 16
        ws.column_dimensions["B"].width = 28
        wb.save(path)
        messagebox.showinfo("完成", f"模板已生成:\n{path}")

    def _re_import_to_list(self):
        """仅把 Excel 里的申请人读入“待提交列表”，不直接提交（提交由“提交申请”负责）。"""
        path = self.re_apply_file_var.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择要导入的 Excel 文件")
            return
        try:
            persons = read_re_persons_from_excel(path)
        except Exception as e:
            messagebox.showerror("读取失败", f"无法读取 Excel：\n{e}")
            return
        if not persons:
            messagebox.showwarning("提示", "未从 Excel 读到有效的申请人（需有 姓名 与 身份证号 两列）")
            return
        # 追加到当前列表（与已有项去重），方便先核对再提交
        exist = {(p["name"], p["cardNum"]) for p in self.re_persons}
        added = 0
        for p in persons:
            key = (p["name"], p["cardNum"])
            if key in exist:
                continue
            exist.add(key)
            self.re_persons.append({"name": p["name"], "cardNum": p["cardNum"]})
            self.re_person_tree.insert("", "end", values=(p["name"], p["cardNum"]))
            added += 1
        self._re_update_apply_count()
        skipped = len(persons) - added
        msg = f"已导入 {added} 人到待提交列表"
        if skipped:
            msg += f"（跳过 {skipped} 个重复项）"
        msg += "，请核对后点“提交申请”。"
        self.status_var.set(msg)
        messagebox.showinfo("导入完成", msg)

    def _re_update_apply_count(self):
        n = len(self.re_person_tree.get_children())
        self.re_apply_count_var.set(f"待提交：{n} 人")

    def _re_add_person(self):
        name = self.re_name_var.get().strip()
        card = self.re_card_var.get().strip()
        if not name or not card:
            messagebox.showwarning("提示", "请输入申请人姓名和身份证号")
            return
        self.re_persons.append({"name": name, "cardNum": card})
        self.re_person_tree.insert("", "end", values=(name, card))
        self.re_name_var.set("")
        self.re_card_var.set("")
        self._re_update_apply_count()

    def _re_remove_person(self):
        for item in self.re_person_tree.selection():
            idx = self.re_person_tree.index(item)
            if idx < len(self.re_persons):
                del self.re_persons[idx]
            self.re_person_tree.delete(item)
        self._re_update_apply_count()

    def _re_clear_apply(self):
        self.re_name_var.set("")
        self.re_card_var.set("")
        self.re_persons.clear()
        for i in self.re_person_tree.get_children():
            self.re_person_tree.delete(i)
        self.re_apply_output.delete("1.0", "end")
        self._re_update_apply_count()

    def _re_run_apply(self):
        persons = list(self.re_persons)
        name = self.re_name_var.get().strip()
        card = self.re_card_var.get().strip()
        if name and card:
            persons.append({"name": name, "cardNum": card})
        if not persons:
            messagebox.showwarning("提示", "请至少添加一个申请人（姓名+身份证号）")
            return
        auth = self._auth_or_warn()
        if not auth:
            return
        client = RealEstateClient(auth)
        self.re_apply_btn.config(state="disabled")
        self.re_apply_output.delete("1.0", "end")
        self.re_apply_output.insert("end", f"正在提交 {len(persons)} 位申请人...\n")
        self.status_var.set("提交申请中...")

        def task():
            try:
                resp = client.submit_application(persons)
                self.root.after(0, self._re_on_apply, resp)
            except Exception as e:
                self.root.after(0, self._on_error, e, traceback.format_exc(), self.re_apply_btn)
        threading.Thread(target=task, daemon=True).start()

    def _re_on_apply(self, resp):
        self.re_apply_btn.config(state="normal")
        self.re_apply_output.delete("1.0", "end")
        data = resp.get("data") if isinstance(resp, dict) else None
        lines = []
        if isinstance(data, list):
            ok = sum(1 for d in data if isinstance(d, dict) and d.get("isSuccess"))
            lines.append(f"提交完成：成功 {ok}/{len(data)} 条。")
            for d in data:
                if isinstance(d, dict):
                    flag = "✓" if d.get("isSuccess") else "✗"
                    lines.append(f"  {flag} {d.get('name','')} {d.get('cardNum','')}  "
                                 f"{d.get('msg','')}")
        else:
            lines.append("提交完成（返回结构见下）。")
        lines.append("注意：本接口不返回 reqOrderNo，请到“② 查询申请记录”按姓名/身份证查询取回请求单号。")
        lines.append("-" * 60)
        self.re_apply_output.insert("end", "\n".join(lines) + "\n")
        try:
            self.re_apply_output.insert("end", json.dumps(resp, ensure_ascii=False, indent=2))
        except Exception:
            self.re_apply_output.insert("end", str(resp))
        self.status_var.set("提交完成")

    # ----- ② 查询申请记录 -----
    def _build_re_records(self, nb):
        frm = ttk.Frame(nb)
        nb.add(frm, text="② 查询申请记录")

        qbar = ttk.Frame(frm)
        qbar.pack(fill="x", padx=8, pady=6)
        ttk.Label(qbar, text="关键词：").pack(side="left")
        self.re_rec_search_var = tk.StringVar()
        ent = ttk.Entry(qbar, textvariable=self.re_rec_search_var, width=22)
        ent.pack(side="left", padx=4)
        ent.bind("<Return>", lambda e: self._re_go_page("first"))
        ttk.Label(qbar, text="(可匹配单号/姓名/身份证)", foreground="#666").pack(side="left")
        ttk.Label(qbar, text="  每页：").pack(side="left")
        self.re_rec_size_var = tk.StringVar(value="10")
        ttk.Entry(qbar, textvariable=self.re_rec_size_var, width=5).pack(side="left", padx=2)
        self.re_records_btn = ttk.Button(qbar, text="查询", command=lambda: self._re_go_page("first"))
        self.re_records_btn.pack(side="left", padx=6)
        ttk.Button(qbar, text="导出Excel", command=self._re_export_records).pack(side="left", padx=2)

        cols = [k for k, _ in RE_RECORD_FIELDS] + ["assetCount"]
        wrap, self.re_records_tree = self._make_tree(frm, cols, height=12, horizontal=True)
        widths = {"reqOrderNo": 230, "name": 90, "cardNum": 180, "status": 70,
                  "callbacked": 70, "callbackId": 110, "id": 150, "assetCount": 60}
        heads = dict(RE_RECORD_FIELDS)
        heads["assetCount"] = "资产条数"
        for c in cols:
            self.re_records_tree.heading(c, text=heads.get(c, c))
            self.re_records_tree.column(c, width=widths.get(c, 100), anchor="w")
        wrap.pack(fill="both", expand=True, padx=8, pady=4)

        # 分页导航条
        self.re_rec_total_page = 1
        nav = ttk.Frame(frm)
        nav.pack(fill="x", padx=8, pady=(2, 0))
        self.re_rec_info_var = tk.StringVar(value="")
        ttk.Label(nav, textvariable=self.re_rec_info_var, foreground="#666").pack(side="left")
        self.re_rec_pagebar = ttk.Frame(nav)
        self.re_rec_pagebar.pack(side="right")
        ttk.Button(self.re_rec_pagebar, text="末页", width=5,
                   command=lambda: self._re_go_page("last")).pack(side="right", padx=1)
        ttk.Button(self.re_rec_pagebar, text="下一页", width=6,
                   command=lambda: self._re_go_page("next")).pack(side="right", padx=1)
        ttk.Button(self.re_rec_pagebar, text="跳转", width=5,
                   command=lambda: self._re_go_page("jump")).pack(side="right", padx=1)
        self.re_rec_totpage_lbl = ttk.Label(self.re_rec_pagebar, text="/ 1 页")
        self.re_rec_totpage_lbl.pack(side="right", padx=(1, 4))
        self.re_rec_page_var = tk.StringVar(value="1")
        pe = ttk.Entry(self.re_rec_pagebar, textvariable=self.re_rec_page_var, width=5,
                       justify="center")
        pe.pack(side="right", padx=1)
        pe.bind("<Return>", lambda e: self._re_go_page("jump"))
        ttk.Label(self.re_rec_pagebar, text="第").pack(side="right", padx=(4, 1))
        ttk.Button(self.re_rec_pagebar, text="上一页", width=6,
                   command=lambda: self._re_go_page("prev")).pack(side="right", padx=1)
        ttk.Button(self.re_rec_pagebar, text="首页", width=5,
                   command=lambda: self._re_go_page("first")).pack(side="right", padx=1)

        # 操作条
        bbar = ttk.Frame(frm)
        bbar.pack(fill="x", padx=8, pady=4)
        ttk.Button(bbar, text="删除选中行", command=self._re_delete_selected_records).pack(side="left", padx=2)
        ttk.Button(bbar, text="清空结果", command=self._re_clear_records).pack(side="left", padx=2)
        ttk.Label(bbar, text="（删除/清空仅清理本地显示，不影响服务器记录）",
                  foreground="#999").pack(side="left", padx=6)
        ttk.Button(bbar, text="选中单号→复制", command=self._re_copy_reqno).pack(side="right", padx=2)
        ttk.Button(bbar, text="选中单号→补录页", command=self._re_send_to_upload).pack(side="right", padx=2)
        ttk.Button(bbar, text="选中单号→回调页", command=self._re_send_to_callback).pack(side="right", padx=2)

    def _re_go_page(self, where):
        """分页导航：first/prev/next/last/jump。设置页码后触发查询。"""
        try:
            cur = int(self.re_rec_page_var.get() or "1")
        except ValueError:
            cur = 1
        tp = max(1, getattr(self, "re_rec_total_page", 1) or 1)
        if where == "first":
            cur = 1
        elif where == "prev":
            cur = cur - 1
        elif where == "next":
            cur = cur + 1
        elif where == "last":
            cur = tp
        # jump: 用输入框里的值，仅做边界裁剪
        cur = max(1, min(cur, tp)) if where != "first" else max(1, cur)
        self.re_rec_page_var.set(str(cur))
        self._re_run_records()

    def _re_delete_selected_records(self):
        sel = self.re_records_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选中要删除的行（可按住 Ctrl/Shift 多选）")
            return
        for i in sorted((self.re_records_tree.index(it) for it in sel), reverse=True):
            if i < len(self.re_records):
                del self.re_records[i]
        for it in sel:
            self.re_records_tree.delete(it)
        n = len(self.re_records_tree.get_children())
        self.re_rec_info_var.set(f"本页剩 {n} 条（已删除 {len(sel)} 行）")
        self.status_var.set(f"已删除 {len(sel)} 行")

    def _re_clear_records(self):
        if not self.re_records_tree.get_children():
            return
        if not messagebox.askyesno("确认", "确认清空当前查询结果？（仅清理本地显示）"):
            return
        self.re_records = []
        for it in self.re_records_tree.get_children():
            self.re_records_tree.delete(it)
        self.re_rec_info_var.set("已清空")
        self.status_var.set("已清空查询结果")

    def _re_run_records(self):
        auth = self._auth_or_warn()
        if not auth:
            return
        try:
            page = int(self.re_rec_page_var.get() or "1")
            size = int(self.re_rec_size_var.get() or "10")
        except ValueError:
            messagebox.showwarning("提示", "页码/每页条数须为数字")
            return
        search = self.re_rec_search_var.get().strip()
        client = RealEstateClient(auth)
        self.re_records_btn.config(state="disabled")
        self.re_rec_info_var.set("查询中...")
        self.status_var.set("查询申请记录中...")

        def task():
            try:
                resp = client.query_records(page, size, search)
                self.root.after(0, self._re_on_records, resp)
            except Exception as e:
                self.root.after(0, self._on_error, e, traceback.format_exc(), self.re_records_btn)
        threading.Thread(target=task, daemon=True).start()

    def _re_on_records(self, resp):
        self.re_records_btn.config(state="normal")
        for i in self.re_records_tree.get_children():
            self.re_records_tree.delete(i)
        records = _re_extract_records(resp)
        self.re_records = records
        for rec in records:
            assets = rec.get("assets") or []
            n = len(assets) if isinstance(assets, list) else 0
            vals = [rec.get(k, "") for k, _ in RE_RECORD_FIELDS] + [n]
            self.re_records_tree.insert("", "end", values=vals)
        total = resp.get("totalCount") if isinstance(resp, dict) else None
        tpage = resp.get("totalPage") if isinstance(resp, dict) else None
        try:
            tpage = int(tpage) if tpage is not None else None
        except (ValueError, TypeError):
            tpage = None
        if tpage and tpage > 0:
            self.re_rec_total_page = tpage
            self.re_rec_totpage_lbl.config(text=f"/ {tpage} 页")
        info = f"本页 {len(records)} 条"
        if total is not None:
            info += f"，共 {total} 条"
        if tpage is not None:
            info += f" / {tpage} 页"
        self.re_rec_info_var.set(info)
        self.status_var.set("查询完成")

    def _re_selected_reqno(self):
        sel = self.re_records_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在列表中选中一行")
            return ""
        vals = self.re_records_tree.item(sel[0], "values")
        return vals[0] if vals else ""

    def _re_copy_reqno(self):
        no = self._re_selected_reqno()
        if not no:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(no)
        self.status_var.set(f"已复制单号: {no}")

    def _re_send_to_upload(self):
        no = self._re_selected_reqno()
        if no:
            self.re_upload_reqno_var.set(no)
            # 补录页默认隐藏，这里自动显示并切换过去
            if not self.re_show_admin_var.get():
                self.re_show_admin_var.set(True)
                self._re_toggle_admin()
            try:
                self.re_nb.select(self.re_tab_upload)
            except Exception:
                pass
            self.status_var.set(f"已填入补录页: {no}")

    def _re_send_to_callback(self):
        no = self._re_selected_reqno()
        if no:
            cur = self.re_cb_orders_text.get("1.0", "end").strip()
            self.re_cb_orders_text.insert("end", (("\n" if cur else "") + no))
            # ⑤ 回调页默认隐藏，加入单号时自动显示并跳过去
            if not self.re_show_admin_var.get():
                self.re_show_admin_var.set(True)
                self._re_toggle_admin()
            try:
                self.re_nb.select(self.re_tab_callback)
            except Exception:
                pass
            self.status_var.set(f"已加入回调页: {no}")

    def _re_export_records(self):
        if not self.re_records:
            messagebox.showinfo("提示", "请先查询到申请记录")
            return
        path = self._ask_save_path(
            f"不动产申请记录_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        if not path:
            return
        try:
            export_realestate_records_to_excel(self.re_records, path)
            self.status_var.set(f"已导出: {path}")
            messagebox.showinfo("完成", f"已导出到:\n{path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    # ----- ③ Excel补录资产结果 -----
    def _build_re_upload(self, nb):
        frm = ttk.Frame(nb)
        self.re_tab_upload = frm  # 由 _re_toggle_admin 控制显示/隐藏

        ttk.Label(
            frm,
            text="为申请记录补录不动产资产结果。Excel 必填列：Name、CardNum、ReqOrderNo（须与申请记录一致）；"
                 "上传时建议带上当前请求单号，服务端会校验一致性。补录成功后申请状态更新为“成功”。",
            foreground="#666", wraplength=940, justify="left",
        ).pack(fill="x", padx=8, pady=(8, 4))

        tbar = ttk.Frame(frm)
        tbar.pack(fill="x", padx=8, pady=4)
        ttk.Button(tbar, text="生成补录模板", command=self._re_gen_template).pack(side="left", padx=2)

        fbar = ttk.Frame(frm)
        fbar.pack(fill="x", padx=8, pady=4)
        ttk.Label(fbar, text="Excel文件：").pack(side="left")
        self.re_upload_path_var = tk.StringVar()
        ttk.Entry(fbar, textvariable=self.re_upload_path_var, width=52).pack(side="left", padx=4)
        ttk.Button(fbar, text="选择...", command=self._re_choose_excel).pack(side="left", padx=2)

        rbar = ttk.Frame(frm)
        rbar.pack(fill="x", padx=8, pady=4)
        ttk.Label(rbar, text="请求单号 reqOrderNo：").pack(side="left")
        self.re_upload_reqno_var = tk.StringVar()
        ttk.Entry(rbar, textvariable=self.re_upload_reqno_var, width=40).pack(side="left", padx=4)
        self.re_upload_btn = ttk.Button(rbar, text="开始补录", command=self._re_run_upload)
        self.re_upload_btn.pack(side="left", padx=6)

        self.re_upload_output = scrolledtext.ScrolledText(
            frm, wrap="word", font=("Consolas", 10), height=12)
        self.re_upload_output.pack(fill="both", expand=True, padx=8, pady=4)

    def _re_gen_template(self):
        path = self._ask_save_path("不动产补录模板.xlsx")
        if not path:
            return
        try:
            generate_re_upload_template(path)
            messagebox.showinfo("完成", f"模板已生成:\n{path}")
        except Exception as e:
            messagebox.showerror("生成失败", str(e))

    def _re_choose_excel(self):
        p = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx;*.xls"), ("All", "*.*")])
        if p:
            self.re_upload_path_var.set(p)

    def _re_run_upload(self):
        path = self.re_upload_path_var.get().strip()
        reqno = self.re_upload_reqno_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请选择有效的 Excel 文件")
            return
        auth = self._auth_or_warn()
        if not auth:
            return
        client = RealEstateClient(auth)
        self.re_upload_btn.config(state="disabled")
        self.re_upload_output.delete("1.0", "end")
        self.re_upload_output.insert("end", f"正在上传补录 {os.path.basename(path)} ...\n")
        self.status_var.set("Excel补录中...")

        def task():
            try:
                resp = client.upload_excel(path, reqno)
                self.root.after(0, self._re_on_upload, resp)
            except Exception as e:
                self.root.after(0, self._on_error, e, traceback.format_exc(), self.re_upload_btn)
        threading.Thread(target=task, daemon=True).start()

    def _re_on_upload(self, resp):
        self.re_upload_btn.config(state="normal")
        self.re_upload_output.delete("1.0", "end")
        if isinstance(resp, dict):
            code = resp.get("code", "")
            msg = resp.get("message", "")
            head = f"补录返回 code={code}"
            if msg:
                head += f"  message={msg}"
            self.re_upload_output.insert("end", head + "\n" + "-" * 60 + "\n")
            self.re_upload_output.insert("end", json.dumps(resp, ensure_ascii=False, indent=2))
        else:
            self.re_upload_output.insert("end", str(resp))
        self.status_var.set("补录完成")

    # ----- ④ 编辑资产结果 -----
    def _build_re_edit(self, nb):
        frm = ttk.Frame(nb)
        self.re_tab_edit = frm  # 由 _re_toggle_admin 控制显示/隐藏

        ttk.Label(
            frm,
            text="补录资产结果后，可按 assetId 编辑单条资产明细。assetId 必填；reqOrderNo 选填（用于二次保护）；"
                 "其余字段留空则不修改。",
            foreground="#666", wraplength=940, justify="left",
        ).pack(fill="x", padx=8, pady=(8, 4))

        form = ttk.Frame(frm)
        form.pack(fill="x", padx=8, pady=4)
        ttk.Label(form, text="资产ID assetId*：").grid(row=0, column=0, sticky="e", pady=3)
        self.re_edit_assetid_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.re_edit_assetid_var, width=28).grid(row=0, column=1, padx=4, pady=3)
        ttk.Label(form, text="请求单号 reqOrderNo：").grid(row=0, column=2, sticky="e", pady=3)
        self.re_edit_reqno_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.re_edit_reqno_var, width=30).grid(row=0, column=3, padx=4, pady=3)

        # 资产字段输入
        self.re_edit_vars = {}
        labels = dict(RE_ASSET_FIELDS)
        order = ["certNo", "unitNo", "ownership", "rightsType", "useTo",
                 "houseArea", "location", "isMortgaged"]
        for i, key in enumerate(order):
            r = 1 + i // 2
            c = (i % 2) * 2
            ttk.Label(form, text=labels[key] + "：").grid(row=r, column=c, sticky="e", pady=3)
            var = tk.StringVar()
            self.re_edit_vars[key] = var
            ttk.Entry(form, textvariable=var, width=28 if c == 0 else 30).grid(
                row=r, column=c + 1, padx=4, pady=3)
        # isSealUp 用下拉（Boolean）
        rr = 1 + len(order) // 2
        ttk.Label(form, text="是否查封 isSealUp：").grid(row=rr, column=0, sticky="e", pady=3)
        self.re_edit_seal_var = tk.StringVar(value="")
        ttk.Combobox(form, textvariable=self.re_edit_seal_var, state="readonly", width=26,
                     values=["", "true", "false"]).grid(row=rr, column=1, padx=4, pady=3)

        sbar = ttk.Frame(frm)
        sbar.pack(fill="x", padx=8, pady=6)
        self.re_edit_btn = ttk.Button(sbar, text="提交修改", command=self._re_run_edit)
        self.re_edit_btn.pack(side="left", padx=2)
        ttk.Button(sbar, text="清空", command=self._re_clear_edit).pack(side="left", padx=2)

        self.re_edit_output = scrolledtext.ScrolledText(
            frm, wrap="word", font=("Consolas", 10), height=9)
        self.re_edit_output.pack(fill="both", expand=True, padx=8, pady=4)

    def _re_clear_edit(self):
        self.re_edit_assetid_var.set("")
        self.re_edit_reqno_var.set("")
        for v in self.re_edit_vars.values():
            v.set("")
        self.re_edit_seal_var.set("")
        self.re_edit_output.delete("1.0", "end")

    def _re_run_edit(self):
        asset_id = self.re_edit_assetid_var.get().strip()
        if not asset_id:
            messagebox.showwarning("提示", "assetId 必填")
            return
        auth = self._auth_or_warn()
        if not auth:
            return
        payload = {"assetId": asset_id}
        reqno = self.re_edit_reqno_var.get().strip()
        if reqno:
            payload["reqOrderNo"] = reqno
        for key, var in self.re_edit_vars.items():
            val = var.get().strip()
            if val:
                payload[key] = val
        seal = self.re_edit_seal_var.get().strip()
        if seal:
            payload["isSealUp"] = (seal == "true")
        client = RealEstateClient(auth)
        self.re_edit_btn.config(state="disabled")
        self.re_edit_output.delete("1.0", "end")
        self.re_edit_output.insert("end", "提交字段:\n" +
                                   json.dumps(payload, ensure_ascii=False, indent=2) + "\n" +
                                   "-" * 60 + "\n正在提交...\n")
        self.status_var.set("编辑资产中...")

        def task():
            try:
                resp = client.update_asset(payload)
                self.root.after(0, self._re_on_edit, resp)
            except Exception as e:
                self.root.after(0, self._on_error, e, traceback.format_exc(), self.re_edit_btn)
        threading.Thread(target=task, daemon=True).start()

    def _re_on_edit(self, resp):
        self.re_edit_btn.config(state="normal")
        self.re_edit_output.delete("1.0", "end")
        try:
            self.re_edit_output.insert("end", json.dumps(resp, ensure_ascii=False, indent=2))
        except Exception:
            self.re_edit_output.insert("end", str(resp))
        self.status_var.set("编辑完成")

    # ----- ⑤ 处理回调 -----
    def _build_re_callback(self, nb):
        frm = ttk.Frame(nb)
        self.re_tab_callback = frm  # 由 _re_toggle_admin 控制显示/隐藏

        ttk.Label(
            frm,
            text="对已补录完成的请求单号触发回调（/callback/handle）。每行一个 reqOrderNo，可填多个；"
                 "至少一条成功即整体返回成功。",
            foreground="#666", wraplength=940, justify="left",
        ).pack(fill="x", padx=8, pady=(8, 4))

        ttk.Label(frm, text="请求单号（每行一个）：").pack(anchor="w", padx=8)
        self.re_cb_orders_text = scrolledtext.ScrolledText(
            frm, wrap="none", font=("Consolas", 10), height=6)
        self.re_cb_orders_text.pack(fill="x", padx=8, pady=4)

        sbar = ttk.Frame(frm)
        sbar.pack(fill="x", padx=8, pady=6)
        self.re_cb_btn = ttk.Button(sbar, text="执行回调", command=self._re_run_callback)
        self.re_cb_btn.pack(side="left", padx=2)
        ttk.Button(sbar, text="清空", command=lambda: self.re_cb_orders_text.delete("1.0", "end")
                   ).pack(side="left", padx=2)

        self.re_cb_output = scrolledtext.ScrolledText(
            frm, wrap="word", font=("Consolas", 10), height=12)
        self.re_cb_output.pack(fill="both", expand=True, padx=8, pady=4)

    def _re_run_callback(self):
        raw = self.re_cb_orders_text.get("1.0", "end")
        orders = [x.strip() for x in raw.splitlines() if x.strip()]
        if not orders:
            messagebox.showwarning("提示", "请至少填写一个请求单号")
            return
        auth = self._auth_or_warn()
        if not auth:
            return
        client = RealEstateClient(auth)
        self.re_cb_btn.config(state="disabled")
        self.re_cb_output.delete("1.0", "end")
        self.re_cb_output.insert("end", f"正在对 {len(orders)} 个单号执行回调...\n")
        self.status_var.set("处理回调中...")

        def task():
            try:
                resp = client.handle_callback(orders)
                self.root.after(0, self._re_on_callback, resp)
            except Exception as e:
                self.root.after(0, self._on_error, e, traceback.format_exc(), self.re_cb_btn)
        threading.Thread(target=task, daemon=True).start()

    def _re_on_callback(self, resp):
        self.re_cb_btn.config(state="normal")
        self.re_cb_output.delete("1.0", "end")
        data = resp.get("data") if isinstance(resp, dict) else None
        lines = []
        if isinstance(data, list):
            ok = sum(1 for d in data if isinstance(d, dict) and d.get("success"))
            lines.append(f"回调完成：成功 {ok}/{len(data)} 条。")
            for d in data:
                if isinstance(d, dict):
                    flag = "✓" if d.get("success") else "✗"
                    line = f"  {flag} {d.get('reqOrderNo','')}"
                    if d.get("errorMessage"):
                        line += f"  失败原因: {d.get('errorMessage')}"
                    if d.get("processTime"):
                        line += f"  处理时间: {d.get('processTime')}"
                    lines.append(line)
        else:
            lines.append("回调完成（返回结构见下）。")
        lines.append("-" * 60)
        self.re_cb_output.insert("end", "\n".join(lines) + "\n")
        try:
            self.re_cb_output.insert("end", json.dumps(resp, ensure_ascii=False, indent=2))
        except Exception:
            self.re_cb_output.insert("end", str(resp))
        self.status_var.set("回调完成")


    # =========================================================
    #  模块四：探知风控数据（apiKey/apiSecret 签名）
    # =========================================================
    def _build_tanzhi_module(self, parent_nb):
        outer = ttk.Frame(parent_nb)
        parent_nb.add(outer, text="探知风控数据")

        # —— API 凭证（与上方 Authorization 不同）——
        authf = ttk.LabelFrame(
            outer, text="探知数据 API 凭证")
        authf.pack(fill="x", padx=8, pady=(6, 2))
        row = ttk.Frame(authf)
        row.pack(fill="x", padx=6, pady=4)
        ttk.Label(row, text="apiKey：").pack(side="left")
        self.tz_key_var = tk.StringVar(value=self.cfg.get("tanzhi_api_key", ""))
        self.tz_key_entry = ttk.Entry(row, textvariable=self.tz_key_var, width=28, show="*")
        self.tz_key_entry.pack(side="left", padx=(0, 10))
        ttk.Label(row, text="apiSecret：").pack(side="left")
        self.tz_secret_var = tk.StringVar(value=self.cfg.get("tanzhi_api_secret", ""))
        self.tz_secret_entry = ttk.Entry(row, textvariable=self.tz_secret_var, width=28, show="*")
        self.tz_secret_entry.pack(side="left", padx=(0, 6))
        self.tz_show_secret = tk.BooleanVar(value=False)
        ttk.Checkbutton(row, text="显示", variable=self.tz_show_secret,
                        command=self._tz_toggle_secret).pack(side="left", padx=4)
        ttk.Button(row, text="保存", command=self._tz_save_auth).pack(side="left", padx=4)

        # —— 产品选择 ——
        selbar = ttk.Frame(outer)
        selbar.pack(fill="x", padx=8, pady=(2, 0))
        ttk.Label(selbar, text="查询产品：").pack(side="left")
        self.tz_product_var = tk.StringVar(value=TANZHI_PRODUCTS[0]["name"])
        cb = ttk.Combobox(selbar, textvariable=self.tz_product_var, state="readonly",
                          width=20, values=[p["name"] for p in TANZHI_PRODUCTS])
        cb.pack(side="left", padx=4)
        cb.bind("<<ComboboxSelected>>", self._tz_on_product_change)
        self.tz_hint_var = tk.StringVar(value="")
        ttk.Label(selbar, textvariable=self.tz_hint_var, foreground="#888").pack(side="left", padx=8)

        nb = ttk.Notebook(outer)
        nb.pack(fill="both", expand=True, padx=4, pady=4)
        self._build_tz_single(nb)
        self._build_tz_batch(nb)
        self._tz_on_product_change()

    def _tz_toggle_secret(self):
        ch = "" if self.tz_show_secret.get() else "*"
        self.tz_key_entry.config(show=ch)
        self.tz_secret_entry.config(show=ch)

    def _tz_save_auth(self):
        self.cfg["tanzhi_api_key"] = self.tz_key_var.get().strip()
        self.cfg["tanzhi_api_secret"] = self.tz_secret_var.get().strip()
        save_config(self.cfg)
        self.status_var.set("探知数据 API 凭证已保存")

    def _tz_auth_or_warn(self):
        k = self.tz_key_var.get().strip()
        s = self.tz_secret_var.get().strip()
        if not k or not s:
            messagebox.showwarning("提示", "请先填写探知数据的 apiKey 与 apiSecret")
            return None
        return k, s

    def _tz_current_product(self):
        return TANZHI_PRODUCT_BY_NAME.get(self.tz_product_var.get(), TANZHI_PRODUCTS[0])

    def _tz_on_product_change(self, event=None):
        p = self._tz_current_product()
        if p.get("input_mode") == "plate":
            self.tz_label1_var.set("VIN(车架号)：" if p.get("vin_field") else "（本产品不需要）")
            self.tz_label2_var.set("车牌号：")
            self.tz_name_entry.config(state=("normal" if p.get("vin_field") else "disabled"))
            self.tz_mobile_entry.config(state="disabled")
            req = ["车牌号"] + (["VIN(车架号)"] if p.get("vin_field") else [])
        else:
            self.tz_label1_var.set("姓名：")
            self.tz_label2_var.set("身份证号：")
            self.tz_name_entry.config(state="normal")
            self.tz_mobile_entry.config(state="normal")
            req = ["身份证号"]
            if p.get("name_req"):
                req.append("姓名")
            if p.get("mobile_use"):
                req.append("手机号")
        self.tz_hint_var.set("必填：" + "、".join(req))

    def _build_tz_single(self, nb):
        frm = ttk.Frame(nb)
        nb.add(frm, text="单个查询")
        bar = ttk.Frame(frm)
        bar.pack(fill="x", padx=8, pady=6)
        self.tz_label1_var = tk.StringVar(value="姓名：")
        ttk.Label(bar, textvariable=self.tz_label1_var).pack(side="left")
        self.tz_name_var = tk.StringVar()
        self.tz_name_entry = ttk.Entry(bar, textvariable=self.tz_name_var, width=18)
        self.tz_name_entry.pack(side="left", padx=(0, 8))
        self.tz_label2_var = tk.StringVar(value="身份证号：")
        ttk.Label(bar, textvariable=self.tz_label2_var).pack(side="left")
        self.tz_id_var = tk.StringVar()
        ttk.Entry(bar, textvariable=self.tz_id_var, width=24).pack(side="left", padx=(0, 8))
        ttk.Label(bar, text="手机号：").pack(side="left")
        self.tz_mobile_var = tk.StringVar()
        self.tz_mobile_entry = ttk.Entry(bar, textvariable=self.tz_mobile_var, width=14)
        self.tz_mobile_entry.pack(side="left", padx=(0, 8))
        self.tz_single_btn = ttk.Button(bar, text="查询", command=self._tz_run_single)
        self.tz_single_btn.pack(side="left", padx=4)
        ttk.Button(bar, text="导出本次结果到Excel",
                   command=self._tz_export_single).pack(side="left", padx=4)

        # 上：注释表  下：原始JSON
        cols = ("field", "cn", "val", "desc")
        names = ("字段", "中文名", "值", "说明")
        wrap, self.tz_tree = self._make_tree(frm, cols, height=14, horizontal=True)
        widths = {"field": 220, "cn": 220, "val": 160, "desc": 460}
        for c, n in zip(cols, names):
            self.tz_tree.heading(c, text=n)
            self.tz_tree.column(c, width=widths[c], anchor="w")
        wrap.pack(fill="both", expand=True, padx=8, pady=4)
        self._attach_tree_tooltip(self.tz_tree)
        ttk.Label(frm, text="原始返回 JSON：").pack(anchor="w", padx=8)
        self.tz_output = scrolledtext.ScrolledText(frm, wrap="word", height=8,
                                                   font=("Consolas", 10))
        self.tz_output.pack(fill="x", padx=8, pady=(0, 6))

    def _build_tz_batch(self, nb):
        frm = ttk.Frame(nb)
        nb.add(frm, text="批量查询（Excel导入）")
        top = ttk.Frame(frm)
        top.pack(fill="x", padx=8, pady=6)
        ttk.Label(top, text="Excel文件：").pack(side="left")
        self.tz_file_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.tz_file_var).pack(
            side="left", fill="x", expand=True, padx=4)
        ttk.Button(top, text="选择...", command=self._tz_choose_excel).pack(side="left", padx=2)
        ttk.Button(top, text="生成导入模板", command=self._tz_gen_template).pack(side="left", padx=2)

        bar = ttk.Frame(frm)
        bar.pack(fill="x", padx=8, pady=2)
        self.tz_batch_btn = ttk.Button(bar, text="开始批量查询", command=self._tz_run_batch)
        self.tz_batch_btn.pack(side="left", padx=2)
        self.tz_cancel_btn = ttk.Button(bar, text="取消", command=self.tz_cancel.set,
                                        state="disabled")
        self.tz_cancel_btn.pack(side="left", padx=2)
        ttk.Button(bar, text="导出全部结果到Excel",
                   command=self._tz_export_batch).pack(side="left", padx=2)
        ttk.Button(bar, text="清空结果", command=self._tz_clear).pack(side="left", padx=2)
        ttk.Label(bar, text="（按上方所选“查询产品”执行；导出按产品分Sheet，含中文注释表头）",
                  foreground="#888").pack(side="left", padx=6)

        prog = ttk.Frame(frm)
        prog.pack(fill="x", padx=8, pady=4)
        self.tz_progress = ttk.Progressbar(prog, mode="determinate")
        self.tz_progress.pack(fill="x", expand=True, side="left")
        self.tz_progress_label = ttk.Label(prog, text="0/0", width=10)
        self.tz_progress_label.pack(side="left", padx=4)

        cols = ("product", "name", "idnum", "mobile", "code", "msg")
        names = ("产品", "姓名", "身份证号", "手机号", "Code", "消息")
        wrap, self.tz_batch_tree = self._make_tree(frm, cols, height=12, horizontal=True)
        widths = {"product": 130, "name": 80, "idnum": 200, "mobile": 120,
                  "code": 70, "msg": 240}
        for c, n in zip(cols, names):
            self.tz_batch_tree.heading(c, text=n)
            self.tz_batch_tree.column(c, width=widths[c], anchor="w")
        wrap.pack(fill="both", expand=True, padx=8, pady=4)
        self._attach_tree_tooltip(self.tz_batch_tree)

    def _tz_choose_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx;*.xlsm"), ("All", "*.*")])
        if p:
            self.tz_file_var.set(p)

    def _tz_gen_template(self):
        path = self._ask_save_path("探知数据_导入模板.xlsx")
        if not path:
            return
        product = self._tz_current_product()
        try:
            if product.get("input_mode") == "plate":
                need_vin = bool(product.get("vin_field"))
                generate_tanzhi_plate_template(path, need_vin=need_vin)
                cols = "车牌号、VIN" if need_vin else "车牌号"
            elif product.get("combo"):
                generate_tanzhi_template(path, combo=True)
                cols = "姓名、身份证号"
            else:
                generate_tanzhi_template(path, combo=False)
                cols = "姓名、身份证号、手机号"
            messagebox.showinfo("完成", "模板已生成（%s）:\n%s" % (cols, path))
        except Exception as e:
            messagebox.showerror("失败", str(e))

    def _tz_run_single(self):
        product = self._tz_current_product()
        name = self.tz_name_var.get().strip()
        idnum = self.tz_id_var.get().strip()
        mobile = self.tz_mobile_var.get().strip()
        if product.get("input_mode") == "plate":
            if not idnum:
                messagebox.showwarning("提示", "请输入车牌号")
                return
            if product.get("vin_field") and not name:
                messagebox.showwarning("提示", "该产品必须提供 VIN(车架号)")
                return
        else:
            if not idnum:
                messagebox.showwarning("提示", "请输入身份证号")
                return
            if product.get("name_req") and not name:
                messagebox.showwarning("提示", "该产品“姓名”为必填")
                return
            if product.get("mobile_use") and not mobile:
                messagebox.showwarning("提示", "该产品“手机号”为必填")
                return
        creds = self._tz_auth_or_warn()
        if not creds:
            return
        client = TanzhiClient(*creds)
        self.tz_single_btn.config(state="disabled")
        self.status_var.set("查询中...")

        def task():
            try:
                if product.get("combo"):
                    info, resp = client.query_chain(product, name, idnum, mobile)
                else:
                    info, resp = client.query_one(product, name, idnum, mobile)
                self.root.after(0, self._tz_on_single_done, info, resp)
            except Exception as e:
                self.root.after(0, self._on_error, e, traceback.format_exc(),
                                self.tz_single_btn)
        threading.Thread(target=task, daemon=True).start()

    def _tz_on_single_done(self, info, resp):
        self.tz_single_btn.config(state="normal")
        self.tz_last = (info, resp)
        ok = info.get("code") == "0000"
        self.status_var.set("查询完成（%s %s）" % (info.get("code", ""), info.get("msg", "")))
        # 注释表
        for i in self.tz_tree.get_children():
            self.tz_tree.delete(i)
        product = TANZHI_PRODUCT_BY_NAME.get(info.get("product", ""), {})
        lookup = tanzhi_lookup(self.tanzhi_dict, product)
        data = tanzhi_extract_data(resp)
        if data is not None:
            for key, val in tanzhi_flatten(data):
                ann = tanzhi_annotate(lookup, key) or {}
                code_note = tanzhi_code_note(key, val)
                expl = "；".join(x for x in (code_note, ann.get("desc", ""), ann.get("note", "")) if x)
                self.tz_tree.insert("", "end", values=(key, ann.get("cn", ""), val, expl))
        # 原始JSON
        self.tz_output.delete("1.0", "end")
        head = ("产品: %s\ncode: %s   msg: %s   token: %s\n%s\n" % (
            info.get("product", ""), info.get("code", ""), info.get("msg", ""),
            info.get("token", ""), "-" * 60))
        self.tz_output.insert("end", head)
        if resp is not None:
            try:
                self.tz_output.insert("end", json.dumps(resp, ensure_ascii=False, indent=2))
            except Exception:
                self.tz_output.insert("end", str(resp))
        if not ok and info.get("msg"):
            pass

    def _tz_export_single(self):
        if not self.tz_last:
            messagebox.showinfo("提示", "请先执行一次单个查询")
            return
        info, resp = self.tz_last
        path = self._ask_save_path("探知_%s_%s.xlsx" % (
            info.get("product", ""), datetime.now().strftime("%Y%m%d_%H%M%S")))
        if not path:
            return
        try:
            export_tanzhi_single(info, resp, self.tanzhi_dict, path)
            self.status_var.set("已导出: %s" % path)
            messagebox.showinfo("完成", "已导出到:\n%s" % path)
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def _tz_run_batch(self):
        product = self._tz_current_product()
        path = self.tz_file_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请选择有效的Excel文件")
            return
        creds = self._tz_auth_or_warn()
        if not creds:
            return
        is_plate = product.get("input_mode") == "plate"
        try:
            persons = read_tanzhi_plate_rows(path) if is_plate else read_persons_from_excel(path)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return
        if not persons:
            need = "车牌号[/VIN]" if is_plate else "姓名/身份证号/手机号"
            messagebox.showinfo("提示", "未在Excel中读取到数据（需有 %s 列）" % need)
            return
        if not messagebox.askyesno(
                "确认", "产品：%s\n共读取到 %d 条，是否开始查询？" % (product["name"], len(persons))):
            return
        client = TanzhiClient(*creds)
        self.tz_cancel.clear()
        self.tz_batch_btn.config(state="disabled")
        self.tz_cancel_btn.config(state="normal")
        self.tz_progress["maximum"] = len(persons)
        self.tz_progress["value"] = 0
        self.tz_progress_label.config(text="0/%d" % len(persons))

        def task():
            done = 0
            for p in persons:
                if self.tz_cancel.is_set():
                    break
                try:
                    if product.get("combo"):
                        info, resp = client.query_chain(
                            product, p.get("name", ""), p.get("idCard", ""), p.get("mobile", ""))
                    else:
                        info, resp = client.query_one(
                            product, p.get("name", ""), p.get("idCard", ""), p.get("mobile", ""))
                except Exception as e:
                    info = {"product": product["name"], "name": p.get("name", ""),
                            "idnum": p.get("idCard", ""), "mobile": p.get("mobile", ""),
                            "code": "异常", "msg": str(e), "token": ""}
                    resp = None
                done += 1
                self.root.after(0, self._tz_on_batch_one, info, resp, done, len(persons))
            self.root.after(0, self._tz_on_batch_done)
        threading.Thread(target=task, daemon=True).start()

    def _tz_on_batch_one(self, info, resp, done, total):
        self.tz_results.append((info, resp))
        self._tz_append_row(info)
        self.tz_progress["value"] = done
        self.tz_progress_label.config(text="%d/%d" % (done, total))

    def _tz_on_batch_done(self):
        self.tz_batch_btn.config(state="normal")
        self.tz_cancel_btn.config(state="disabled")
        self.status_var.set("批量查询结束")
        messagebox.showinfo("完成", "批量查询已结束，可点击导出按钮保存结果")

    def _tz_append_row(self, info):
        self.tz_batch_tree.insert("", "end", values=(
            info.get("product", ""), info.get("name", ""), info.get("idnum", ""),
            info.get("mobile", ""), info.get("code", ""), info.get("msg", "")))

    def _tz_export_batch(self):
        if not self.tz_results:
            messagebox.showinfo("提示", "暂无结果可导出")
            return
        path = self._ask_save_path(
            "探知数据查询结果_%s.xlsx" % datetime.now().strftime("%Y%m%d_%H%M%S"))
        if not path:
            return
        try:
            export_tanzhi_batch(self.tz_results, self.tanzhi_dict, path)
            self.status_var.set("已导出: %s" % path)
            messagebox.showinfo("完成", "已导出到:\n%s" % path)
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def _tz_clear(self):
        if not self.tz_results:
            return
        if not messagebox.askyesno("确认", "确认清空探知数据批量结果？"):
            return
        self.tz_results.clear()
        for i in self.tz_batch_tree.get_children():
            self.tz_batch_tree.delete(i)
        self.tz_progress["value"] = 0
        self.tz_progress_label.config(text="0/0")
        self.status_var.set("已清空")


def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()




