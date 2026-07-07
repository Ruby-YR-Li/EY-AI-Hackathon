"""
费用底稿Review Agent - 规则核对逻辑元数据
为每条规则提供结构化核对逻辑描述，供网页弹窗和Excel核对逻辑Sheet使用。
"""

from typing import Optional, Dict

# 规则ID → 核对逻辑元数据
# source: 取数Sheet
# position: 取数位置/关键词
# data: 取数内容描述
# method: 核对方法
# pass_cond: 检查通过条件
# fail_reason: 问题原因
RULE_CHECK_LOGIC: Dict[str, dict] = {
    # ── PSP执行完整性 ────────────────────────────────────
    "UEXP-PSP-001": {
        "source": "汇总",
        "position": "G列（执行状态）",
        "data": "汇总页中每条PSP的执行状态（是/否）",
        "method": "遍历汇总页G列，确认所有PSP的执行状态均已填写",
        "pass_cond": "所有PSP的G列均有值（是/否）；不存在空行",
        "fail_reason": "存在未标记执行状态的PSP",
    },
    "UEXP-PSP-002": {
        "source": "汇总 + BKD",
        "position": "汇总G/H列、BKD数据",
        "data": "不执行的PSP及其拒绝理由，结合BKD数据验证",
        "method": "AI判断：PSP拒绝理由是否充分，与BKD数据是否一致",
        "pass_cond": "AI判断所有拒绝理由均合理、具体、充分",
        "fail_reason": "存在拒绝理由不充分或与实际数据矛盾的PSP",
    },
    # ── 格式与基础信息 ────────────────────────────────────
    "UEXP-FMT-001": {
        "source": "Uexp.00 Lead",
        "position": "C2-C8（客户名称、期末、TE/SAD、会计准则、本位币）",
        "data": "Lead Sheet基本信息的填写情况",
        "method": "按行读取关键字段值，检查是否为空值；会计准则应填写具体准则名称（如'企业会计准则'），不能为'CNY'（货币代码）；记账本位币应填写货币名称（如'人民币'），不能为'IFRS'（准则代码）",
        "pass_cond": "客户名称、期末、分析日期、TE、SAD、会计准则、记账本位币7个字段全部非空；会计准则≠'CNY'，记账本位币≠'IFRS'",
        "fail_reason": "Lead Sheet基本信息字段存在空缺或填写错误",
    },
    "UEXP-FMT-002": {
        "source": "Uexp.00 Lead",
        "position": "C3（期末）、C4（分析日期）",
        "data": "分析日期与期末日期",
        "method": "比较C3（期末/资产负债表日）和C4（分析日期）的值",
        "pass_cond": "分析日期 ≠ 期末日期",
        "fail_reason": "分析日期与期末日期相同，分析日期应为底稿编制日期而非资产负债表日当天",
    },
    "UEXP-FMT-003": {
        "source": "汇总",
        "position": "PSP执行状态列",
        "data": "汇总页中PSP程序代码和执行状态",
        "method": "检查汇总页中PSP的G列（执行）是否至少有一个为'是'",
        "pass_cond": "G列中存在'是'标记的PSP",
        "fail_reason": "所有PSP均标记为不执行或未标记",
    },
    "UEXP-FMT-004": {
        "source": "汇总",
        "position": "G列/H列",
        "data": "PSP执行状态和拒绝理由",
        "method": "检查G列执行状态是否全部填写；检查不执行的PSP在H列是否填写理由",
        "pass_cond": "汇总页所有PSP行G列均已填写（是/否）；G列为'否'的行H列有拒绝理由",
        "fail_reason": "PSP执行状态未完整填写或拒绝理由缺失",
    },
    # ── 勾稽关系 ────────────────────────────────────────
    "UEXP-GL-001": {
        "source": "Uexp.00 Lead + BKD",
        "position": "Lead C26/C27（销售费用、管理费用），BKD合计行",
        "data": "Lead Sheet中销售费用/管理费用本期账面数，VC.00/VD.00 BKD合计行金额",
        "method": "分别提取Lead C26（销售费用）、C27（管理费用）的本期账面数，与对应BKD合计行的本期账面数比较",
        "pass_cond": "Leader本期账面数与BKD合计金额的差异 ≤ max(SAD×1%, 100)",
        "fail_reason": "Lead Sheet本期账面数与BKD合计金额不一致",
    },
    "UEXP-GL-002": {
        "source": "Uexp.00 Lead",
        "position": "B14-C18 签核标记区",
        "data": "TB/PY/A3/A5/FS标记和签核内容",
        "method": "检查B列标签（TB/PY/A3/A5/FS）是否完整；检查C列是否有签核确认文字",
        "pass_cond": "B14-B18标签区域完整，且C列对应位置有签核确认内容",
        "fail_reason": "Lead Sheet签核标记不完整或A3核对确认缺失",
    },
    "UEXP-GL-003": {
        "source": "Uexp.00 Lead",
        "position": "C5（TE）、C6（SAD）",
        "data": "可容忍误差（TE）和名义金额（SAD）",
        "method": "提取TE和SAD值，检查是否为正数且SAD＜TE",
        "pass_cond": "TE > 0 且 SAD > 0 且 SAD < TE",
        "fail_reason": "TE/SAD值缺失或异常（TE≤0、SAD≤0、SAD≥TE）",
    },
    "UEXP-GL-004": {
        "source": "Uexp.00 Lead + 用户配置 + BKD",
        "position": "C5（TE）、C6（SAD）、BKD CRA行",
        "data": "用户输入的TE/SAD和CRA配置，底稿Lead Sheet中的TE/SAD值，BKD中各认定的CRA水平",
        "method": "比较用户输入的TE/SAD与Lead Sheet填列值是否一致；比较用户配置的CRA与BKD中各认定的CRA水平是否一致",
        "pass_cond": "用户配置的TE/SAD/CRA与底稿填列内容一致（差异<1%）",
        "fail_reason": "用户输入的TE/SAD/CRA与底稿不一致",
    },
    # ── BKD波动分析 ──────────────────────────────────────
    "UEXP-BKD-001": {
        "source": "VC.00 销售费用BKD / VD.00 管理费用BKD",
        "position": "预期记录区域（B17附近）",
        "data": "BKD中对费用变动的预期描述",
        "method": "扫描BKD中'预期'相关行，检查是否有超过20字的预期描述内容",
        "pass_cond": "预期区域有详细的预期描述，含预期方向、依据和理由",
        "fail_reason": "BKD中未记录预期或仅记录结论性描述，缺少依据和理由",
    },
    "UEXP-BKD-002": {
        "source": "VC.00 销售费用BKD / VD.00 管理费用BKD",
        "position": "BKD表1的数据行排列顺序",
        "data": "各费用明细项的排列顺序",
        "method": "检查费用明细按科目编码排序的规律；检查审定金额是否降序排列",
        "pass_cond": "费用明细按金额大小降序排列（非科目编码顺序）",
        "fail_reason": "费用明细按科目编码排列而非按金额大小排序",
    },
    "UEXP-BKD-003": {
        "source": "VC.00 销售费用BKD / VD.00 管理费用BKD",
        "position": "P列（进一步调查基于波动幅度？）",
        "data": "费用明细项变动金额和变动率",
        "method": "逐行计算变动金额和变动率，与波动阈值比较；检查超过阈值的项目P列是否标记'是'",
        "pass_cond": "所有超过波动阈值且变动率＞10%的项目均在P列标记为'是'",
        "fail_reason": "存在超过波动阈值但未在P列标记的项目",
    },
    "UEXP-BKD-004": {
        "source": "VC.00 销售费用BKD / VD.00 管理费用BKD",
        "position": "Q列（进一步调查基于定性考虑？）",
        "data": "费用明细项名称和金额",
        "method": "识别本期新增费用项（上期=0）、'其他'项金额大、法律诉讼等异常项目，检查Q列是否标记",
        "pass_cond": "所有新增、异常费用项均在Q列标记为'是'",
        "fail_reason": "存在应做定性分析但未在Q列标记的费用项目",
    },
    "UEXP-BKD-005": {
        "source": "VC.00 销售费用BKD / VD.00 管理费用BKD",
        "position": "O列（Notes）、下方Notes区域",
        "data": "标记为'是'的项目的Notes索引和ARP分析",
        "method": "检查P列或Q列为'是'的项目，其O列是否填写Note索引，下方是否有ARP分析记录",
        "pass_cond": "每个标记项在O列有Note索引，且下方有对应的ARP分析记录",
        "fail_reason": "标记项缺少Notes索引或ARP分析记录",
    },
    "UEXP-BKD-006": {
        "source": "VC.00 销售费用BKD / VD.00 管理费用BKD",
        "position": "BKD底部'波动说明'区域",
        "data": "总体波动分析说明",
        "method": "查找'波动说明'行，检查下方是否有详细的波动分析内容",
        "pass_cond": "波动说明区域有详细的变动分析，含金额、比例、原因和索引",
        "fail_reason": "波动说明内容不充分或过于简略",
    },
    "UEXP-BKD-008": {
        "source": "VC.00 销售费用BKD / VD.00 管理费用BKD",
        "position": "N列（For X-ref）",
        "data": "职工薪酬、折旧摊销等项目的交叉索引",
        "method": "检查职工薪酬、折旧摊销等已在其他底稿执行程序的项目，N列是否填写交叉索引",
        "pass_cond": "所有需交叉索引的项目均在N列填写了X-ref",
        "fail_reason": "存在应交叉索引但缺失的项目",
    },
    # ── TOD ──────────────────────────────────────────────
    "UEXP-TOD-001": {
        "source": "VC&VD.01.2 TOD 预审+剩余期间",
        "position": "D列（总体名称）、F/G/H列（金额）",
        "data": "TOD样本总体的构成",
        "method": "检查D列列示的样本总体名称，识别是否包含非同一SCOT的科目（制造费用、研发支出等）",
        "pass_cond": "样本总体仅包含同一SCOT的费用类型（销售费用/管理费用）",
        "fail_reason": "样本总体混合了不同SCOT的费用类型",
    },
    "UEXP-TOD-002": {
        "source": "VC&VD.01.2 TOD 预审+剩余期间",
        "position": "E列（数据来源 X-ref）",
        "data": "每个样本总体的数据来源索引",
        "method": "检查E列是否填写数据来源X-ref（如<VC.00销售费用BKD>）",
        "pass_cond": "每个样本总体均在E列填写了数据来源X-ref",
        "fail_reason": "存在未填写数据来源X-ref的样本总体",
    },
    # ── 截止性测试 ────────────────────────────────────────
    "UEXP-CO-001": {
        "source": "VC&VD.01.4 截止性测试",
        "position": "表1（确定涵盖期间）",
        "data": "截止测试期间的确定依据",
        "method": "检查表1中是否有对测试期间确定依据的文字说明",
        "pass_cond": "存在详细的期间确定依据记录（>30字）",
        "fail_reason": "未记录截止性测试期间的确定依据",
    },
    "UEXP-CO-002": {
        "source": "VC&VD.01.4 截止性测试",
        "position": "表2（记录策略）",
        "data": "截止测试样本选取策略",
        "method": "检查表2中是否有'策略''key item''关键项'等字样的策略描述",
        "pass_cond": "存在截止测试选样策略记录",
        "fail_reason": "未记录截止测试样本选取策略",
    },
    "UEXP-CO-003": {
        "source": "VC&VD.01.4 截止性测试",
        "position": "表3（测试样本明细）",
        "data": "截止测试选取的样本金额",
        "method": "提取资产负债表日后样本的金额，检查是否筛选了超过测试阈值的交易",
        "pass_cond": "截止测试样本中包含金额较大的交易（接近或超过测试阈值）",
        "fail_reason": "截止测试样本金额均较小，可能未充分筛选关键项",
    },
    # ── AI规则 ────────────────────────────────────────────
    "UEXP-AI-001": {
        "source": "BKD",
        "position": "各费用明细项",
        "data": "大幅波动的费用项目",
        "method": "AI判断：大幅波动的合理性及是否需要进一步调查",
        "pass_cond": "AI判断波动均合理或已在底稿中分析",
        "fail_reason": "AI识别到需进一步调查的异常波动",
    },
    "UEXP-AI-002": {
        "source": "BKD",
        "position": "科目名称",
        "data": "可能存在分类错误风险的科目",
        "method": "AI判断：费用科目分类是否符合会计准则（如运输费/质保费/佣金等）",
        "pass_cond": "AI判断未发现分类异常",
        "fail_reason": "AI识别到可能存在分类错误的费用项目",
    },
    "UEXP-AI-003": {
        "source": "VD.00 管理费用BKD + 汇总",
        "position": "管理费用BKD + 汇总页PSP状态",
        "data": "法律费用相关科目及PSP执行情况",
        "method": "AI判断：BKD中法律诉讼费用与PSP执行状态是否一致",
        "pass_cond": "BKD无法律费用或PSP已执行",
        "fail_reason": "BKD存在法律费用但对应PSP拒绝执行",
    },
}


def _location_text(sheet: str = "", cell: str = "") -> str:
    if sheet and cell:
        return f"{sheet}!{cell}"
    if sheet:
        return sheet
    return ""


def build_execution_summary(rule_id: str, message: str = "", expected: str = "",
                            actual: str = "", sheet: str = "", cell: str = "",
                            context: Optional[dict] = None,
                            status: str = "FAILED") -> Optional[dict]:
    """构建面向使用者的精简核对摘要：取数、方法、结论。"""
    logic = RULE_CHECK_LOGIC.get(rule_id)
    if not logic:
        return None
    if status == "PASSED":
        conclusion = logic["pass_cond"]
    elif status == "FAILED":
        conclusion = logic["fail_reason"]
    elif status in ("SKIPPED", "AI_DISABLED"):
        conclusion = "未检查（" + status + "）"
    else:
        conclusion = ""
    return {
        "source": logic["source"],
        "position": logic["position"],
        "data": logic["data"],
        "method": logic["method"],
        "pass_cond": logic["pass_cond"],
        "fail_reason": logic["fail_reason"],
        "conclusion": conclusion,
    }


def build_execution_detail(rule_id: str, message: str = "", expected: str = "",
                           actual: str = "", sheet: str = "", cell: str = "",
                           context: Optional[dict] = None,
                           status: str = "FAILED") -> str:
    """构建结构化的执行详情文本（含取数 -> 方法 -> 结论）。"""
    summary = build_execution_summary(
        rule_id=rule_id,
        message=message,
        expected=expected,
        actual=actual,
        sheet=sheet,
        cell=cell,
        context=context,
        status=status,
    )
    if not summary:
        return ""
    return "\n".join([
        f"取数：{summary['data']}",
        f"方法：{summary['method']}",
        f"结论：{summary['conclusion']}",
    ])
