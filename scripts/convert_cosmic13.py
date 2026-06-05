"""
COSMIC Excel → Word 需求文档生成器（通用版）
自动检测列名格式，支持：
  - 7列格式：一级模块、二级模块、三级模块、功能过程、子过程描述、数据组、数据属性
  - 13列格式：功能需求主键、一级模块、二级模块、三级模块、功能用户、OPEX需求、触发事件、
              功能过程、子过程描述、数据移动类型、数据组、数据属性、CFP
层级：一级模块(H2) → 二级模块(H3) → 三级模块(H4) → 功能过程(H5) → 子过程描述(段落)
"""

import os
import pandas as pd
from collections import OrderedDict


def _detect_columns(df: pd.DataFrame) -> dict:
    """
    自动检测Excel列名映射，兼容7列和13列格式。

    Returns:
        dict: 列名映射，如 {"一级模块": "一级模块", "功能过程": "功能过程", ...}
    """
    cols = list(df.columns)

    # 基础层级列（两种格式名称一致）
    mapping = {
        "一级模块": None,
        "二级模块": None,
        "三级模块": None,
        "功能过程": None,
        "子过程描述": None,
        "数据组": None,
        "数据属性": None,
    }

    # 可选列（13列格式独有）
    optional = {
        "功能用户": None,
        "OPEX需求": None,
        "触发事件": None,
        "数据移动类型": None,
        "CFP": None,
    }

    # 列名关键词映射
    keyword_map = {
        "一级模块": ["一级"],
        "二级模块": ["二级"],
        "三级模块": ["三级"],
        "功能过程": ["功能过程"],
        "子过程描述": ["子过程描述", "子过程"],
        "数据组": ["数据组"],
        "数据属性": ["数据属性"],
        "功能用户": ["功能用户"],
        "OPEX需求": ["OPEX", "功能用户需求"],
        "触发事件": ["触发事件"],
        "数据移动类型": ["数据移动类型"],
        "CFP": ["CFP"],
    }

    for target_key, keywords in keyword_map.items():
        for col in cols:
            col_str = str(col)
            if col_str == target_key:
                if target_key in mapping:
                    mapping[target_key] = col
                else:
                    optional[target_key] = col
                break
            for kw in keywords:
                if kw in col_str:
                    if target_key in mapping:
                        mapping[target_key] = col
                    else:
                        optional[target_key] = col
                    break
            if target_key in mapping and mapping[target_key] is not None:
                break

    # 也尝试匹配四级模块/五级模块（标准COSMIC模板）
    if mapping["功能过程"] is None:
        for col in cols:
            if "四级" in str(col):
                mapping["功能过程"] = col
                break
    if mapping["子过程描述"] is None:
        for col in cols:
            if "五级" in str(col):
                mapping["子过程描述"] = col
                break

    result = {**mapping, **optional}
    return result


def convert_cosmic13_to_word(input_xlsx: str, output_docx: str) -> str:
    """
    将COSMIC Excel（7列或13列）转换为Word需求文档。

    Args:
        input_xlsx: 输入Excel文件路径
        output_docx: 输出Word文件路径

    Returns:
        生成的Word文件路径
    """
    import html

    # ---- 读取并预处理数据 ----
    df = pd.read_excel(input_xlsx, sheet_name=0)
    cm = _detect_columns(df)

    # 前向填充层级列
    for key in ["一级模块", "二级模块", "三级模块"]:
        col = cm[key]
        if col and col in df.columns:
            df[col] = df[col].ffill()

    l1_col = cm["一级模块"]
    l2_col = cm["二级模块"]
    l3_col = cm["三级模块"]
    fp_col = cm["功能过程"]
    sub_col = cm["子过程描述"]
    dg_col = cm["数据组"]
    da_col = cm["数据属性"]
    opex_col = cm["OPEX需求"]
    trigger_col = cm["触发事件"]
    dm_col = cm["数据移动类型"]
    cfp_col = cm["CFP"]

    # 如果列名未检测到，回退到标准列名
    if l1_col is None:
        l1_col = "一级模块"
    if l2_col is None:
        l2_col = "二级模块"
    if l3_col is None:
        l3_col = "三级模块"
    if fp_col is None:
        fp_col = "功能过程"
    if sub_col is None:
        sub_col = "子过程描述"
    if dg_col is None:
        dg_col = "数据组"
    if da_col is None:
        da_col = "数据属性"

    # ---- 构建层级映射 ----
    l1_to_l2 = OrderedDict()
    l2_to_l3 = OrderedDict()
    l3_to_fp = OrderedDict()

    for _, row in df.iterrows():
        l1 = row.get(l1_col, None)
        l2 = row.get(l2_col, None)
        l3 = row.get(l3_col, None)
        fp = row.get(fp_col, None)

        if pd.notna(l1) and pd.notna(l2):
            if l1 not in l1_to_l2:
                l1_to_l2[l1] = []
            if l2 not in l1_to_l2[l1]:
                l1_to_l2[l1].append(l2)

        if pd.notna(l2) and pd.notna(l3):
            if l2 not in l2_to_l3:
                l2_to_l3[l2] = []
            if l3 not in l2_to_l3[l2]:
                l2_to_l3[l2].append(l3)

        if pd.notna(l3) and pd.notna(fp):
            if l3 not in l3_to_fp:
                l3_to_fp[l3] = []
            if fp not in l3_to_fp[l3]:
                l3_to_fp[l3].append(fp)

    # ---- 辅助函数 ----
    def esc(t):
        return html.escape(str(t)) if pd.notna(t) and str(t).strip() else ""

    current = {"l1": "", "l2": "", "l3": "", "fp": ""}
    parts = []

    for idx, row in df.iterrows():
        l1 = row.get(l1_col, None)
        l2 = row.get(l2_col, None)
        l3 = row.get(l3_col, None)
        fp = row.get(fp_col, None)
        sub_desc = row.get(sub_col, None)
        dm_type = row.get(dm_col, None) if dm_col else None
        data_group = row.get(dg_col, None)
        data_attr = row.get(da_col, None)
        cfp = row.get(cfp_col, None) if cfp_col else None

        # (1) 一级标题 H2
        if pd.notna(l1) and str(l1).strip() and l1 != current["l1"]:
            parts.append(f"<h2>{esc(l1)}</h2>")
            titles = l1_to_l2.get(l1, [])
            if titles:
                parts.append(f"<p><b>需求说明：</b>{esc(l1)}模块包括：{esc('，'.join(titles))}。</p>")
            current["l1"] = l1
            current["l2"] = ""
            current["l3"] = ""
            current["fp"] = ""

        # (2) 二级标题 H3
        if pd.notna(l2) and str(l2).strip() and l2 != current["l2"]:
            parts.append(f"<h3>{esc(l2)}</h3>")
            titles = l2_to_l3.get(l2, [])
            if titles:
                parts.append(f"<p><b>需求说明：</b>{esc(l2)}功能包括：{esc('，'.join(titles))}。</p>")
            current["l2"] = l2
            current["l3"] = ""
            current["fp"] = ""

        # (3) 三级标题 H4
        if pd.notna(l3) and str(l3).strip() and l3 != current["l3"]:
            parts.append(f"<h4>{esc(l3)}</h4>")
            # 三级需求说明
            sub = df[df[l3_col] == l3]
            desc_parts = []
            if opex_col and opex_col in df.columns:
                opex_vals = sub[opex_col].dropna().unique()
                if len(opex_vals) > 0:
                    desc_parts.append(f"用户需求包括：{esc('；'.join(opex_vals[:3]))}")
            if trigger_col and trigger_col in df.columns:
                trigger_vals = sub[trigger_col].dropna().unique()
                if len(trigger_vals) > 0:
                    desc_parts.append(f"触发事件包括：{esc('；'.join(trigger_vals[:3]))}")
            fps = l3_to_fp.get(l3, [])
            if fps:
                desc_parts.append(f"包含{len(fps)}个功能过程")
            if desc_parts:
                parts.append(f"<p><b>需求说明：</b>{'。'.join(desc_parts)}。</p>")
            current["l3"] = l3
            current["fp"] = ""

        # (4) 功能过程 H5
        if pd.notna(fp) and str(fp).strip() and fp != current["fp"]:
            parts.append(f"<h5>{esc(fp)}</h5>")
            fp_desc = []
            if opex_col and opex_col in df.columns:
                fp_sub = df[(df[l3_col] == l3) & (df[fp_col] == fp)]
                fp_opex = fp_sub[opex_col].dropna().unique()
                if len(fp_opex) > 0:
                    fp_desc.append(esc(fp_opex[0]))
            if trigger_col and trigger_col in df.columns:
                fp_sub2 = df[(df[l3_col] == l3) & (df[fp_col] == fp)]
                fp_trigger = fp_sub2[trigger_col].dropna().unique()
                if len(fp_trigger) > 0:
                    fp_desc.append(f"触发事件：{esc(fp_trigger[0])}")
            if fp_desc:
                parts.append(f"<p><b>需求说明：</b>{'；'.join(fp_desc)}。</p>")
            current["fp"] = fp

        # (5) 子过程描述段落（纯内容，不加"需求说明："前缀）
        if pd.notna(sub_desc) and str(sub_desc).strip():
            detail_parts = [f"<b>{esc(sub_desc)}</b>"]
            if pd.notna(dm_type):
                detail_parts.append(f"数据移动类型：{esc(dm_type)}")
            if pd.notna(data_group):
                detail_parts.append(f"数据组：{esc(data_group)}")
            if pd.notna(data_attr):
                detail_parts.append(f"数据属性：{esc(data_attr)}")
            if pd.notna(cfp):
                detail_parts.append(f"CFP：{int(cfp)}")
            line = "；".join(detail_parts) + "。"
            parts.append(f'<p style="text-indent:2em;margin:4px 0;line-height:1.6;">{line}</p>')

    body_content = "\n".join(parts)

    # ---- 统计摘要 ----
    total_cfp = df[cfp_col].sum() if cfp_col and cfp_col in df.columns else 0
    total_fp = df[fp_col].dropna().nunique() if fp_col in df.columns else 0
    l3_count = df[l3_col].nunique() if l3_col in df.columns else 0

    cfp_info = f"，合计{int(total_cfp)}CFP" if total_cfp > 0 else ""

    # ---- 组装完整HTML ----
    html_full = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>全双工二期工程需求描述</title>
<style>
body {{
    font-family: "宋体", SimSun, "Microsoft YaHei", Arial, sans-serif;
    font-size: 10.5pt;
    line-height: 1.8;
    color: #333333;
    margin: 40px 60px;
}}
h1 {{
    font-size: 18pt;
    color: #000000;
    border-bottom: 2px solid #000000;
    padding-bottom: 6px;
    margin-top: 36px;
    margin-bottom: 12px;
}}
h2 {{
    font-size: 16pt;
    color: #000000;
    border-bottom: 1px solid #666666;
    padding-bottom: 4px;
    margin-top: 28px;
    margin-bottom: 10px;
}}
h3 {{
    font-size: 14pt;
    color: #333333;
    margin-top: 22px;
    margin-bottom: 8px;
}}
h4 {{
    font-size: 12pt;
    color: #333333;
    font-weight: bold;
    margin-top: 16px;
    margin-bottom: 6px;
}}
h5 {{
    font-size: 11pt;
    color: #444444;
    font-weight: bold;
    margin-top: 12px;
    margin-bottom: 4px;
}}
p {{
    font-size: 10.5pt;
    text-indent: 2em;
    margin: 4px 0;
    line-height: 1.6;
}}
</style>
</head>
<body>
<h1>全双工二期工程需求描述</h1>
<p><b>需求说明：</b>本文档基于全双工二期COSMIC功能点拆分表，按照一级模块、二级模块、三级模块、功能过程、子过程的层级结构，对各功能模块进行系统性描述。共包含{len(l1_to_l2)}个一级模块、{sum(len(v) for v in l1_to_l2.values())}个二级模块、{l3_count}个三级模块、{total_fp}个功能过程{cfp_info}。</p>
{body_content}
</body>
</html>"""

    # ---- 写入临时HTML ----
    temp_html = output_docx.rsplit(".", 1)[0] + "_temp.html"
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(html_full)

    # ---- 使用html2docx转换 ----
    import sys
    skill_path = os.getenv("skill_path", r"C:\Users\小黑\AppData\Roaming\WPS 灵犀\serverdir\skills")
    sys.path.insert(0, os.path.join(skill_path, "docx", "scripts"))
    from html2docx import html2docx

    result = html2docx(input_html=temp_html, output_docx=output_docx)

    # 清理临时HTML
    try:
        os.remove(temp_html)
    except OSError:
        pass

    if not result.get("ok"):
        raise RuntimeError(f"Word转换失败: {result.get('message', '未知错误')}")

    return output_docx


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        print("用法: python convert_cosmic13.py <输入.xlsx> <输出.docx>")
        sys.exit(1)
    output = convert_cosmic13_to_word(sys.argv[1], sys.argv[2])
    print(f"生成完成: {output}")
