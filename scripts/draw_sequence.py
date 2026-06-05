"""
时序图绘制引擎 v3
功能：从COSMIC Excel数据推导参与者（4-6个）和交互消息（20-25条），绘制紧凑时序图PNG。
特点：宽度自适应、消息文字精确居中、参与者按业务语义分类。
依赖：Pillow (PIL), pandas
"""

import os
import re
import math
from PIL import Image, ImageDraw, ImageFont

# ========== 全局配置 ==========
FONT_PATH = None
BG_COLOR = (245, 249, 255)
PART_BG = (225, 237, 255)
PART_BORDER = (90, 130, 200)
LABEL_COLOR = (40, 40, 40)
LIFELINE_COLOR = (190, 190, 190)
TEXT_COLOR = (50, 50, 50)
ARROW_COLOR = (70, 70, 70)
DPI = 150


# ========== 字体 ==========
def get_font(size, bold=False):
    global FONT_PATH
    if FONT_PATH is None:
        for c in ['C:/Windows/Fonts/msyh.ttc', 'msyh.ttc']:
            if os.path.exists(c):
                FONT_PATH = c
                break
    try:
        return ImageFont.truetype(FONT_PATH, size)
    except Exception:
        return ImageFont.load_default()


# ========== 绘图辅助 ==========
def draw_rounded_rect(draw, xy, radius, fill, outline=None, width=1):
    x1, y1, x2, y2 = xy
    r = int(min(radius, (x2 - x1) // 2, (y2 - y1) // 2))
    draw.rounded_rectangle(xy, radius=r, fill=fill, outline=outline, width=int(width))


def draw_dashed_line(draw, x1, y1, x2, y2, color, width=1, dash_len=6, gap_len=4):
    length = math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
    if length == 0:
        return
    dx = (x2 - x1) / length
    dy = (y2 - y1) / length
    pos = 0
    while pos < length:
        end = min(pos + dash_len, length)
        draw.line(
            [(int(x1 + dx * pos), int(y1 + dy * pos)),
             (int(x1 + dx * end), int(y1 + dy * end))],
            fill=color, width=int(width),
        )
        pos += dash_len + gap_len


def draw_arrow(draw, x1, y1, x2, y2, dashed=False, color=ARROW_COLOR, width=1):
    if dashed:
        draw_dashed_line(draw, x1, y1, x2, y2, color, width)
    else:
        draw.line([(int(x1), int(y1)), (int(x2), int(y2))], fill=color, width=int(width))
    arrow_len = 7
    angle = math.atan2(y2 - y1, x2 - x1)
    a1 = angle + math.pi / 6
    a2 = angle - math.pi / 6
    draw.polygon(
        [
            (int(x2), int(y2)),
            (int(x2 - arrow_len * math.cos(a1)), int(y2 - arrow_len * math.sin(a1))),
            (int(x2 - arrow_len * math.cos(a2)), int(y2 - arrow_len * math.sin(a2))),
        ],
        fill=color,
    )


# ========== 核心绘制 ==========
def generate_seq_diagram(title, participants, messages, output_path, max_messages=25):
    """
    紧凑时序图绘制。
    宽度完全自适应，消息文字精确居中。

    Args:
        title: 时序图标题
        participants: 参与者名称列表
        messages: 消息列表，每条 {"from": int, "to": int, "text": str, "dashed": bool}
        output_path: 输出PNG路径
        max_messages: 最大消息条数
    """
    font_label = get_font(11, bold=True)
    font_msg = get_font(10)
    font_title = get_font(12, bold=True)

    n = len(participants)
    if len(messages) > max_messages:
        messages = messages[:max_messages]

    for msg in messages:
        text = msg.get('text', '')
        if len(text) > 28:
            msg['text'] = text[:26] + ".."

    # 测量参与者框宽度
    max_name_w = max(font_label.getlength(p) for p in participants)
    part_w = int(max_name_w) + 18
    part_w = max(part_w, 60)
    part_h = 36
    part_r = 6

    # 测量消息文字宽度，动态计算间距
    max_msg_w = 0
    for m in messages:
        text = m.get('text', '')
        if text:
            max_msg_w = max(max_msg_w, font_msg.getlength(text))

    min_gap = part_w + int(max_msg_w) + 20
    part_gap = max(min_gap, 40)

    margin_l = 15
    margin_r = 15
    total_w = margin_l + n * part_w + (n - 1) * part_gap + margin_r

    part_x = [margin_l + part_w // 2 + i * (part_w + part_gap) for i in range(n)]
    last_right = part_x[-1] + part_w // 2
    total_w = max(total_w, last_right + margin_r)

    margin_t = 45
    margin_b = 30
    msg_gap = 28
    img_h = margin_t + part_h + 16 + len(messages) * msg_gap + margin_b + part_h

    img = Image.new('RGB', (total_w, img_h), BG_COLOR)
    draw = ImageDraw.Draw(img)

    # 标题
    tw = font_title.getlength(title)
    draw.text(((total_w - tw) / 2, 10), title, fill=(26, 86, 160), font=font_title)

    def draw_part_box(cx, cy, text):
        x1, x2 = cx - part_w // 2, cx + part_w // 2
        draw_rounded_rect(draw, (x1, cy, x2, cy + part_h), part_r,
                          fill=PART_BG, outline=PART_BORDER, width=1)
        nw = font_label.getlength(text)
        draw.text((cx - nw / 2, cy + (part_h - 14) / 2), text,
                  fill=LABEL_COLOR, font=font_label)

    # 顶部参与者框
    top_y = margin_t
    for i in range(n):
        draw_part_box(part_x[i], top_y, participants[i])

    # 生命线
    ll_start = top_y + part_h
    ll_end = img_h - margin_b - part_h
    for i in range(n):
        draw_dashed_line(draw, part_x[i], ll_start, part_x[i], ll_end, LIFELINE_COLOR, 1)

    # 消息箭头
    cur_y = ll_start + 14
    for msg in messages:
        fi = min(msg.get('from', 0), n - 1)
        ti = min(msg.get('to', 0), n - 1)
        text = msg.get('text', '')
        dashed = msg.get('dashed', False)
        fx, tx = part_x[fi], part_x[ti]

        draw_arrow(draw, fx, cur_y, tx, cur_y, dashed=dashed)

        if text and abs(fx - tx) > 10:
            mid_x = (fx + tx) / 2
            mw = font_msg.getlength(text)
            draw.text((mid_x - mw / 2, cur_y - 16), text, fill=TEXT_COLOR, font=font_msg)

        cur_y += msg_gap

    # 底部参与者框
    for i in range(n):
        draw_part_box(part_x[i], ll_end, participants[i])

    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
    img.save(output_path, 'PNG', dpi=(DPI, DPI))
    return output_path


# ========== 参与者推导 ==========
def derive_participants(l2_name, l2_rows):
    """
    推导4-6个参与者。
    基于二级模块名称中的关键词匹配业务领域，分配合适的参与者集合。

    Returns:
        list[str]: 参与者名称列表
    """
    name = str(l2_name)

    # --- 管理运营类 ---
    mgmt_keywords = [
        "管理", "运营", "配置", "角色", "权限", "日志", "服务号运营",
        "报表", "指标", "敏感词", "设备安全", "群组管理", "消息管理",
        "文件管理", "快通知管理",
    ]
    if any(kw in name for kw in mgmt_keywords):
        if "角色" in name or "权限" in name:
            return ["运营人员", "管理平台", "权限服务", "数据库", "日志服务"]
        elif "日志" in name:
            return ["运营人员", "管理平台", "日志采集", "数据存储", "查询引擎"]
        elif "消息" in name or "敏感词" in name:
            return ["运营人员", "管理平台", "消息引擎", "过滤服务", "存储层"]
        elif "文件" in name:
            return ["运营人员", "管理平台", "文件服务", "存储引擎", "分发通道"]
        elif "群组" in name:
            return ["运营人员", "管理平台", "群组服务", "权限校验", "数据层"]
        elif "报表" in name or "指标" in name:
            return ["运营人员", "管理平台", "数据采集", "统计引擎", "报表服务"]
        elif "设备安全" in name:
            return ["运营人员", "管理平台", "设备注册", "安全校验", "策略中心"]
        else:
            return ["运营人员", "管理平台", "业务服务", "数据库", "缓存层"]

    # --- 客户端交互类 ---
    client_keywords = ["Android", "iOS", "鸿蒙", "WinPC", "MacPC", "PC端", "PC国产化"]
    for kw in client_keywords:
        if kw in name:
            if "文件" in name and ("限制" in name or "禁止" in name):
                return ["用户", kw + "客户端", "文件服务", "策略引擎", "缓存层"]
            elif "登录" in name or "认证" in name:
                return ["用户", kw + "客户端", "认证服务", "统一认证", "数据库"]
            elif "通讯录" in name:
                return ["用户", kw + "客户端", "通讯录服务", "组织架构", "缓存"]
            elif "群聊" in name or "群组" in name or "机器人" in name:
                return ["用户", kw + "客户端", "MQTT通道", "机器人服务", "消息网关"]
            elif "语音" in name or "TTS" in name:
                return ["用户", kw + "客户端", "TTS引擎", "音频服务", "缓存"]
            elif "智能" in name or "问答" in name:
                return ["用户", kw + "客户端", "AI引擎", "知识库", "消息队列"]
            elif "头像" in name:
                return ["用户", kw + "客户端", "头像服务", "CDN", "缓存"]
            else:
                return ["用户", kw + "客户端", "业务服务", "数据库", "缓存"]

    # --- 灵犀/助手类 ---
    lingxi_keywords = ["灵犀", "助手", "智能体", "备忘"]
    if any(kw in name for kw in lingxi_keywords):
        return ["用户", "灵犀客户端", "灵犀引擎", "AI服务", "知识库"]

    # --- 系统对接/数据同步类 ---
    sync_keywords = ["同步", "对接", "适配", "接入", "门户"]
    if any(kw in name for kw in sync_keywords):
        if "门户" in name:
            return ["运营人员", "管理平台", "门户服务", "数据服务", "权限中心"]
        elif "同步" in name:
            return ["源系统", "同步引擎", "数据清洗", "目标存储", "监控服务", "日志"]
        else:
            return ["管理平台", "适配层", "目标系统", "配置中心", "数据库"]

    # --- 默认 ---
    return ["用户", "客户端", "业务服务", "数据库"]


# ========== 消息推导 ==========
def derive_messages(l2_name, l2_rows, participants):
    """
    推导20-25条交互消息。
    从四级模块语义提取业务动作，从数据属性提取字段，按阶段展开。

    Returns:
        list[dict]: 消息列表 [{"from": int, "to": int, "text": str, "dashed": bool}]
    """
    n = len(participants)

    # 提取四级模块文本
    l4_texts = []
    for _, row in l2_rows.iterrows():
        l4 = str(row.get('四级模块', ''))
        if l4.strip() and l4.strip() != 'nan':
            l4_texts.append(l4.strip())

    # 提取数据属性字段
    da_all = []
    for _, row in l2_rows.iterrows():
        da = str(row.get('数据属性', ''))
        if da.strip() and da.strip() != 'nan':
            da_all.append(da.strip())
    da_combined = '\uff0c'.join(da_all)
    fields = [f.strip() for f in re.split(r'[\u3001\uff0c,]', da_combined) if f.strip()][:10]

    l3_names = [str(l3).strip() for l3 in l2_rows['三级模块'].dropna().unique()
                if str(l3).strip()]

    # 从L4提取动作短语
    def extract_actions(text):
        parts = re.split(r'[\uff0c,;\u3001\u5e76]', text)
        actions = []
        for part in parts:
            part = part.strip()
            for prefix in ['用户', '系统', '客户端', '管理', '运营', '管理员']:
                if part.startswith(prefix):
                    part = part[len(prefix):]
                    break
            part = part.strip('\u7684').strip()
            if len(part) >= 4:
                if len(part) > 18:
                    part = part[:16] + ".."
                actions.append(part)
        return actions

    all_actions = []
    for l4 in l4_texts:
        all_actions.extend(extract_actions(l4))
    seen = set()
    unique_actions = []
    for a in all_actions:
        if a not in seen:
            unique_actions.append(a)
            seen.add(a)

    def short_field(f):
        return f if len(f) <= 14 else f[:12] + ".."

    messages = []
    last_data_idx = n - 1
    second_last = n - 2 if n >= 4 else n - 1
    d_idx = min(3, n - 1)

    # 阶段1: 请求发起
    act = unique_actions[0] if unique_actions else (l3_names[0] if l3_names else "\u53d1\u8d77\u8bf7\u6c42")
    messages.append({"from": 0, "to": 1, "text": act[:16]})
    if fields:
        messages.append({"from": 0, "to": 1, "text": f"\u4f20\u5165{short_field(fields[0])}", "dashed": True})

    # 阶段2: 权限校验
    act2 = unique_actions[1] if len(unique_actions) > 1 else "\u8eab\u4efd\u6821\u9a8c"
    messages.append({"from": 1, "to": 2, "text": act2[:16]})
    if fields and len(fields) > 1:
        messages.append({"from": 2, "to": second_last, "text": f"\u9a8c\u8bc1{short_field(fields[1])}"})
    if n >= 5:
        messages.append({"from": second_last, "to": last_data_idx, "text": "\u67e5\u8be2\u6743\u9650"})
        messages.append({"from": last_data_idx, "to": second_last, "text": "\u6743\u9650\u6709\u6548", "dashed": True})
    messages.append({"from": 2, "to": 1, "text": "\u6821\u9a8c\u901a\u8fc7", "dashed": True})

    # 阶段3: 核心处理
    act3 = unique_actions[2] if len(unique_actions) > 2 else "\u4e1a\u52a1\u5904\u7406"
    messages.append({"from": 1, "to": 2, "text": act3[:16]})
    if n >= 4:
        messages.append({"from": 2, "to": 3, "text": "\u903b\u8f91\u6821\u9a8c"})
    if n >= 5 and fields and len(fields) > 2:
        messages.append({"from": 3, "to": 4, "text": f"\u5904\u7406{short_field(fields[2])}"})

    # 阶段4: 数据操作
    act4 = unique_actions[3] if len(unique_actions) > 3 else "\u6570\u636e\u64cd\u4f5c"
    messages.append({"from": 2, "to": d_idx, "text": act4[:16]})
    if fields and len(fields) > 3:
        messages.append({"from": d_idx, "to": min(d_idx + 1, n - 1), "text": f"\u67e5\u8be2{short_field(fields[3])}"})
    messages.append({"from": d_idx, "to": 2, "text": "\u8fd4\u56de\u6570\u636e\u96c6", "dashed": True})

    # 阶段5: 结果组装+缓存
    act5 = unique_actions[4] if len(unique_actions) > 4 else "\u7ed3\u679c\u7ec4\u88c5"
    messages.append({"from": 2, "to": 1, "text": act5[:16], "dashed": True})
    if n >= 5:
        messages.append({"from": 2, "to": last_data_idx, "text": "\u66f4\u65b0\u7f13\u5b58"})
        messages.append({"from": last_data_idx, "to": 2, "text": "\u7f13\u5b58\u5df2\u66f4\u65b0", "dashed": True})

    # 阶段6: 返回客户端
    act6 = unique_actions[5] if len(unique_actions) > 5 else "\u8fd4\u56de\u7ed3\u679c"
    messages.append({"from": 1, "to": 0, "text": act6[:16], "dashed": True})

    # 阶段7: 用剩余字段补充
    remaining = fields[4:]
    for f in remaining[:2]:
        if len(messages) >= 25:
            break
        messages.append({"from": 1, "to": 2, "text": f"\u5904\u7406{short_field(f)}"})
        if n >= 4:
            messages.append({"from": 2, "to": d_idx, "text": f"\u6821\u9a8c{short_field(f)}"})
            messages.append({"from": d_idx, "to": 2, "text": "\u6821\u9a8c\u901a\u8fc7", "dashed": True})

    # 补充到至少20条
    while len(messages) < 20:
        extras_4 = [
            {"from": 1, "to": 2, "text": "\u683c\u5f0f\u5316\u6570\u636e"},
            {"from": 2, "to": d_idx, "text": "\u5199\u5165\u8bb0\u5f55"},
            {"from": d_idx, "to": 2, "text": "\u5199\u5165\u5b8c\u6210", "dashed": True},
            {"from": 2, "to": 1, "text": "\u5904\u7406\u5b8c\u6bd5", "dashed": True},
            {"from": 1, "to": 0, "text": "\u72b6\u6001\u5237\u65b0", "dashed": True},
            {"from": 2, "to": last_data_idx, "text": "\u8bb0\u5f55\u65e5\u5fd7"},
        ]
        extras_5 = [
            {"from": 2, "to": 3, "text": "\u6570\u636e\u6c47\u603b"},
            {"from": 3, "to": 4, "text": "\u540c\u6b65\u66f4\u65b0"},
            {"from": 4, "to": 3, "text": "\u540c\u6b65\u5b8c\u6210", "dashed": True},
            {"from": 3, "to": 2, "text": "\u6c47\u603b\u7ed3\u679c", "dashed": True},
            {"from": 2, "to": 1, "text": "\u63a8\u9001\u901a\u77e5", "dashed": True},
            {"from": 1, "to": 0, "text": "\u5c55\u793a\u66f4\u65b0", "dashed": True},
        ]
        pool = extras_5 if n >= 5 else extras_4
        added = False
        for ex in pool:
            if len(messages) >= 20:
                break
            key = f"{ex['from']}_{ex['to']}_{ex['text']}"
            if not any(f"{m['from']}_{m['to']}_{m['text']}" == key for m in messages):
                messages.append(ex)
                added = True
        if not added:
            break

    return messages[:25]


# ========== 描述生成 ==========
def generate_description(l2_name, l2_rows, participants, messages):
    """
    产品经理视角的精简流程描述。
    从消息中按阶段归纳，每阶段一句话。
    格式：["xxx", "xxx", ...]

    Returns:
        list[str]: 描述列表
    """
    n = len(participants)
    user_role = participants[0]
    client_role = participants[1]
    desc_items = []

    # 阶段1: 用户操作
    for msg in messages:
        if msg['from'] == 0 and msg['to'] == 1 and not msg.get('dashed'):
            desc_items.append(f"{user_role}{msg['text']}")
            break

    # 阶段2: 请求转发
    for msg in messages:
        if msg['from'] == 1 and msg['to']  == 2 and not msg.get('dashed'):
            desc_items.append(f"{client_role}\u5c06\u8bf7\u6c42\u8f6c\u53d1\u81f3{participants[2]}")
            break

    # 阶段3: 参数校验
    validated = False
    for msg in messages:
        if '\u6821\u9a8c\u901a\u8fc7' in msg['text'] and msg.get('dashed'):
            desc_items.append(f"{participants[2]}\u5b8c\u6210\u53c2\u6570\u6821\u9a8c\u4e0e\u6743\u9650\u9a8c\u8bc1")
            validated = True
            break
    if not validated:
        desc_items.append(f"{participants[2]}\u6267\u884c\u53c2\u6570\u6821\u9a8c\u4e0e\u6743\u9650\u9a8c\u8bc1")

    # 阶段4: 核心业务处理
    if n >= 4:
        desc_items.append(f"{participants[2]}\u8c03\u7528{participants[3]}\u6267\u884c\u6838\u5fc3\u4e1a\u52a1\u903b\u8f91")

    # 阶段5: 数据操作
    data_ops = [m for m in messages if m.get('dashed') and '\u8fd4\u56de' in m['text'] and m['from'] >= 3]
    if data_ops:
        desc_items.append(f"{participants[min(3, n - 1)]}\u5b8c\u6210\u6570\u636e\u67e5\u8be2\u4e0e\u6301\u4e45\u5316\u64cd\u4f5c")

    # 阶段6: 结果组装与缓存
    has_cache = any('\u7f13\u5b58' in msg['text'] for msg in messages)
    if has_cache:
        desc_items.append(f"{participants[2]}\u7ec4\u88c5\u54cd\u5e94\u6570\u636e\u5e76\u66f4\u65b0\u7f13\u5b58")
    else:
        desc_items.append(f"{participants[2]}\u7ec4\u88c5\u54cd\u5e94\u6570\u636e")

    # 阶段7: 返回结果
    for msg in messages:
        if msg['from'] == 1 and msg['to'] == 0 and msg.get('dashed') and '\u7ed3\u679c' in msg['text']:
            desc_items.append(f"{client_role}\u5411{user_role}\u8fd4\u56de\u5904\u7406\u7ed3\u679c")
            break
    else:
        for msg in messages:
            if msg['from'] == 1 and msg['to'] == 0 and msg.get('dashed'):
                desc_items.append(f"{client_role}\u5411{user_role}\u8fd4\u56de\u6267\u884c\u7ed3\u679c")
                break

    # 阶段8: 后续操作
    has_log = any('\u65e5\u5fd7' in msg['text'] for msg in messages)
    has_sync = any('\u540c\u6b65' in msg['text'] for msg in messages)
    if has_log and has_sync:
        desc_items.append(f"{participants[2]}\u8bb0\u5f55\u64cd\u4f5c\u65e5\u5fd7\u5e76\u540c\u6b65\u72b6\u6001\u53d8\u66f4")
    elif has_log:
        desc_items.append(f"{participants[2]}\u8bb0\u5f55\u64cd\u4f5c\u65e5\u5fd7")
    elif has_sync:
        desc_items.append(f"{participants[2]}\u540c\u6b65\u72b6\u6001\u53d8\u66f4\u81f3\u4e0b\u6e38\u670d\u52a1")

    return desc_items
