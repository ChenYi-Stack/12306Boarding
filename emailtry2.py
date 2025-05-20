import imaplib
import email
from email.header import decode_header
from email.utils import parseaddr
import pandas as pd
import os
import re
from openpyxl.styles import PatternFill
from datetime import datetime

# ================= 邮箱配置区 =================
EMAIL = "EMAIL" #YOUR EMAIL
PASSWORD = "PASSWORD"  #QQ邮箱需授权码
#QQ邮箱登录页面->右上角“账号与安全”->安全设置->(下拉)生成授权码->生成。
#温馨提示：收取选项建议收取“全部”的邮件，默认为15日。
IMAP_SERVER = "imap.qq.com"  # QQ邮箱IMAP服务器地址

YOUR_NAME = "NAME"  # 需要筛选的乘客姓名


# =============================================

def connect_to_mailbox():
    """连接到邮箱服务器"""
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, PASSWORD)
        mail.select("inbox")
        return mail
    except Exception as e:
        print(f"连接或登录失败: {e}")
        exit(1)

def extract_12306_data(msg, mail_content):
    """从邮件内容提取12306相关数据（最终优化版）"""
    patterns = {
        "购票日期": r"您于(\d{4}年\d{1,2}月\d{1,2}日)在中国铁路客户服务中心网站",
        "发车日期": r"(\d{4}年\d{1,2}月\d{1,2}日)\d{1,2}:\d{2}开",
        "发车时间": r"\d{4}年\d{1,2}月\d{1,2}日(\d{1,2}:\d{2})开",
        "车次": r"([A-Z]?\d{1,4})(?:次|次列车|车次|$)",
        "出发站": r"(\w+站)[—\-]",
        "到达站": r"[—\-](\w+站)",
        # 改进车厢号匹配：确保无座情况能获取完整车厢号
        "车厢号": r"(\d+[A-Z]?车)(?:无座|\d+[A-Z]?号)",  # 修改正则表达式结构
        "座位号": r"\d+[A-Z]?车(\d+[A-Z]?号|无座)",
        # 强化姓名匹配：使用全名精确匹配
        "乘客姓名": rf"({re.escape(YOUR_NAME)})，\d{{4}}年\d{{1,2}}月\d{{1,2}}日",
        "座位等级": r"(?:^|[\s，,])([一二]等座|商务座|特等座|硬[卧座]|软[卧座]|硬卧[上中下]铺|软卧[上下]铺|无座)(?:$|[\s，。])",
        "票价": r"(?:票价|票款|金额|应付金额)(?:\s*[:：]|\s+)?(\d+\.?\d{0,2})元",
        "订单号": r"订单号码\s*([A-Z0-9]{8,})"
    }

    data = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, mail_content, re.IGNORECASE)
        data[key] = match.group(1).upper() if match and match.group(1) else None

    # 强制统一格式
    data = {k: v.upper() if isinstance(v, str) and k in ["车次", "车厢号", "座位号"] else v
            for k, v in data.items()}

    # 改进点1：处理无座信息（确保车厢号正确获取）
    if data["车厢号"] and data["座位号"]:
        # 移除多余的车字（正则已包含"车"字）
        clean_carriage = data['车厢号'].rstrip('车')  # 示例："17车" → "17"
        data["座位"] = f"{clean_carriage}车{data['座位号']}"
    elif data.get("座位等级") == "无座":
        # 当存在无座标识但无车厢号时，尝试从座位字段提取
        if not data["车厢号"] and '车' in mail_content:
            carriage_match = re.search(r'(\d+车)无座', mail_content)
            if carriage_match:
                data["车厢号"] = carriage_match.group(1).strip()
        data["座位"] = f"{data['车厢号']}无座" if data["车厢号"] else "无座"
        data["座位号"] = "无座"
    else:
        data["座位"] = None

    # 处理票价格式
    if data["票价"]:
        try:
            data["票价"] = f"{round(float(data['票价']), 1):.1f}"
        except:
            data["票价"] = None

    # 改进点2：严格验证乘客身份
    data["本人车票"] = "是" if (data.get("乘客姓名") and data["乘客姓名"].strip() == YOUR_NAME) else "否"

    # 添加元数据
    # 状态判断逻辑保留
    data["状态"] = "已退" if re.search(r"退票成功|已退票|退单", mail_content) else "有效"

    # 购票日期使用正则提取结果
    if not data.get("购票日期"):
        # 添加兜底日期格式匹配（针对不同邮件模板）
        fallback_date = re.search(r"订单生成时间[:：]\s*(\d{4}-\d{2}-\d{2})", mail_content)
        data["购票日期"] = fallback_date.group(1).replace("-", "年", 1).replace("-",
                                                                                "月") + "日" if fallback_date else None

    return data


def process_subject(subject):
    """处理邮件主题（增加退单过滤）"""
    if "候补订单退单通知" in subject:
        return None

    prefixes = ["网上购票系统-", "列车"]
    for prefix in prefixes:
        if subject.startswith(prefix):
            return subject[len(prefix):]
    return subject


def parse_email(msg):
    """解析邮件内容"""
    content = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            charset = part.get_content_charset() or 'utf-8'
            try:
                payload = part.get_payload(decode=True)
                if not payload:
                    continue

                decoded = payload.decode(charset, errors='replace')

                if content_type == "text/plain":
                    content = decoded
                    break
                elif content_type == "text/html":
                    html_cleaner = re.compile(r'<[^>]+>|[\s]{2,}|&nbsp;|\\n|\\r|\\t')
                    content = re.sub(html_cleaner, ' ', decoded)
                    content = re.sub(r'[【】（）()]', ' ', content)
            except Exception as e:
                print(f"解码邮件内容失败: {e}")
    else:
        charset = msg.get_content_charset() or 'utf-8'
        try:
            payload = msg.get_payload(decode=True)
            if payload:
                content = payload.decode(charset, errors='replace')
        except Exception as e:
            print(f"解码邮件内容失败: {e}")

    return extract_12306_data(msg, content)


def fetch_all_12306_uids(mail):
    """获取所有历史邮件的UID"""
    try:
        status, data = mail.uid('search', None, '(FROM "12306@rails.com.cn")')
        if status != "OK":
            print("邮件搜索失败")
            return []

        uids = data[0].split()
        print(f"找到 {len(uids)} 封潜在12306邮件")
        return uids
    except Exception as e:
        print(f"邮件搜索失败: {e}")
        return []


def fetch_email_batch(mail, uids, batch_size=100):
    """分批次获取邮件内容"""
    valid_emails = []
    total = len(uids)

    for i in range(0, total, batch_size):
        batch = uids[i:i + batch_size]
        decoded_batch = [uid.decode('utf-8') if isinstance(uid, bytes) else str(uid) for uid in batch]
        uid_str = ",".join(decoded_batch)

        status, data = mail.uid('fetch', uid_str, '(RFC822)')
        if status != "OK":
            continue

        for item in data:
            if isinstance(item, tuple):
                raw_email = item[1]
                msg = email.message_from_bytes(raw_email)
                from_email = parseaddr(msg["From"])[1]
                if from_email == "12306@rails.com.cn":
                    valid_emails.append(msg)

        progress = min(i + batch_size, total) / total * 100
        print(f"加载进度：{progress:.1f}% ({min(i + batch_size, total)}/{total})")

    print(f"有效12306邮件数量：{len(valid_emails)}")
    return valid_emails


def save_to_excel(df, output_path):
    """保存数据到Excel（含样式设置）"""
    # 列顺序调整
    columns_order = [
        "发车日期", "发车时间", "出发站", "到达站", "车次",
        "座位", "车厢号", "座位号", "座位等级", "票价",
        "主题", "订单号", "购票日期", "状态", "订单来源"
    ]

    # 日期格式转换
    df['发车日期'] = pd.to_datetime(
        df['发车日期'].str.replace('年', '-').str.replace('月', '-').str.replace('日', ''),
        format='%Y-%m-%d'
    )
    df['发车时间'] = pd.to_datetime(df['发车时间'], format='%H:%M').dt.strftime('%H:%M')

    # 排序处理
    df = df.sort_values(by=['发车日期', '发车时间'])
    df['发车日期'] = df['发车日期'].dt.strftime('%Y年%m月%d日')

    # 改进点3：过滤非本人车票（新增过滤条件）
    df = df[df['本人车票'] == '是']

    # 改进点4：注释退票过滤逻辑（原问题3需求）
    ################################################################
    # 原退票过滤逻辑已注释，可根据需要恢复
    # canceled_orders = df[df['状态'] == '已退']['订单号'].dropna().unique()
    # final_df = df[~df['订单号'].isin(canceled_orders)]
    # canceled_pairs = df[df['状态'] == '已退'][['发车日期', '车次']].drop_duplicates()
    # final_df = final_df.merge(canceled_pairs, how='left', indicator=True)
    # final_df = final_df[final_df['_merge'] == 'left_only'].drop(columns=['_merge'])
    ################################################################
    final_df = df  # 直接使用未过滤的原始数据

    # 保存文件
    final_df[columns_order].to_excel(output_path, index=False)
    print(f"数据已保存至：{output_path}")


def main():
    mail = connect_to_mailbox()
    uids = fetch_all_12306_uids(mail)

    if not uids:
        print("没有需要处理的邮件")
        return

    print("\n开始加载邮件内容...")
    all_emails = fetch_email_batch(mail, uids)

    print("\n开始解析邮件内容...")
    results = []
    for idx, msg in enumerate(all_emails):
        try:
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode(decode_header(msg["Subject"])[0][1] or 'utf-8')

            processed_subject = process_subject(subject)
            if not processed_subject:
                continue

            mail_content = parse_email(msg)
            mail_content.update({
                "主题": processed_subject,
                "订单来源": "12306",
                "状态": "已退" if any(kw in processed_subject for kw in ["退票", "退单"]) else "有效"
            })

            if any(mail_content.values()):
                results.append(mail_content)

            if (idx + 1) % 10 == 0:
                print(f"解析进度：{idx + 1}/{len(all_emails)}")

        except Exception as e:
            print(f"处理第 {idx + 1} 封邮件出错: {e}")

    if not results:
        print("未提取到有效数据！")
        return

    df = pd.DataFrame(results)
    output_path = os.path.join(os.path.expanduser("~"), "Desktop", "12306车票统计.xlsx")
    save_to_excel(df, output_path)


if __name__ == "__main__":
    main()