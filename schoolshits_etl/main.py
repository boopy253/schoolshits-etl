import argparse
import re
import time
from copy import copy

import pandas as pd
from loguru import logger
from openpyxl import load_workbook


# 参数解析
def parse_args():
    parser = argparse.ArgumentParser(description="根据 source Excel 填充 target 模板（支持多种 source 格式）")

    parser.add_argument("--source", default="source.xlsx")
    parser.add_argument("--template", default="target.xlsx")
    parser.add_argument("--output", default="target_filled.xlsx")
    parser.add_argument("--side-output", default="version_grade_book.xlsx")
    parser.add_argument("--school", default="示例小学")
    parser.add_argument("--year", default="2025")
    parser.add_argument("--start-row", type=int, default=4)

    return parser.parse_args()


def normalize_text(text):
    if not isinstance(text, str):
        return ""
    return text.replace("（", "(").replace("）", ")").strip()


def copy_row_style(ws, src_row, tgt_row, max_col):
    for col in range(1, max_col + 1):
        src = ws.cell(row=src_row, column=col)
        tgt = ws.cell(row=tgt_row, column=col)
        tgt.value = src.value
        if src.has_style:
            tgt._style = copy(src._style)


# Source 兼容 + 归一化
def load_and_normalize_source(path):
    raw = pd.read_excel(path, header=None)
    first_cell = str(raw.iloc[0, 0])

    if "发货单明细" in first_cell:
        logger.info("识别为【发货单明细】格式（header 在第 2 行）")
        df = pd.read_excel(path, header=1)

        def split_row(r):
            free = str(r.get("是否免费")).strip() == "是"
            return pd.Series(
                {
                    "书名": r.get("产品名称"),
                    "版别": None,
                    "单价": r.get("定价"),
                    "非免费订数": 0 if free else r.get("发货数"),
                    "免费订数": r.get("发货数") if free else 0,
                }
            )

        df2 = df.apply(split_row, axis=1)

    elif "教辅目录" in first_cell:
        logger.info("识别为【小学教辅目录】格式（header 在第 2 行）")
        df = pd.read_excel(path, header=1)
        student_col = df.columns[-2]
        teacher_col = df.columns[-1]

        df2 = pd.DataFrame(
            {
                "书名": df["名称"],
                "版别": df["版本"],
                "单价": df["单价"],
                "非免费订数": df[student_col],
                "免费订数": df[teacher_col],
            }
        )

    elif "免费教材" in first_cell:
        logger.info("识别为【免费教材征订目录】格式（header 在第 3 行）")
        df = pd.read_excel(path, header=2)

        def find_col_by_keywords(df, keywords):
            for col in df.columns:
                col_str = str(col).replace(" ", "").replace("\u3000", "")
                if all(k in col_str for k in keywords):
                    return col
            return None

        name_col = find_col_by_keywords(df, ["书", "名"])
        version_col = find_col_by_keywords(df, ["版"])

        df2 = pd.DataFrame(
            {
                "书名": df[name_col],
                "版别": df[version_col] if version_col else None,
                "单价": 0,
                "非免费订数": 0,
                "免费订数": df.iloc[:, -1],
            }
        )

    else:
        logger.info("识别为【订数统计表】格式")
        df = pd.read_excel(path)
        df2 = df[["书名", "版别", "单价", "非免费订数", "免费订数"]].copy()

    for col in ["非免费订数", "免费订数"]:
        df2[col] = pd.to_numeric(df2[col], errors="coerce").fillna(0).astype(int)  # type: ignore

    df2["单价"] = pd.to_numeric(df2["单价"], errors="coerce").fillna(0)  # type: ignore

    return df2


# 业务解析逻辑
def parse_grade_and_term(book_name):
    if not isinstance(book_name, str):
        return None, None

    text = normalize_text(book_name)

    grade_map = {
        "一": "一年级",
        "二": "二年级",
        "三": "三年级",
        "四": "四年级",
        "五": "五年级",
        "六": "六年级",
        "1": "一年级",
        "2": "二年级",
        "3": "三年级",
        "4": "四年级",
        "5": "五年级",
        "6": "六年级",
    }

    grade = None
    term = None

    m = re.search(r"((一|二|三|四|五|六)|([1-6]))年级", text)
    if m:
        grade = grade_map[m.group(1)]
    else:
        m = re.search(r"([一二三四五六])(?=[上下])", text)
        if m:
            grade = grade_map[m.group(1)]

    if re.search(r"上册?|\(上\)", text):
        term = "上学期"
    elif re.search(r"下册?|\(下\)", text):
        term = "下学期"

    return grade, term


# 科目识别
def parse_subject(book_name):
    if not isinstance(book_name, str):
        return None

    subjects = [
        "道德与法治",
        "语文",
        "数学",
        "英语",
        "科学",
        "书法",
        "品德与社会",
        "音乐",
        "美术",
        "信息技术",
        "综合实践",
        "体育与健康",
        "体育",
    ]

    for s in subjects:
        if s in book_name:
            return "体育" if s in ("体育", "体育与健康") else s

    return None


def parse_category(book_name):
    if not isinstance(book_name, str):
        return None
    for k in [
        "教参",
        "教师用书",
        "学生活动手册",
        "练习",
        "光盘",
        "学具",
        "字典",
        "教案",
        "辅导",
        "同步",
        "导学",
        "训练",
    ]:
        if k in book_name:
            return "教辅"
    return "教材"


# 版本识别
def parse_version(book_name, fallback_version):
    if not isinstance(book_name, str):
        return fallback_version

    text = book_name.replace("（", "(").replace("）", ")")

    explicit_patterns = [
        r"北师大版",
        r"北师大",
        r"人教版",
        r"粤教版",
        r"人音版",
        r"人美版",
        r"语文S版",
        r"语文社S版",
    ]
    for p in explicit_patterns:
        m = re.search(p, text)
        if m:
            v = m.group(0)
            return "北师大版" if v == "北师大" else v

    m = re.search(r"配\s*([^\)（）]+)", text)
    if m:
        v = m.group(1).strip()
        return "北师大版" if "北师大" in v else v

    for pub in ["湖南美术", "江苏教育", "粤教科技", "粤教育", "粤高教"]:
        if pub in text:
            return pub

    return fallback_version


# 主程序
def main():
    args = parse_args()
    start_time = time.time()

    logger.info(f"读取 Source：{args.source}")
    df_src = load_and_normalize_source(args.source)

    # 年级排序
    df_src["_grade_order"] = df_src["书名"].apply(  # type: ignore
        lambda x: {
            "一年级": 1,
            "二年级": 2,
            "三年级": 3,
            "四年级": 4,
            "五年级": 5,
            "六年级": 6,
        }.get(parse_grade_and_term(x)[0] or "", 99)
    )

    df_src = df_src.sort_values(by="_grade_order").drop(columns="_grade_order")  # type: ignore

    # 主输出
    wb = load_workbook(args.template)
    ws = wb.active

    if ws is None:
        raise ValueError("Template Excel file has no active worksheet.")

    max_col = ws.max_column
    current_row = args.start_row

    for idx, (_, r) in enumerate(df_src.iterrows()):
        if current_row > ws.max_row:
            copy_row_style(ws, args.start_row, current_row, max_col)

        ws.cell(row=current_row, column=1, value=idx + 1)

        book = r["书名"]
        grade, term = parse_grade_and_term(book)

        values = [
            args.year,
            term,
            grade,
            parse_subject(book),
            parse_category(book),
            book,
            r["版别"],
            r["非免费订数"] + r["免费订数"],
            r["单价"],
        ]

        for col, value in enumerate(values, start=3):
            ws.cell(row=current_row, column=col, value=value)

        current_row += 1

    wb.save(args.output)
    logger.info(f"主输出完成：{args.output}")

    # 副输出
    logger.info("生成副输出文件")

    side_rows = []
    for _, r in df_src.iterrows():
        book = r["书名"]
        grade, _ = parse_grade_and_term(book)

        side_rows.append(
            {
                "版本": parse_version(book, r["版别"]),
                "年级": grade,
                "书名": book,
                "单价": r["单价"],
                "数量": r["非免费订数"] + r["免费订数"],
                "类别": parse_category(book),
                "科目": parse_subject(book),
            }
        )

    pd.DataFrame(side_rows).to_excel(args.side_output, index=False)
    logger.info(f"副输出完成：{args.side_output}")

    logger.info(f"全部完成 ✅ 用时 {time.time() - start_time:.2f} 秒")
