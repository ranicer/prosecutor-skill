#!/usr/bin/env python3
"""
检察文书格式化脚本
使用 python-docx 生成符合公文格式的 Word 检察文书

公文格式要求：
- 字体：方正仿宋_GBK（正文三号），方正小标宋_GBK（标题二号），方正楷体_GBK（引用三号）
- 行距：固定值28磅
- 页边距：上37mm，下35mm，左28mm，右26mm
- 页码：居中，四号半角阿拉伯数字
"""

import argparse
import json
import sys
from pathlib import Path

try:
    from docx import Document
    from docx.shared import Pt, Cm, Mm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
except ImportError:
    print("错误：需要安装 python-docx 库")
    print("请运行：pip install python-docx")
    sys.exit(1)


def set_font(run, font_name="方正仿宋_GBK", font_size=Pt(16), bold=False):
    """设置字体"""
    run.font.size = font_size
    run.font.bold = bold
    run.font.name = font_name
    # 设置中文字体
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = r.makeelement(qn("w:rFonts"), {})
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), font_name)


def set_paragraph_format(paragraph, line_spacing=Pt(28), alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                         first_line_indent=None, space_before=Pt(0), space_after=Pt(0)):
    """设置段落格式"""
    pf = paragraph.paragraph_format
    pf.line_spacing = line_spacing
    pf.alignment = alignment
    pf.space_before = space_before
    pf.space_after = space_after
    if first_line_indent is not None:
        pf.first_line_indent = first_line_indent


def set_page_margins(document, top=Mm(37), bottom=Mm(35), left=Mm(28), right=Mm(26)):
    """设置页边距"""
    section = document.sections[0]
    section.top_margin = top
    section.bottom_margin = bottom
    section.left_margin = left
    section.right_margin = right


def add_title(document, text, font_name="方正小标宋_GBK", font_size=Pt(22)):
    """添加标题（居中，二号字）"""
    paragraph = document.add_paragraph()
    set_paragraph_format(paragraph, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                         first_line_indent=None)
    run = paragraph.add_run(text)
    set_font(run, font_name=font_name, font_size=font_size, bold=False)
    return paragraph


def add_heading_text(document, text, font_name="方正黑体_GBK", font_size=Pt(16)):
    """添加小标题（三号黑体）"""
    paragraph = document.add_paragraph()
    set_paragraph_format(paragraph, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                         first_line_indent=Cm(0.74))  # 首行缩进两字符
    run = paragraph.add_run(text)
    set_font(run, font_name=font_name, font_size=font_size, bold=True)
    return paragraph


def add_body_text(document, text, font_name="方正仿宋_GBK", font_size=Pt(16),
                  indent=True, bold=False):
    """添加正文段落（三号仿宋，首行缩进两字符）"""
    paragraph = document.add_paragraph()
    indent_val = Cm(0.74) if indent else None  # 三号字约0.74cm缩进两字符
    set_paragraph_format(paragraph, first_line_indent=indent_val)
    run = paragraph.add_run(text)
    set_font(run, font_name=font_name, font_size=font_size, bold=bold)
    return paragraph


def add_reference_text(document, text, font_name="方正楷体_GBK", font_size=Pt(16)):
    """添加引用文字（三号楷体）"""
    paragraph = document.add_paragraph()
    set_paragraph_format(paragraph, first_line_indent=Cm(0.74))
    run = paragraph.add_run(text)
    set_font(run, font_name=font_name, font_size=font_size)
    return paragraph


def add_empty_line(document, count=1):
    """添加空行"""
    for _ in range(count):
        document.add_paragraph()


def create_qishu(data):
    """
    生成起诉书

    data: dict, 包含以下字段：
    - procuratorate: 机关名称（如"北京市海淀区人民检察院"）
    - case_number: 文号（如"京海检刑诉〔2024〕123号"）
    - defendant: 被告人信息 dict (name, gender, birth_date, id_number, ethnicity,
                                    education, household, address, occupation)
    - arrest_date: 被拘留日期
    - arrest_approval_date: 批准逮捕日期
    - arrest_execution_date: 执行逮捕日期
    - detention_place: 羁押场所
    - facts: 案件事实 (str)
    - evidence: 证据列表 (list of str)
    - legal_basis: 法律依据 (str)
    - sentencing_circumstances: 量刑情节 (str)
    - plea_agreement: 认罪认罚情况 (str or None)
    - sentencing_recommendation: 量刑建议 (str)
    """
    doc = Document()
    set_page_margins(doc)

    # 机关名称
    add_title(doc, data.get("procuratorate", "×××人民检察院"),
              font_name="方正小标宋_GBK", font_size=Pt(22))

    # 文书名称
    add_title(doc, "起 诉 书",
              font_name="方正小标宋_GBK", font_size=Pt(22))

    # 文号
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    run = p.add_run(data.get("case_number", "×检×刑诉〔20××〕×号"))
    set_font(run, font_size=Pt(16))

    add_empty_line(doc)

    # 被告人基本情况
    add_heading_text(doc, "一、被告人基本情况")
    d = data.get("defendant", {})
    gender = d.get("gender", "男/女")
    defendant_info = (
        f"被告人{d.get('name', '×××')}，{gender}，"
        f"{d.get('birth_date', '×年×月×日')}出生，"
        f"公民身份号码{d.get('id_number', '×××')}，"
        f"{d.get('ethnicity', '×')}族，"
        f"文化程度{d.get('education', '×××')}，"
        f"户籍所在地{d.get('household', '×××')}，"
        f"现住{d.get('address', '×××')}，"
        f"职业{d.get('occupation', '×××')}。"
    )
    if data.get("arrest_date"):
        defendant_info += f"因涉嫌×××罪，于{data['arrest_date']}被×××公安局刑事拘留，"
    if data.get("arrest_approval_date"):
        defendant_info += f"于{data['arrest_approval_date']}经本院批准逮捕，"
    if data.get("arrest_execution_date"):
        defendant_info += f"于{data['arrest_execution_date']}由×××公安局执行逮捕，"
    defendant_info += f"现羁押于{data.get('detention_place', '×××看守所')}。"
    add_body_text(doc, defendant_info)

    add_empty_line(doc)

    # 案件事实
    add_heading_text(doc, "二、案件事实")
    add_body_text(doc, "经依法审查查明：")
    add_body_text(doc, data.get("facts", "（案件事实）"))

    add_empty_line(doc)

    # 证据
    add_heading_text(doc, "认定上述事实的证据如下：")
    for i, evidence in enumerate(data.get("evidence", []), 1):
        add_body_text(doc, f"{i}. {evidence}")

    add_empty_line(doc)

    # 起诉理由
    add_heading_text(doc, "三、起诉理由和法律依据")
    legal_basis = data.get("legal_basis", "×××罪")
    add_body_text(
        doc,
        f"本院认为，被告人×××的行为触犯了{legal_basis}，"
        f"犯罪事实清楚，证据确实、充分，应当以×××罪追究其刑事责任。"
    )

    # 量刑情节
    if data.get("sentencing_circumstances"):
        add_body_text(doc, "量刑情节分析：")
        add_body_text(doc, data["sentencing_circumstances"])

    # 认罪认罚
    if data.get("plea_agreement"):
        add_body_text(doc, f"认罪认罚情况：{data['plea_agreement']}")

    # 量刑建议
    add_body_text(doc, f"量刑建议：{data.get('sentencing_recommendation', '（量刑建议）')}")

    add_body_text(doc, "根据《中华人民共和国刑事诉讼法》第一百七十六条的规定提起公诉。")

    add_empty_line(doc, 2)

    # 此致
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                         first_line_indent=Cm(0.74))
    run = p.add_run("此致")
    set_font(run)

    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                         first_line_indent=Cm(0.74))
    run = p.add_run("×××人民法院")
    set_font(run)

    add_empty_line(doc, 2)

    # 签名
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.RIGHT)
    run = p.add_run("检察员：×××")
    set_font(run)

    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.RIGHT)
    run = p.add_run(data.get("date", "20××年×月×日"))
    set_font(run)

    return doc


def create_buqisu(data):
    """
    生成不起诉决定书

    data: dict, 包含以下字段：
    - procuratorate: 机关名称
    - case_number: 文号
    - unprosecuted_person: 被不起诉人信息 (同 defendant)
    - facts: 案件事实
    - unprosecution_type: 不起诉类型 ("法定" / "酌定" / "存疑" / "附条件")
    - reasons: 不起诉理由
    - legal_basis: 法律依据
    - victim_info: 被害人信息（用于告知申诉权利）
    """
    doc = Document()
    set_page_margins(doc)

    add_title(doc, data.get("procuratorate", "×××人民检察院"))
    add_title(doc, "不 起 诉 决 定 书")

    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    run = p.add_run(data.get("case_number", "×检×刑不诉〔20××〕×号"))
    set_font(run)

    add_empty_line(doc)

    # 被不起诉人基本情况
    add_heading_text(doc, "一、被不起诉人基本情况")
    d = data.get("unprosecuted_person", {})
    gender = d.get("gender", "男/女")
    info = (
        f"被不起诉人{d.get('name', '×××')}，{gender}，"
        f"{d.get('birth_date', '×年×月×日')}出生，"
        f"公民身份号码{d.get('id_number', '×××')}，"
        f"{d.get('ethnicity', '×')}族，"
        f"文化程度{d.get('education', '×××')}，"
        f"户籍所在地{d.get('household', '×××')}，"
        f"现住{d.get('address', '×××')}，"
        f"职业{d.get('occupation', '×××')}。"
    )
    add_body_text(doc, info)

    add_empty_line(doc)

    # 案件事实
    add_heading_text(doc, "二、案件事实")
    add_body_text(doc, data.get("facts", "（案件事实）"))

    add_empty_line(doc)

    # 不起诉理由
    add_heading_text(doc, "三、不起诉理由和法律依据")
    unprosecution_type = data.get("unprosecution_type", "酌定")

    if unprosecution_type == "法定":
        add_body_text(doc,
            "本院认为，被不起诉人没有犯罪事实/具有《中华人民共和国刑事诉讼法》"
            "第十六条规定的情形之一。")
        add_body_text(doc,
            "根据《中华人民共和国刑事诉讼法》第一百七十七条第一款的规定，"
            "决定对被不起诉人不起诉。")
    elif unprosecution_type == "酌定":
        add_body_text(doc,
            f"本院认为，被不起诉人的行为触犯了{data.get('legal_basis', '×××')}"
            "，构成×××罪。但鉴于犯罪情节轻微，"
            f"{data.get('reasons', '不需要判处刑罚')}。")
        add_body_text(doc,
            "根据《中华人民共和国刑事诉讼法》第一百七十七条第二款的规定，"
            "决定对被不起诉人不起诉。")
    elif unprosecution_type == "存疑":
        add_body_text(doc, "本院认为，本案证据不足，不符合起诉条件。")
        add_body_text(doc,
            "根据《中华人民共和国刑事诉讼法》第一百七十五条第四款的规定，"
            "决定对被不起诉人不起诉。")
    elif unprosecution_type == "附条件":
        add_body_text(doc,
            "被不起诉人（未成年人）实施了×××行为，涉嫌×××罪，"
            "可能判处一年有期徒刑以下刑罚，符合起诉条件，但有悔罪表现。")
        add_body_text(doc,
            "根据《中华人民共和国刑事诉讼法》第二百八十二条第一款的规定，"
            "决定对被不起诉人附条件不起诉。")

    add_empty_line(doc, 2)

    # 签名
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.RIGHT)
    run = p.add_run("检察员：×××")
    set_font(run)

    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.RIGHT)
    run = p.add_run(data.get("date", "20××年×月×日"))
    set_font(run)

    return doc


def create_jianchayijian(data):
    """
    生成审查逮捕意见书

    data: dict
    """
    doc = Document()
    set_page_margins(doc)

    add_title(doc, data.get("procuratorate", "×××人民检察院"))
    add_title(doc, "审 查 逮 捕 意 见 书")

    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    run = p.add_run(data.get("case_number", "×检×捕审〔20××〕×号"))
    set_font(run)

    add_empty_line(doc)

    # 案件来源
    add_heading_text(doc, "一、案件来源")
    add_body_text(doc, data.get("case_source",
        "犯罪嫌疑人×××涉嫌×××罪一案，由×××公安局提请批准逮捕。"))

    add_empty_line(doc)

    # 犯罪嫌疑人基本情况
    add_heading_text(doc, "二、犯罪嫌疑人基本情况")
    d = data.get("suspect", {})
    info = (
        f"犯罪嫌疑人{d.get('name', '×××')}，{d.get('gender', '男')}，"
        f"{d.get('birth_date', '×年×月×日')}出生，"
        f"公民身份号码{d.get('id_number', '×××')}，"
        f"{d.get('ethnicity', '×')}族，"
        f"文化程度{d.get('education', '×××')}，"
        f"户籍所在地{d.get('household', '×××')}，"
        f"现住{d.get('address', '×××')}，"
        f"职业{d.get('occupation', '×××')}。"
    )
    add_body_text(doc, info)

    add_empty_line(doc)

    # 案件事实
    add_heading_text(doc, "三、案件事实")
    add_body_text(doc, data.get("facts", "（案件事实）"))

    add_empty_line(doc)

    # 证据情况
    add_heading_text(doc, "四、证据情况")
    for i, evidence in enumerate(data.get("evidence", []), 1):
        add_body_text(doc, f"{i}. {evidence}")

    add_empty_line(doc)

    # 审查意见
    add_heading_text(doc, "五、审查意见")

    add_body_text(doc, "（一）犯罪构成分析", bold=True)
    add_body_text(doc, data.get("composition_analysis", "（犯罪构成分析）"))

    add_body_text(doc, "（二）证据审查", bold=True)
    add_body_text(doc, data.get("evidence_review", "（证据审查意见）"))

    add_body_text(doc, "（三）逮捕必要性审查", bold=True)
    add_body_text(doc, data.get("necessity_review", "（社会危险性评估）"))

    add_body_text(doc, "（四）认罪认罚情况", bold=True)
    add_body_text(doc, data.get("plea_status", "（认罪认罚情况）"))

    add_empty_line(doc)

    # 处理意见
    add_heading_text(doc, "六、处理意见")
    opinion = data.get("opinion",
        "综上所述，犯罪嫌疑人×××的行为涉嫌×××罪，事实清楚，证据确实、充分，"
        "符合逮捕条件。本院拟作出批准逮捕的决定。")
    add_body_text(doc, opinion)

    add_empty_line(doc, 2)

    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.RIGHT)
    run = p.add_run("承办人：×××")
    set_font(run)

    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.RIGHT)
    run = p.add_run(data.get("date", "20××年×月×日"))
    set_font(run)

    return doc


def main():
    parser = argparse.ArgumentParser(description="检察文书格式化工具")
    parser.add_argument("type", choices=["起诉书", "不起诉决定书", "审查逮捕意见书"],
                        help="文书类型")
    parser.add_argument("json_file", help="案件信息 JSON 文件路径")
    parser.add_argument("-o", "--output", help="输出 Word 文件路径（默认与 JSON 文件同名）")

    args = parser.parse_args()

    # 读取 JSON 数据
    with open(args.json_file, "r", encoding="utf-8") as f:
        data = json.load(f)

    # 生成文书
    if args.type == "起诉书":
        doc = create_qishu(data)
    elif args.type == "不起诉决定书":
        doc = create_buqisu(data)
    elif args.type == "审查逮捕意见书":
        doc = create_jianchayijian(data)

    # 保存文件
    if args.output:
        output_path = args.output
    else:
        json_path = Path(args.json_file)
        output_path = str(json_path.with_suffix(".docx"))

    doc.save(output_path)
    print(f"文书已生成：{output_path}")


# 供 Claude Code 直接调用的函数接口
def generate_from_dict(doc_type, data, output_path):
    """
    从字典数据生成检察文书

    Args:
        doc_type: 文书类型 ("起诉书" / "不起诉决定书" / "审查逮捕意见书")
        data: 案件信息字典
        output_path: 输出 Word 文件路径

    Returns:
        输出文件路径
    """
    if doc_type == "起诉书":
        doc = create_qishu(data)
    elif doc_type == "不起诉决定书":
        doc = create_buqisu(data)
    elif doc_type == "审查逮捕意见书":
        doc = create_jianchayijian(data)
    else:
        raise ValueError(f"不支持的文书类型：{doc_type}")

    doc.save(output_path)
    return output_path


if __name__ == "__main__":
    main()
