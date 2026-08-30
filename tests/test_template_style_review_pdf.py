from pathlib import Path
from pypdf import PdfReader


def test_template_style_review_pdf_exists_and_has_all_sources():
    pdf = Path(r"C:/ppt-creater/output/production-template-style-review.pdf")
    assert pdf.exists()
    reader = PdfReader(str(pdf))
    # Six source templates are shown as six slides per landscape review page.
    assert len(reader.pages) >= 20
    titles = [item.title for item in reader.outline if hasattr(item, "title")]
    for title in [
        "01 现代汇报 / AI / 投资人 / 科研",
        "02 数据经营 / 管理层 / 企业通用",
        "03 企业介绍 / 品牌提案",
        "04 专业活动 / 文化文旅",
        "05 婚礼 / 婚庆",
        "06 简历 / 个人作品集",
    ]:
        assert title in titles
