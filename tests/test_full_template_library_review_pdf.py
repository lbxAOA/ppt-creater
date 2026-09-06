from pathlib import Path
from pypdf import PdfReader


def test_full_template_library_review_pdf_covers_local_and_network_families():
    pdf = Path(r"C:/ppt-creater/output/full-production-template-library-review.pdf")
    assert pdf.exists()
    reader = PdfReader(str(pdf))
    assert len(reader.pages) >= 35
    titles = [item.title for item in reader.outline if hasattr(item, "title")]
    expected = {
        "L01 Modern Report | AI / Investor / Academic",
        "L02 Data & Boardroom | Corporate",
        "L03 Company Profile | Brand",
        "L04 Event & Culture | Local",
        "L05 Wedding | Local",
        "L06 Career & Portfolio | Local",
        "N01 Investor Modern | NET-01",
        "N02 Company Professional | NET-02",
        "N03 Career Portfolio | NET-03",
        "N04 Wedding Amelia | NET-04",
    }
    assert expected <= set(titles)
