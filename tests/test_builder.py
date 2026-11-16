import pandas as pd
from docx import Document
from PIL import Image

from simpledocbuilder import DocumentConfig, SimpleDocBuilder, __version__
from simpledocbuilder.utils import eng_string


def test_version_exposed():
    assert isinstance(__version__, str)
    assert len(__version__) > 0


def test_eng_string():
    assert eng_string(1230.0) == "1.23e3"
    assert eng_string(1230.0, si=True) == "1.23k"
    assert eng_string(-1230000.0, si=True) == "-1.23M"
    assert eng_string(0.0) == "0"
    assert eng_string(1e-4) == "0"


def test_build_empty_document(tmp_path):
    output = tmp_path / "empty.docx"
    with SimpleDocBuilder() as doc:
        doc.build(output)
    assert not output.exists()


def test_build_text_table_and_image(tmp_path):
    img_path = tmp_path / "sample.png"
    Image.new("RGB", (40, 20), color="blue").save(img_path)

    df = pd.DataFrame(
        {
            "Parameter": ["Voltage", "Current", "Missing", "Active"],
            "Value": [220.5, 1500.0, float("nan"), 0.0],
            "Flag": [True, False, True, False],
        }
    )

    output = tmp_path / "out" / "report.docx"
    config = DocumentConfig(default_image_width_mm=100)

    with SimpleDocBuilder(config=config) as doc:
        doc.add_title("Engineering Report", level=1)
        doc.add_text("Introductory paragraph.")
        doc.add_table(df, title="Measurements")
        doc.add_image(img_path, title="Sample Image")
        doc.add_image(tmp_path / "does_not_exist.png", title="Missing Image")
        doc.build(output)

    assert output.exists()
    generated = Document(str(output))
    full_text = "\n".join(p.text for p in generated.paragraphs)
    assert "Engineering Report" in full_text
    assert "Introductory paragraph." in full_text
    assert "Measurements" in full_text
    assert "[ERREUR: Image introuvable - does_not_exist.png]" in full_text
    assert len(generated.tables) == 1
