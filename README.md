# SimpleDocBuilder

**SimpleDocBuilder** is a powerful yet simple Python library designed to build Microsoft Word documents (`.docx`) programmatically, block by block. It adopts a modular approach where each element (text, image, table, etc.) is handled as an independent block, ensuring robustness and flexibility.

## Features

*   **Modular Architecture**: Build documents sequentially using a fluent API.
*   **Robust Context Management**: Uses `with SimpleDocBuilder() as doc:` to safely manage temporary files and cleanup.
*   **Rich Content Support**:
    *   **Text & Titles**: Add styled paragraphs and headings easily.
    *   **Pandas DataFrames**: Convert DataFrames to formatted Word tables with automatic "Engineering Notation" support.
    *   **Images**: Insert images with optional captions and resizing.
    *   **Templates**: Inject content into existing `.docx` templates using `docxtpl`.
    *   **LaTeX & HTML**: Convert LaTeX equations and HTML content directly to Word (requires `pandoc`).
    *   **Great Tables / Complex HTML**: Render complex HTML tables as high-quality images.

## Installation

This project uses `uv` for dependency management.

```bash
# Clone the repository
git clone <repository-url>
cd simpledocbuilder

# Install dependencies
uv sync
```

### External Dependencies
For full functionality, you need:
*   **Pandoc**: For LaTeX/HTML conversion (`pypandoc`).
*   **wkhtmltopdf**: For rendering complex HTML to images (`imgkit`).

## Usage

Here is a quick example of how to use `SimpleDocBuilder`:

```python
import pandas as pd
from simpledocbuilder import SimpleDocBuilder

# Create some data
df = pd.DataFrame({
    'Parameter': ['Voltage', 'Current', 'Power'],
    'Value': [220.5, 0.05, 11.025],
    'Unit': ['V', 'A', 'W']
})

# Build the document
with SimpleDocBuilder() as doc:
    
    # Add Title and Text
    doc.add_title("Engineering Report")
    doc.add_text("This report was generated automatically.")
    
    # Add a Table (with automatic engineering formatting for numbers)
    doc.add_title("Measurements", level=2)
    doc.add_table(df, title="Experimental Data")
    
    # Add an Image
    doc.add_title("Visuals", level=2)
    # doc.add_image("path/to/chart.png", title="System Chart")
    
    # Add LaTeX Equation
    doc.add_title("Theory", level=2)
    doc.add_latex(r"P = U \times I")

    # Finalize
    doc.build("final_report.docx")
```

## Advanced Features

### Engineering Notation
The `add_table` method automatically formats numerical values in DataFrames using engineering notation (e.g., `0.0012` -> `1.20m`) thanks to the internal `utils.eng_string` helper.

### Templating
You can use `docxtpl` templates for specific blocks:

```python
doc.add_image("photo.jpg", template_path="templates/image_container.docx")
```

## Contributing

1.  Fork the repository.
2.  Create your feature branch.
3.  Commit your changes.
4.  Push to the branch.
5.  Open a Pull Request.

## License

[Add License Here]
