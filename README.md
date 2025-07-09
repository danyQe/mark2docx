# Markdown to Word Converter

[![Python Version](https://img.shields.io/badge/python-3.7%2B-blue)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![PyPI Version](https://img.shields.io/pypi/v/markdown-to-word)](https://pypi.org/project/markdown-to-word/)

A Python library and command-line tool to convert Markdown files to Microsoft Word documents (.docx) with proper formatting and styling.

## Features

- 📝 **Full Markdown Support**: Headings, bold, italic, links, code blocks, blockquotes, and more
- 📊 **Table Support**: Converts Markdown tables to properly formatted Word tables
- 🎨 **Customizable Styles**: Easily modify fonts, colors, and sizes
- 🔧 **CLI and API**: Use as a command-line tool or Python library
- 📦 **Minimal Dependencies**: Only requires python-docx, markdown2, and beautifulsoup4
- 🚀 **Fast and Efficient**: Processes documents quickly with low memory usage

## Installation

### From PyPI

```bash
pip install mark2docx
```

### From Source

```bash
git clone https://github.com/danyQe/mark2docx.git
cd mark2docx
pip install -e .
```

## Quick Start

### Command Line Usage

Convert a Markdown file to Word:

```bash
mark2docx input.md
```

Specify a custom output filename:

```bash
mark2docx input.md -o output.docx
```

### Python API Usage

```python
from mark2docx import MarkdownToWordConverter

# Create converter instance
converter = MarkdownToWordConverter()

# Convert markdown file
converter.convert_markdown_to_word("content", "output.docx")

# Or convert from string
with open("input.md", "r") as f:
    markdown_content = f.read()
    
converter.convert_markdown_to_word(markdown_content, "output.docx")
```

## Supported Markdown Elements

| Element | Markdown Syntax | Word Output |
|---------|----------------|-------------|
| Heading 1 | `# Heading` | Heading 1 style |
| Heading 2-6 | `## Heading` | Heading 2-6 styles |
| Bold | `**text**` | Bold text |
| Italic | `*text*` | Italic text |
| Code | `` `code` `` | Monospace font |
| Link | `[text](url)` | Blue underlined text with URL |
| Unordered List | `- item` | Bulleted list |
| Ordered List | `1. item` | Numbered list |
| Blockquote | `> quote` | Indented paragraph |
| Code Block | ` ```code``` ` | Monospace block |
| Table | Pipe syntax | Word table with borders |
| Horizontal Rule | `---` | Line separator |

## Examples

### Basic Document

```markdown
# Project Report

## Introduction
This is a **comprehensive** report with *various* formatting options.

### Key Features
- Feature 1: Fast conversion
- Feature 2: Preserves formatting
- Feature 3: Easy to use

### Code Example
```python
def convert_document(input_file):
    converter = MarkdownToWordConverter()
    converter.convert_markdown_to_word(input_file, "output.docx")
```

> Important: Always backup your files before conversion.
```

### Advanced Usage

```python
from markdown_to_word import MarkdownToWordConverter
from docx.shared import Pt, RGBColor

# Create converter with custom configuration
converter = MarkdownToWordConverter()

# Customize heading colors
doc = converter.doc
for i in range(1, 7):
    heading_style = doc.styles[f'Heading {i}']
    heading_style.font.color.rgb = RGBColor(0, 0, 255)  # Blue headings

# Convert with custom styles
converter.convert_markdown_to_word(markdown_content, "styled_output.docx")
```

## Configuration

You can customize the output by modifying the converter's styles:

```python
converter = MarkdownToWordConverter()

# Access the document object
doc = converter.doc

# Modify styles
code_style = doc.styles['Code']
code_style.font.name = 'Monaco'
code_style.font.size = Pt(11)
```

## Development

### Setting Up Development Environment

```bash
# Clone the repository
git clone https://github.com/danyQe/mark2docx.git
cd markdown-to-word

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install in development mode
pip install -e ".[dev]"
```

### Running Tests

```bash
pytest tests/
```

### Building Documentation

```bash
cd docs
make html
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [python-docx](https://python-docx.readthedocs.io/) for Word document manipulation
- [markdown2](https://github.com/trentm/python-markdown2) for Markdown parsing
- [Beautiful Soup](https://www.crummy.com/software/BeautifulSoup/) for HTML parsing

## Changelog

### Version 1.0.0 (2024-01-15)
- Initial release
- Full Markdown to Word conversion support
- Command-line interface
- Python API

## Support

- 📧 Email: raogoutham374@example.com
- 🐛 Issues: [GitHub Issues](https://github.com/danyQe/mark2docx/issues)
- 💬 Discussions: [GitHub Discussions](https://github.com/danyQe/mark2docx/discussions)
