"""
Generate PDF Documentation from Markdown
=========================================
This script converts the UserForm documentation markdown to a PDF.
Works on any platform (Windows, Linux, Mac).

Requirements:
    pip install reportlab

Usage:
    python generate_userform_pdf.py
"""

import os
import re
from pathlib import Path

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
    from reportlab.lib import colors
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, PageBreak,
        Preformatted, Table, TableStyle, KeepTogether
    )
    from reportlab.lib.colors import HexColor
except ImportError:
    print("Error: reportlab is required")
    print("Install with: pip install reportlab")
    exit(1)


class MarkdownToPDF:
    """Convert markdown documentation to PDF."""

    def __init__(self, markdown_file, output_pdf):
        self.markdown_file = Path(markdown_file)
        self.output_pdf = Path(output_pdf)
        self.styles = getSampleStyleSheet()
        self._create_custom_styles()
        self.elements = []

    def _create_custom_styles(self):
        """Create custom paragraph styles."""
        # Title style
        self.styles.add(ParagraphStyle(
            name='CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=24,
            textColor=HexColor('#1f4788'),
            spaceAfter=30,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        ))

        # H1 style
        self.styles.add(ParagraphStyle(
            name='CustomH1',
            parent=self.styles['Heading1'],
            fontSize=18,
            textColor=HexColor('#2e5c8a'),
            spaceAfter=12,
            spaceBefore=12,
            fontName='Helvetica-Bold'
        ))

        # H2 style
        self.styles.add(ParagraphStyle(
            name='CustomH2',
            parent=self.styles['Heading2'],
            fontSize=14,
            textColor=HexColor('#4a7ba7'),
            spaceAfter=10,
            spaceBefore=10,
            fontName='Helvetica-Bold'
        ))

        # H3 style
        self.styles.add(ParagraphStyle(
            name='CustomH3',
            parent=self.styles['Heading3'],
            fontSize=12,
            textColor=HexColor('#5a8bb8'),
            spaceAfter=8,
            spaceBefore=8,
            fontName='Helvetica-Bold'
        ))

        # Code style
        self.styles.add(ParagraphStyle(
            name='CodeBlock',
            parent=self.styles['Code'],
            fontSize=9,
            fontName='Courier',
            leftIndent=20,
            rightIndent=20,
            spaceAfter=10,
            spaceBefore=10,
            backColor=HexColor('#f5f5f5'),
            borderColor=HexColor('#cccccc'),
            borderWidth=1,
            borderPadding=10
        ))

        # Body text
        self.styles.add(ParagraphStyle(
            name='CustomBody',
            parent=self.styles['BodyText'],
            fontSize=11,
            alignment=TA_JUSTIFY,
            spaceAfter=10
        ))

        # Bullet points
        self.styles.add(ParagraphStyle(
            name='CustomBullet',
            parent=self.styles['BodyText'],
            fontSize=11,
            leftIndent=30,
            spaceAfter=5
        ))

    def parse_markdown(self):
        """Parse markdown file and convert to PDF elements."""
        with open(self.markdown_file, 'r', encoding='utf-8') as f:
            content = f.read()

        lines = content.split('\n')
        i = 0
        in_code_block = False
        code_lines = []

        while i < len(lines):
            line = lines[i]

            # Code block
            if line.startswith('```'):
                if in_code_block:
                    # End of code block
                    code_text = '\n'.join(code_lines)
                    pre = Preformatted(code_text, self.styles['Code'])
                    self.elements.append(pre)
                    self.elements.append(Spacer(1, 0.2*inch))
                    code_lines = []
                    in_code_block = False
                else:
                    # Start of code block
                    in_code_block = True
                i += 1
                continue

            if in_code_block:
                code_lines.append(line)
                i += 1
                continue

            # Headings
            if line.startswith('# '):
                text = line[2:].strip()
                if i == 0:  # First heading is title
                    self.elements.append(Spacer(1, 1*inch))
                    self.elements.append(Paragraph(text, self.styles['CustomTitle']))
                else:
                    self.elements.append(Spacer(1, 0.3*inch))
                    self.elements.append(Paragraph(text, self.styles['CustomH1']))
            elif line.startswith('## '):
                text = line[3:].strip()
                self.elements.append(Paragraph(text, self.styles['CustomH2']))
            elif line.startswith('### '):
                text = line[4:].strip()
                self.elements.append(Paragraph(text, self.styles['CustomH3']))
            elif line.startswith('#### '):
                text = line[5:].strip()
                self.elements.append(Paragraph(text, self.styles['Heading4']))

            # Horizontal rule
            elif line.strip() == '---':
                self.elements.append(Spacer(1, 0.2*inch))
                self.elements.append(Table(
                    [['']],
                    colWidths=[7*inch],
                    style=TableStyle([
                        ('LINEABOVE', (0,0), (-1,0), 2, colors.grey),
                    ])
                ))
                self.elements.append(Spacer(1, 0.2*inch))

            # Bullet points
            elif line.strip().startswith('- ') or line.strip().startswith('* '):
                text = line.strip()[2:]
                # Handle checkboxes
                text = text.replace('[ ]', '☐').replace('[x]', '☑').replace('✅', '✓')
                bullet = Paragraph(f"• {text}", self.styles['CustomBullet'])
                self.elements.append(bullet)

            # Bold text in lists
            elif line.strip().startswith('**') and line.strip().endswith('**'):
                text = line.strip()[2:-2]
                self.elements.append(Paragraph(f"<b>{text}</b>", self.styles['CustomBody']))

            # Regular paragraph
            elif line.strip() and not line.startswith('┌') and not line.startswith('│') and not line.startswith('└') and not line.startswith('├') and not line.startswith('├'):
                # Handle ASCII art boxes as code
                if any(c in line for c in ['┌', '┐', '└', '┘', '│', '─', '├', '┤', '┬', '┴', '┼']):
                    # Collect box lines
                    box_lines = [line]
                    i += 1
                    while i < len(lines) and any(c in lines[i] for c in ['┌', '┐', '└', '┘', '│', '─', '├', '┤', '┬', '┴', '┼']):
                        box_lines.append(lines[i])
                        i += 1

                    box_text = '\n'.join(box_lines)
                    pre = Preformatted(box_text, self.styles['Code'])
                    self.elements.append(pre)
                    self.elements.append(Spacer(1, 0.2*inch))
                    continue
                else:
                    # Regular text - handle inline formatting
                    text = line.strip()
                    # Convert markdown bold
                    text = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', text)
                    # Convert markdown italic
                    text = re.sub(r'\*(.*?)\*', r'<i>\1</i>', text)
                    # Convert markdown code
                    text = re.sub(r'`(.*?)`', r'<font name="Courier">\1</font>', text)

                    self.elements.append(Paragraph(text, self.styles['CustomBody']))

            # Empty line
            elif not line.strip():
                self.elements.append(Spacer(1, 0.1*inch))

            i += 1

    def generate_pdf(self):
        """Generate the PDF document."""
        print(f"Generating PDF: {self.output_pdf}")

        # Create document
        doc = SimpleDocTemplate(
            str(self.output_pdf),
            pagesize=letter,
            rightMargin=0.75*inch,
            leftMargin=0.75*inch,
            topMargin=0.75*inch,
            bottomMargin=0.75*inch
        )

        # Parse markdown
        self.parse_markdown()

        # Build PDF
        doc.build(self.elements)
        print(f"PDF generated successfully: {self.output_pdf}")


def main():
    """Main execution function."""
    print("=" * 70)
    print("UserForm Documentation PDF Generator")
    print("=" * 70)
    print()

    script_dir = Path(__file__).parent
    markdown_file = script_dir / "USERFORM_DOCUMENTATION.md"
    output_pdf = script_dir / "UserForm_Documentation.pdf"

    if not markdown_file.exists():
        print(f"Error: Markdown file not found: {markdown_file}")
        return

    try:
        converter = MarkdownToPDF(markdown_file, output_pdf)
        converter.generate_pdf()

        print("\n" + "=" * 70)
        print("SUCCESS!")
        print("=" * 70)
        print(f"\nPDF documentation: {output_pdf}")
        print(f"File size: {output_pdf.stat().st_size / 1024:.1f} KB")

    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
