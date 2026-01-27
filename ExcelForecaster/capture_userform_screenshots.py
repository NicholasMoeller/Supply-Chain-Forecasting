"""
UserForm Screenshot Capture Script
====================================
This script automates Excel to display the UserForms and capture screenshots,
then generates a PDF document showing what the compiled UserForms look like.

Requirements:
    pip install pywin32 pillow reportlab

Usage:
    python capture_userform_screenshots.py

Note: This script must be run on Windows with Excel installed.
"""

import os
import sys
import time
from pathlib import Path

try:
    import win32com.client
    from PIL import ImageGrab, Image
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
    from reportlab.platypus import Table, TableStyle, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
except ImportError as e:
    print(f"Error: Missing required package: {e}")
    print("\nPlease install required packages:")
    print("  pip install pywin32 pillow reportlab")
    sys.exit(1)


class UserFormCapturer:
    """Captures screenshots of Excel UserForms and generates PDF documentation."""

    def __init__(self, excel_file_path):
        """
        Initialize the capturer.

        Args:
            excel_file_path: Path to the Excel file containing the UserForms
        """
        self.excel_file = Path(excel_file_path).resolve()
        self.output_dir = self.excel_file.parent / "screenshots"
        self.output_dir.mkdir(exist_ok=True)
        self.excel = None
        self.workbook = None
        self.screenshots = []

    def start_excel(self):
        """Start Excel application and open the workbook."""
        print("Starting Excel...")
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.Visible = True
        self.excel.DisplayAlerts = False

        print(f"Opening workbook: {self.excel_file}")
        self.workbook = self.excel.Workbooks.Open(str(self.excel_file))
        time.sleep(2)  # Wait for Excel to fully load

    def capture_userform(self, form_name, description):
        """
        Display a UserForm and capture its screenshot.

        Args:
            form_name: Name of the UserForm to display
            description: Description for the documentation
        """
        try:
            print(f"Capturing {form_name}...")

            # Show the UserForm using VBA
            vba_code = f"{form_name}.Show vbModeless"
            self.excel.Run(vba_code)

            # Wait for form to display
            time.sleep(2)

            # Capture screenshot
            screenshot = ImageGrab.grab()
            screenshot_path = self.output_dir / f"{form_name}.png"
            screenshot.save(screenshot_path)

            print(f"  Screenshot saved: {screenshot_path}")

            # Store screenshot info
            self.screenshots.append({
                'name': form_name,
                'path': screenshot_path,
                'description': description
            })

            # Close the UserForm
            try:
                self.excel.Run(f"Unload {form_name}")
            except:
                pass

            time.sleep(1)

        except Exception as e:
            print(f"  Warning: Could not capture {form_name}: {e}")

    def generate_pdf(self, output_pdf):
        """
        Generate a PDF document with all captured screenshots.

        Args:
            output_pdf: Path to the output PDF file
        """
        print(f"\nGenerating PDF: {output_pdf}")

        doc = SimpleDocTemplate(
            str(output_pdf),
            pagesize=letter,
            rightMargin=0.5*inch,
            leftMargin=0.5*inch,
            topMargin=0.5*inch,
            bottomMargin=0.5*inch
        )

        # Container for the 'Flowable' objects
        elements = []

        # Define styles
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=24,
            textColor=colors.HexColor('#1f4788'),
            spaceAfter=30,
            alignment=1  # Center
        )
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontSize=16,
            textColor=colors.HexColor('#2e5c8a'),
            spaceAfter=12
        )

        # Title page
        elements.append(Spacer(1, 2*inch))
        elements.append(Paragraph("Supply Chain Forecasting Tool", title_style))
        elements.append(Paragraph("UserForm Interface Documentation", heading_style))
        elements.append(Spacer(1, 0.5*inch))
        elements.append(Paragraph(
            f"Generated: {time.strftime('%Y-%m-%d %H:%M:%S')}",
            styles['Normal']
        ))
        elements.append(PageBreak())

        # Add each screenshot
        for idx, screenshot_info in enumerate(self.screenshots, 1):
            # Section heading
            elements.append(Paragraph(
                f"{idx}. {screenshot_info['name']}",
                heading_style
            ))
            elements.append(Spacer(1, 0.2*inch))

            # Description
            elements.append(Paragraph(
                screenshot_info['description'],
                styles['Normal']
            ))
            elements.append(Spacer(1, 0.3*inch))

            # Screenshot image
            img = Image.open(screenshot_info['path'])

            # Resize image to fit page width
            max_width = 7*inch
            max_height = 8*inch
            img_width, img_height = img.size

            # Calculate scaling
            scale = min(max_width / img_width, max_height / img_height, 1.0)
            new_width = img_width * scale
            new_height = img_height * scale

            # Add image to PDF
            img_flowable = RLImage(
                str(screenshot_info['path']),
                width=new_width,
                height=new_height
            )
            elements.append(img_flowable)

            # Page break between screenshots
            if idx < len(self.screenshots):
                elements.append(PageBreak())

        # Build PDF
        doc.build(elements)
        print(f"PDF generated successfully: {output_pdf}")

    def close_excel(self):
        """Close Excel application."""
        if self.workbook:
            self.workbook.Close(SaveChanges=False)
        if self.excel:
            self.excel.Quit()
        print("Excel closed.")


def main():
    """Main execution function."""
    print("=" * 70)
    print("UserForm Screenshot Capture and PDF Generation")
    print("=" * 70)
    print()

    # Get the Excel file path
    script_dir = Path(__file__).parent
    excel_file = script_dir / "TimeSeriesForecaster.xlsm"

    # Check if file exists
    if not excel_file.exists():
        print(f"Error: Excel file not found: {excel_file}")
        print("\nPlease ensure the Excel workbook is in the same directory as this script.")
        print("Expected file: TimeSeriesForecaster.xlsm")
        sys.exit(1)

    # Output PDF path
    output_pdf = script_dir / "UserForm_Documentation.pdf"

    # Create capturer
    capturer = UserFormCapturer(excel_file)

    try:
        # Start Excel
        capturer.start_excel()

        # Capture ForecastGUI
        capturer.capture_userform(
            "ForecastGUI",
            "The main Time Series Forecasting Tool interface. This form allows users to "
            "load CSV data, configure forecasting parameters (frequency, horizon, seasonal type), "
            "and run analysis using Simple Exponential Smoothing and Holt-Winters methods. "
            "Results can be viewed as charts and exported to CSV."
        )

        # Capture BatchForecastGUI
        capturer.capture_userform(
            "BatchForecastGUI",
            "The Batch Forecasting Tool for multi-component analysis. This advanced interface "
            "supports processing 50-60+ components simultaneously. Users can choose between "
            "single component or batch processing modes, configure data formats (wide/long), "
            "and enable full diagnostics or quick mode for efficient processing."
        )

        # Generate PDF
        capturer.generate_pdf(output_pdf)

        print("\n" + "=" * 70)
        print("SUCCESS!")
        print("=" * 70)
        print(f"\nScreenshots saved to: {capturer.output_dir}")
        print(f"PDF documentation: {output_pdf}")

    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()

    finally:
        # Always close Excel
        capturer.close_excel()


if __name__ == "__main__":
    main()
