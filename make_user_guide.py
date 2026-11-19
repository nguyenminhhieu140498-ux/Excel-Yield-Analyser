from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt
from pathlib import Path

OUTPUT_FILE = "Excel_Yield_Analyzer_Beginner_Guide.pptx"

# Colors
NAVY = RGBColor(25, 45, 80)
BLUE = RGBColor(54, 105, 170)
LIGHT_BLUE = RGBColor(226, 238, 250)
PALE = RGBColor(246, 248, 251)
DARK = RGBColor(42, 48, 58)
GRAY = RGBColor(102, 112, 128)
WHITE = RGBColor(255, 255, 255)
GREEN = RGBColor(39, 130, 87)
ORANGE = RGBColor(220, 137, 48)
RED = RGBColor(185, 66, 66)


def set_background(slide, color=PALE):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_text(
    slide, text, x, y, w, h, size=20, color=DARK,
    bold=False, align=PP_ALIGN.LEFT,
):
    box = slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = Inches(0.06)
    frame.margin_right = Inches(0.06)
    frame.margin_top = Inches(0.04)
    frame.margin_bottom = Inches(0.03)
    frame.vertical_anchor = MSO_ANCHOR.TOP

    lines = str(text).split("\n")
    for index, line in enumerate(lines):
        paragraph = (
            frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        )
        paragraph.text = line
        paragraph.font.name = "Aptos"
        paragraph.font.size = Pt(size)
        paragraph.font.bold = bold
        paragraph.font.color.rgb = color
        paragraph.alignment = align
        paragraph.space_after = Pt(7)

    return box


def add_title(slide, title, subtitle=None, number=None):
    add_text(slide, title, 0.6, 0.35, 12.1, 0.55, 28, NAVY, True)
    if subtitle:
        add_text(slide, subtitle, 0.63, 0.95, 12.0, 0.45, 14, GRAY)
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.62), Inches(1.43), Inches(12.05), Inches(0.035)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = BLUE
    line.line.fill.background()

    if number is not None:
        add_text(
            slide,
            f"Excel Yield Analyzer  •  Beginner Guide  •  {number}",
            0.65, 7.12, 12.0, 0.22, 9, GRAY,
        )


def add_card(slide, title, body, x, y, w, h, accent=BLUE):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = WHITE
    shape.line.color.rgb = RGBColor(215, 222, 232)

    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(0.09), Inches(h)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = accent
    bar.line.fill.background()

    add_text(slide, title, x + 0.22, y + 0.16, w - 0.42, 0.42, 17, NAVY, True)
    add_text(slide, body, x + 0.22, y + 0.67, w - 0.42, h - 0.8, 14, DARK)
    return shape


def add_step(slide, number, title, body, y, accent=BLUE):
    circle = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(0.8), Inches(y), Inches(0.52), Inches(0.52)
    )
    circle.fill.solid()
    circle.fill.fore_color.rgb = accent
    circle.line.fill.background()
    tf = circle.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = str(number)
    p.alignment = PP_ALIGN.CENTER
    p.font.name = "Aptos"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = WHITE

    add_text(slide, title, 1.55, y - 0.01, 10.8, 0.35, 18, NAVY, True)
    add_text(slide, body, 1.55, y + 0.38, 10.8, 0.58, 14, DARK)


def new_slide(prs, title, subtitle, number):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide)
    add_title(slide, title, subtitle, number)
    return slide


def build_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Slide 1: Cover
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, NAVY)

    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0), Inches(0.2), Inches(7.5)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = RGBColor(83, 158, 220)
    accent.line.fill.background()

    add_text(
        slide, "EXCEL YIELD ANALYZER",
        0.85, 1.45, 11.6, 0.75, 34, WHITE, True,
    )
    add_text(
        slide,
        "Beginner’s step-by-step guide",
        0.9, 2.35, 11.2, 0.55, 25, RGBColor(188, 215, 245),
    )
    add_text(
        slide,
        "Import yield reports • Enrich them with mapping data • "
        "Build Pass/Fail commonality tables • Export results",
        0.92, 3.25, 10.8, 1.0, 19, WHITE,
    )
    add_text(
        slide,
        "A practical guide for manufacturing, yield, quality, and process teams",
        0.92, 6.55, 11.0, 0.35, 13, RGBColor(208, 218, 232),
    )

    # Slide 2: Start the app
    slide = new_slide(
        prs,
        "1. Open the application",
        "Start the Python source version from the project folder.",
        2,
    )
    add_step(
        slide, 1, "Open the project in VS Code",
        "Open the folder containing app.py and engine.py. In VS Code, choose "
        "Terminal → New Terminal.",
        1.8,
    )
    add_step(
        slide, 2, "Activate the project environment",
        "Run:  source .venv/bin/activate\n"
        "The terminal prompt should begin with (.venv).",
        2.85,
    )
    add_step(
        slide, 3, "Start the app",
        "Run:  python app.py\n"
        "If your terminal says “python: command not found,” activate .venv "
        "or run the Python executable available on your Mac.",
        3.9,
    )
    add_card(
        slide, "To stop the app",
        "Click the Terminal and press Control + C. If you are using the "
        "packaged .app, close its window normally.",
        0.85, 5.25, 11.6, 1.1, ORANGE,
    )

    # Slide 3: Import yield workbook
    slide = new_slide(
        prs,
        "2. Import your yield report",
        "Choose the workbook and the worksheet or worksheets to analyze.",
        3,
    )
    add_step(
        slide, 1, "Open Import Files",
        "Click Choose Yield File… and select your .xlsx or .xlsm report.",
        1.8,
    )
    add_step(
        slide, 2, "Choose worksheets",
        "Select the relevant sheets, such as FATP, CR, CQA, or ORT. "
        "Use multi-select if you want to combine several sheets.",
        2.85,
    )
    add_step(
        slide, 3, "Check the preview",
        "Review the column names and the first rows. Confirm that headers "
        "are in the first row and that the expected data was loaded.",
        3.9,
    )
    add_card(
        slide, "Tip",
        "Use clear, consistent headers. Check that serial numbers and "
        "Pass/Fail results are present before continuing.",
        0.85, 5.35, 11.6, 1.0, GREEN,
    )

    # Slide 4: Mapping and merge
    slide = new_slide(
        prs,
        "3. Add mapping data and merge",
        "Mapping lists can add vendor, product, line, station, or issue details.",
        4,
    )
    add_step(
        slide, 1, "Choose the mapping workbook",
        "In Import Files, click Choose Mapping File… and select its worksheet.",
        1.8,
    )
    add_step(
        slide, 2, "Set the merge keys",
        "Open Mapping. Select the matching yield and mapping columns, "
        "for example Serial_Number and SN.",
        2.85,
    )
    add_step(
        slide, 3, "Choose mapping fields",
        "Select the mapping columns to append, such as Vendor, Product, "
        "or Station. Start with a Left join unless you need another join type.",
        3.9,
    )
    add_step(
        slide, 4, "Merge and review",
        "Click Merge data. Check matched rows, unmatched rows, match rate, "
        "and duplicate mapping keys.",
        4.95,
    )
    add_card(
        slide, "Fuzzy matching",
        "Enable similar-key matching only when appropriate. Review the "
        "results carefully: similar serial numbers can be matched incorrectly.",
        0.85, 6.0, 11.6, 0.72, ORANGE,
    )

    # Slide 5: Analysis layout
    slide = new_slide(
        prs,
        "4. Build a Tabulate-style analysis",
        "Drag fields from the left into the drop zones on the right.",
        5,
    )
    add_card(
        slide, "Rows",
        "Put grouping fields here, such as phase, TC vendor, or TP vendor. "
        "These identify each result row.",
        0.75, 1.8, 5.7, 1.15, BLUE,
    )
    add_card(
        slide, "Columns",
        "Optional additional grouping fields. Use this zone when you want "
        "another category to organize the results.",
        6.75, 1.8, 5.7, 1.15, BLUE,
    )
    add_card(
        slide, "Values",
        "Put the Pass/Fail test-result columns here, for example "
        "31(TP2TC offset) and 30(TP2TC GAP).",
        0.75, 3.25, 5.7, 1.15, GREEN,
    )
    add_card(
        slide, "Statistics",
        "Drag N for an integer count and Row % for the percentage of Pass "
        "or Fail within each group.",
        6.75, 3.25, 5.7, 1.15, GREEN,
    )
    add_text(
        slide,
        "To remove a field: select it in a drop zone and click Remove selected, "
        "or drag it back to the available-column list. Click Start over to "
        "clear the layout.",
        0.85, 5.15, 11.7, 0.95, 16, DARK,
    )

    # Slide 6: Example configuration
    slide = new_slide(
        prs,
        "5. Example: Pass/Fail table by vendor",
        "Use this setup to get separate Fail and Pass columns for each test.",
        6,
    )
    add_card(
        slide, "Drag into Rows",
        "phase\nTC vendor\nTP vendor",
        0.8, 1.9, 3.7, 2.0, BLUE,
    )
    add_card(
        slide, "Drag into Values",
        "31(TP2TC offset)\n30(TP2TC GAP)",
        4.8, 1.9, 3.7, 2.0, GREEN,
    )
    add_card(
        slide, "Drag into Statistics",
        "N\nRow %",
        8.8, 1.9, 3.7, 2.0, ORANGE,
    )
    add_text(
        slide,
        "Click Update table. For each value field, the results are organized "
        "under Fail and Pass, with N and Row % beneath each status. For example, "
        "“31(TP2TC offset) – Fail – N” and “31(TP2TC offset) – Pass – Row %.”",
        0.9, 4.55, 11.6, 1.15, 17, DARK,
    )
    add_card(
        slide, "Reading Row %",
        "Row % is calculated within each displayed group. Fail % plus Pass % "
        "should be approximately 100% when all results are recognized.",
        0.85, 6.0, 11.6, 0.72, BLUE,
    )

    # Slide 7: Export and box plot
    slide = new_slide(
        prs,
        "6. Export the report",
        "Export the analysis and optionally add a vendor box plot.",
        7,
    )
    add_step(
        slide, 1, "Run the analysis first",
        "Click Update table so the latest Tabulate result is ready for export.",
        1.8,
    )
    add_step(
        slide, 2, "Configure the optional box plot",
        "On Export, choose Vendor, a numeric measurement, and the Pass/Fail "
        "status column. Check the Pass and Fail labels used in your data.",
        2.85,
    )
    add_step(
        slide, 3, "Save the Excel report",
        "Click Export Excel Report… and choose a destination. The workbook "
        "contains the merged data, analysis summaries, and optional chart.",
        3.9,
    )
    add_card(
        slide, "Box plot note",
        "A box plot needs numeric measurements. A categorical Pass/Fail "
        "column alone cannot show a measurement distribution; the chart also "
        "overlays recognized Pass/Fail points.",
        0.85, 5.4, 11.6, 1.0, ORANGE,
    )

    # Slide 8: Troubleshooting
    slide = new_slide(
        prs,
        "7. Quick troubleshooting",
        "Common beginner issues and what to check first.",
        8,
    )
    add_card(
        slide, "“python: command not found”",
        "Activate the virtual environment: source .venv/bin/activate. "
        "If needed, use the Python command installed on your Mac.",
        0.8, 1.8, 5.75, 1.35, RED,
    )
    add_card(
        slide, "App opens, then exits with an error",
        "Read the Terminal traceback. Check that dependencies are installed "
        "in the active environment and that engine.py is beside app.py.",
        6.75, 1.8, 5.75, 1.35, ORANGE,
    )
    add_card(
        slide, "No matching rows",
        "Confirm the selected key columns and inspect spelling, spaces, "
        "blank keys, and merge statistics. Use fuzzy matching cautiously.",
        0.8, 3.55, 5.75, 1.35, BLUE,
    )
    add_card(
        slide, "Pass/Fail columns look empty",
        "Check the test fields contain recognized labels such as Pass/Fail "
        "or OK/NG, and check that the correct fields are in Values.",
        6.75, 3.55, 5.75, 1.35, GREEN,
    )
    add_text(
        slide,
        "If you need help, copy the complete Terminal error message and note "
        "which button you clicked immediately before it appeared.",
        0.9, 5.55, 11.5, 0.7, 16, NAVY, True,
    )

    prs.save(OUTPUT_FILE)
    print(f"Created: {Path(OUTPUT_FILE).resolve()}")


if __name__ == "__main__":
    build_deck()