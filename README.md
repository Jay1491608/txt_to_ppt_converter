# Theme-Adaptive TXT-to-PPTX Converter

A Python command-line utility to convert structured text course materials (like markdown and formatted text modules) into PowerPoint presentations (`.pptx`) that dynamically inherit layouts, fonts, colors, and margins from an existing PowerPoint template.

---

## Key Features

1. **Theme-Adaptive (Style Inheritance)**:
   * No hardcoded fonts, sizes, colors, or absolute shapes inside the Python script.
   * Everything inherits styling from the template slide master. Swapping the PowerPoint template or the MS Office theme automatically repaints the entire slide deck correctly.
2. **Flexible Parsing Formats**:
   * Auto-detects four different text schemas (`m1`, `m89`, `plaintext`, and `simple`).
   * Supports slides with titles, bulleted lists (with nested indent levels), and detailed speaker notes.
3. **PowerPoint Sections (`p:extLst` Integration)**:
   * Auto-generates native collapsible PowerPoint sections (e.g., `M1T1`, `M1T2`, `M2T1`, etc.) grouping slides inside the presentation view via low-level XML trees.
4. **Batch & Merging Capabilities**:
   * Convert files individually, or batch compile multiple module text files into separate slide decks.
   * Combine multiple text files into a single, cohesive master presentation using the `--combine` flag.

---

## Installation & Setup

1. Clone or download this directory.
2. Make sure Python 3.8+ is installed.
3. Install the required dependencies from [requirements.txt]:
   ```bash
   pip install -r requirements.txt
   ```

---

## Quick Start & Usage Examples

### 1. Generating a Starter Template
If you do not have a slide template, you can generate a default blank PowerPoint format using [create_template.py](file:///c:/Users/jay/Documents/Study/ppt_converter/create_template.py):
```bash
python create_template.py
```
This creates a template at `templates/default.pptx`. You can open this file in PowerPoint and edit its Slide Master to set up customized layout placeholders.

### 2. Single-File Conversion
Convert a single structured text file to PowerPoint:
```bash
python converter.py examples/sample_input.txt templates/default.pptx output/intro_deck.pptx
```

### 3. Batch Convert Multiple Files
Convert multiple module text files into separate slide decks in the `output/` directory:
```bash
python converter.py --batch examples/M1.txt examples/M2.txt -t templates/default.pptx
```

### 4. Combine Modules into a Single Master Presentation
Parse multiple module files, merge them sequentially, group them into PowerPoint sections, and compile them into a single presentation deck:
```bash
python converter.py --batch examples/M1.txt examples/M2.txt examples/M3.txt -t templates/default.pptx --combine
```
*Outputs: `output/full_course.pptx`*

---

## Command Line Arguments

Run `python converter.py --help` to inspect all options:
*   `input_txt`: Path to a single text/markdown input file (single-file mode).
*   `template_pptx` / `-t`: Path to reference presentation template (theme and layouts).
*   `output_pptx`: Output presentation file path (single-file mode).
*   `--batch`: List of files to process in batch mode.
*   `--combine`: Merges all batch files into one presentation.
*   `-o` / `--combined-output`: Custom path for the combined output presentation (defaults to `output/full_course.pptx`).
*   `--output-dir`: Output directory for batch runs (defaults to `output`).
*   `-v` / `--verbose`: Enable debug logging to troubleshoot slide layout/placeholder indices matching.
*   `--keep-template-slides`: Retain existing slides present in the template file (otherwise cleared).
*   `--no-sections`: Disable grouping slides into PowerPoint collapsible sections.

---

## Input Formats

### Simple Markdown Format (e.g. [sample_input.txt](file:///c:/Users/jay/Documents/Study/ppt_converter/examples/sample_input.txt))
*   `# Section Title`: Generates a new Section Divider slide.
*   `## Slide Title`: Generates a standard content slide.
*   `- Bullet`: Top-level bullet points.
*   `  - Sub Bullet`: Nested bullet points (indented with 2 spaces).
*   `Notes: ...`: Speaker notes.

### Module Format (e.g. [M1.txt](file:///c:/Users/jay/Documents/Study/ppt_converter/examples/M1.txt))
*   `## TOPIC X.X — Name`: Topic section boundary.
*   `### Slide N — Slide Title`: Initiates a new slide.
*   `**Title:** ...`: Explicitly overrides title (especially on section slides).
*   `**Bullets:**`: List items starting with `-`.
*   `**Speaker Notes:**`: Appends text to slide notes.
*   `**Q&A:**`: Appends structured question-and-answer texts to speaker notes.

---

## Project Structure & Files

*   **[converter.py](file:///c:/Users/jay/Documents/Study/ppt_converter/converter.py)**: Contains the main CLI interface, parsers for different text formats, placeholder indexing, and sections injection.
*   **[inspect_template.py](file:///c:/Users/jay/Documents/Study/ppt_converter/inspect_template.py)**: Prints layout and placeholder index details from a chosen template for debugging.
*   **[inspect_deck.py](file:///c:/Users/jay/Documents/Study/ppt_converter/inspect_deck.py)**: Displays generated slides and their layout names.
*   **`examples/`**: Preconfigured structured course files (e.g., [M1.txt](file:///c:/Users/jay/Documents/Study/ppt_converter/examples/M1.txt) to [M9.txt](file:///c:/Users/jay/Documents/Study/ppt_converter/examples/M9.txt)).
