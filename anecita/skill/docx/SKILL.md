# Docx Builder Skill

Use python-docx (https://python-docx.readthedocs.io/) to generate DOCX files programmatically.

## Capabilities
- Create documents with headings, paragraphs, lists, tables, images
- Apply styles, fonts, colors, alignment
- Add headers, footers, page numbers
- Insert tables with merged cells and formatting
- Add images with sizing and positioning
- Handle document sections and page breaks

## Conventions
- Use `Document()` from `docx` module
- Use `python-docx` library (install via `pip install python-docx`)
- Save output files with `.docx` extension
- Use `Inches` or `Cm` from `docx.shared` for dimensions
- Use `Pt` from `docx.shared` for font sizes
- Use `RGBColor` from `docx.shared` for colors
- Use `WD_ALIGN_PARAGRAPH` from `docx.enum.text` for alignment

## Common Workflow
1. Create a `Document()` instance
2. Add content (paragraphs, tables, images)
3. Style appropriately
4. Save with `document.save(path)`
