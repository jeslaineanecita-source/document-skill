# Report Generation

## Template
Use `ReportTemplate` for structured reports with:
- Title page (title, author, date)
- Multiple sections with headings and body content
- Tables, images, paragraphs per section

```python
from anecita import ReportTemplate

report = ReportTemplate(
    title="Annual Report 2026",
    author="Jane Doe",
)
report.build([
    {
        "heading": "Executive Summary",
        "body": [
            {"type": "paragraph", "text": "This quarter exceeded targets..."},
        ],
    },
    {
        "heading": "Revenue Breakdown",
        "body": [
            {"type": "table", "headers": ["Quarter", "Revenue"], "rows": [["Q1", "$50K"]]},
        ],
    },
])
report.builder.save("report.docx")
```

## Style Notes
- Use `Colors.PRIMARY` for headings, `Colors.DARK` for body
- Tables use `Table Grid` style with colored headers
