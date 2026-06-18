# Letter Generation

## Template
Use `LetterTemplate` for formal correspondence with:
- Sender and recipient blocks
- Date, subject line (Re:)
- Body paragraphs, closing, signature

```python
from anecita import LetterTemplate

letter = LetterTemplate(
    sender="Dr. Emily Carter\n789 Pine Road\nLos Angeles, CA 90001",
    recipient="Prof. Michael Torres\n321 Cedar Lane\nChicago, IL 60601",
    subject="Conference Speaker Invitation",
)
letter.build(
    body_paragraphs=[
        "Dear Prof. Torres,",
        "We would be honored to invite you as a keynote speaker at the "
        "2026 International Conference on Artificial Intelligence, "
        "scheduled for September 15-17 in Los Angeles.",
        "Your expertise in machine learning ethics would provide invaluable "
        "insights to our attendees. Please let us know your availability "
        "by August 1st.",
    ],
    closing="Warm regards",
    signature_name="Dr. Emily Carter",
)
letter.builder.save("letter.docx")
```

## Style Notes
- Wide margins for formal appearance
- Bold subject line in primary color
- Standard business letter formatting
