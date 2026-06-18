# Invoice Generation

## Template
Use `InvoiceTemplate` for professional invoices with:
- Invoice number, company/client info
- Line items table (qty, unit price, total)
- Subtotal, tax, grand total calculation

```python
from anecita import InvoiceTemplate

invoice = InvoiceTemplate(
    invoice_number="INV-2026-001",
    company_name="Acme Corp\n123 Main St\nNew York, NY 10001",
    client_name="Client Name\n456 Oak Ave",
)
invoice.build(
    line_items=[
        {"description": "Web Development", "qty": 40, "unit_price": 150},
        {"description": "Hosting (Monthly)", "qty": 1, "unit_price": 29.99},
    ],
    tax_rate=0.08,
    notes="Payment due within 30 days.",
)
invoice.builder.save("invoice.docx")
```

## Style Notes
- Right-aligned title with invoice number
- Auto-calculated totals section
