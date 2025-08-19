import pdfplumber

def classify_page(text):
    """
    Simple classification logic based on keywords.
    You can expand this with your own business rules.
    """
    text_lower = text.lower()

    if "invoice" in text_lower:
        return "Invoice Page"
    elif "purchase order" in text_lower:
        return "Purchase Order Page"
    elif "packing list" in text_lower:
        return "Packing List"
    else:
        return "Other / Unclassified"

def classify_pdf_pages(pdf_path):
    results = []

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            classification = classify_page(text)
            results.append({
                "page": i,
                "classification": classification
            })

    return results

# -------- MAIN PROGRAM --------
if __name__ == "__main__":
    pdf_path = input("Enter full PDF file path: ").strip()
    results = classify_pdf_pages(pdf_path)

    print("\nClassification Results:")
    for r in results:
        print(f"Page {r['page']}: {r['classification']}")
