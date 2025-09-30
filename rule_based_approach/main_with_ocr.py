import fitz  # PyMuPDF
import yaml
import os
from collections import defaultdict
from PIL import Image
import pytesseract

def load_doc_type_rules(yml_path):
    if not os.path.exists(yml_path):
        print(f"Error: The rules file '{yml_path}' does not exist.")
        return None
    with open(yml_path, "r", encoding="utf-8") as f:
        rules = yaml.safe_load(f)
    return rules.get("doc_types", None)

def extract_text_with_ocr(page):
    text = page.get_text("text").strip()
    if text:
        return text
    # Fallback to OCR with advanced preprocessing if no text extracted
    pix = page.get_pixmap(dpi=400)  # Further increase DPI for better OCR
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    # Preprocess: convert to grayscale
    img = img.convert("L")
    # Preprocess: increase contrast
    from PIL import ImageEnhance
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(2.5)
    # Preprocess: binarize with adjustable threshold
    threshold = 110  # Lower threshold for faint text
    img = img.point(lambda x: 0 if x < threshold else 255, '1')
    # Use pytesseract config for better accuracy
    text = pytesseract.image_to_string(img, config='--psm 6')
    return text.strip() if text.strip() else "[EMPTY]"

def identify_doc_type(text, rules):
    text_clean = text.lower().replace("'", "").replace("-", " ").replace(",", " ").replace(".", " ")
    for doc_type, props in rules.items():
        for kw in props["match_keywords"]:
            kw_clean = kw.lower().replace("'", "").replace("-", " ").replace(",", " ").replace(".", " ")
            if kw_clean in text_clean:
                return doc_type
    return None

def split_pdf_by_doc_type(pdf_path, yml_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    rules = load_doc_type_rules(yml_path)
    if not rules:
        print("No rules found or failed to load rules.")
        return
    pdf = fitz.open(pdf_path)
    print(f"Total pages in PDF: {len(pdf)}")
    grouped_pages = defaultdict(list)
    unknown_pages = []
    page_summaries = []
    for i, page in enumerate(pdf):
        text = extract_text_with_ocr(page)
        # Save extracted text to file for inspection
        text_file_path = os.path.join(output_dir, f"page_{i+1}_extracted.txt")
        with open(text_file_path, "w", encoding="utf-8") as txt_file:
            txt_file.write(text)
        preview = text[:200].replace('\n', ' ')
        print(f"--- Page {i} | Text Length: {len(text)} ---")
        doc_type = identify_doc_type(text, rules)
        print(f"Page {i}: Classified as {doc_type if doc_type else 'Unknown'} | Preview: {preview}...")
        page_summaries.append({"page": i, "type": doc_type if doc_type else 'Unknown', "preview": preview})
        if doc_type:
            grouped_pages[doc_type].append(i)
        else:
            unknown_pages.append(i)
    print("\n=== Document Summary ===")
    for summary in page_summaries:
        print(f"Page {summary['page']}: {summary['type']} | {summary['preview']}...")
    # Save each group to a separate PDF
    for doc_type, page_indices in grouped_pages.items():
        new_doc = fitz.open()
        for idx in page_indices:
            new_doc.insert_pdf(pdf, from_page=idx, to_page=idx)
        output_file = os.path.join(output_dir, f"{doc_type}.pdf")
        new_doc.save(output_file)
        new_doc.close()
        print(f"✅ Saved: {output_file} ({len(page_indices)} pages)")
    print(f"\nTotal original pages: {len(pdf)}")
    total_extracted = sum(len(pages) for pages in grouped_pages.values())
    print(f"Total classified pages: {total_extracted}")
    print(f"Total unclassified pages: {len(unknown_pages)}")
    if unknown_pages:
        print(f"⚠️ {len(unknown_pages)} pages could not be classified.")
        unknown_doc = fitz.open()
        for idx in unknown_pages:
            unknown_doc.insert_pdf(pdf, from_page=idx, to_page=idx)
        unknown_doc.save(os.path.join(output_dir, "unclassified.pdf"))
        unknown_doc.close()

if __name__ == "__main__":
    split_pdf_by_doc_type(
        pdf_path=r"C:\Users\PawanMagapalli\Downloads\ilovepdf_merged.pdf",
        yml_path=r"rule.yml",
        output_dir="output_docs"
    )
