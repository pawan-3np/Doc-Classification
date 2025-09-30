import fitz  # PyMuPDF
import yaml
import os
from collections import defaultdict

# === Load YAML Rules ===
def load_doc_type_rules(yml_path):
    if not os.path.exists(yml_path):
        print(f"Error: The rules file '{yml_path}' does not exist.")
        return None
    with open(yml_path, "r", encoding="utf-8") as f:
        rules = yaml.safe_load(f)
    return rules.get("doc_types", None)

# === Extract Text Per Page ===
def extract_texts(pdf_path):
    pdf = fitz.open(pdf_path)
    return [pdf[i].get_text().lower() for i in range(len(pdf))]

def identify_doc_type(text, rules):
    for doc_type, props in rules.items():
        for kw in props["match_keywords"]:
            # Normalize keyword and text
            kw_clean = kw.lower().replace("'", "").replace("-", " ").replace(",", " ").replace(".", " ")
            text_clean = text.replace("'", "").replace("-", " ").replace(",", " ").replace(".", " ")
            if kw_clean in text_clean:
                return doc_type
    return None

# === Group Pages by Document Type ===
def split_pdf_by_doc_type(pdf_path, yml_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    rules = load_doc_type_rules(yml_path)
    if not rules:
        print("No rules found or failed to load rules.")
        return

    texts = extract_texts(pdf_path)
    print(f"Total pages extracted: {len(texts)}")

    grouped_pages = defaultdict(list)
    unknown_pages = []

    for idx, text in enumerate(texts):
        doc_type = identify_doc_type(text, rules)
        print(f"Page {idx}: Classified as {doc_type if doc_type else 'Unknown'}")
        if doc_type:
            grouped_pages[doc_type].append(idx)
        else:
            unknown_pages.append(idx)

    # Re-open PDF for actual page extraction
    pdf = fitz.open(pdf_path)

    # Save each group to a separate PDF
    for doc_type, page_nums in grouped_pages.items():
        new_doc = fitz.open()
        for page_num in page_nums:
            new_doc.insert_pdf(pdf, from_page=page_num, to_page=page_num)
        output_file = os.path.join(output_dir, f"{doc_type}.pdf")
        new_doc.save(output_file)
        new_doc.close()
        print(f"✅ Saved: {output_file} ({len(page_nums)} pages)")

    if unknown_pages:
        print(f"⚠️ {len(unknown_pages)} pages could not be classified.")
        unknown_doc = fitz.open()
        for page_num in unknown_pages:
            unknown_doc.insert_pdf(pdf, from_page=page_num, to_page=page_num)
        unknown_doc.save(os.path.join(output_dir, "unclassified.pdf"))
        unknown_doc.close()

    pdf.close()

# === Run It ===
if __name__ == "__main__":
    split_pdf_by_doc_type(
        pdf_path=r"C:\Users\PawanMagapalli\Downloads\document\Doc-Classification\merged doc.pdf",
        yml_path=r"rule.yml",
        output_dir="output_docs"
    )
