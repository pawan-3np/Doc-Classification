# import fitz  # PyMuPDF
# import yaml
# import os
# from collections import defaultdict
# from PIL import Image
# import pytesseract

# def load_doc_type_rules(yml_path):
#     if not os.path.exists(yml_path):
#         print(f"Error: The rules file '{yml_path}' does not exist.")
#         return None
#     with open(yml_path, "r", encoding="utf-8") as f:
#         rules = yaml.safe_load(f)
#     return rules.get("doc_types", None)

# def extract_text_with_ocr(page):
#     text = page.get_text("text").strip()
#     if text:
#         return text
#     # Fallback to OCR with advanced preprocessing if no text extracted
#     pix = page.get_pixmap(dpi=400)  # Further increase DPI for better OCR
#     img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
#     # Preprocess: convert to grayscale
#     img = img.convert("L")
#     # Preprocess: increase contrast
#     from PIL import ImageEnhance
#     enhancer = ImageEnhance.Contrast(img)
#     img = enhancer.enhance(2.5)
#     # Preprocess: binarize with adjustable threshold
#     threshold = 110  # Lower threshold for faint text
#     img = img.point(lambda x: 0 if x < threshold else 255, '1')
#     # Use pytesseract config for better accuracy
#     text = pytesseract.image_to_string(img, config='--psm 6')
#     return text.strip() if text.strip() else "[EMPTY]"

# def identify_doc_type(text, rules):
#     text_clean = text.lower().replace("'", "").replace("-", " ").replace(",", " ").replace(".", " ")
#     for doc_type, props in rules.items():
#         for kw in props["match_keywords"]:
#             kw_clean = kw.lower().replace("'", "").replace("-", " ").replace(",", " ").replace(".", " ")
#             if kw_clean in text_clean:
#                 return doc_type
#     return None

# def split_pdf_by_doc_type(pdf_path, yml_path, output_dir):
#     os.makedirs(output_dir, exist_ok=True)
#     rules = load_doc_type_rules(yml_path)
#     if not rules:
#         print("No rules found or failed to load rules.")
#         return
#     pdf = fitz.open(pdf_path)
#     print(f"Total pages in PDF: {len(pdf)}")
#     grouped_pages = defaultdict(list)
#     unknown_pages = []
#     page_summaries = []
#     for i, page in enumerate(pdf):
#         text = extract_text_with_ocr(page)
#         # Save extracted text to file for inspection
#         text_file_path = os.path.join(output_dir, f"page_{i+1}_extracted.txt")
#         with open(text_file_path, "w", encoding="utf-8") as txt_file:
#             txt_file.write(text)
#         preview = text[:200].replace('\n', ' ')
#         print(f"--- Page {i} | Text Length: {len(text)} ---")
#         doc_type = identify_doc_type(text, rules)
#         print(f"Page {i}: Classified as {doc_type if doc_type else 'Unknown'} | Preview: {preview}...")
#         page_summaries.append({"page": i, "type": doc_type if doc_type else 'Unknown', "preview": preview})
#         if doc_type:
#             grouped_pages[doc_type].append(i)
#         else:
#             unknown_pages.append(i)
#     print("\n=== Document Summary ===")
#     for summary in page_summaries:
#         print(f"Page {summary['page']}: {summary['type']} | {summary['preview']}...")
#     # Save each group to a separate PDF
#     for doc_type, page_indices in grouped_pages.items():
#         new_doc = fitz.open()
#         for idx in page_indices:
#             new_doc.insert_pdf(pdf, from_page=idx, to_page=idx)
#         output_file = os.path.join(output_dir, f"{doc_type}.pdf")
#         new_doc.save(output_file)
#         new_doc.close()
#         print(f"✅ Saved: {output_file} ({len(page_indices)} pages)")
#     print(f"\nTotal original pages: {len(pdf)}")
#     total_extracted = sum(len(pages) for pages in grouped_pages.values())
#     print(f"Total classified pages: {total_extracted}")
#     print(f"Total unclassified pages: {len(unknown_pages)}")
#     if unknown_pages:
#         print(f"⚠️ {len(unknown_pages)} pages could not be classified.")
#         unknown_doc = fitz.open()
#         for idx in unknown_pages:
#             unknown_doc.insert_pdf(pdf, from_page=idx, to_page=idx)
#         unknown_doc.save(os.path.join(output_dir, "unclassified.pdf"))
#         unknown_doc.close()

# if __name__ == "__main__":
#     split_pdf_by_doc_type(
#         pdf_path=r"C:\Users\PawanMagapalli\Downloads\ilovepdf_merged (1).pdf",
#         yml_path=r"rule.yml",
#         output_dir="output_docs"
#     )

# import fitz  # PyMuPDF
# import yaml
# import os
# from collections import defaultdict
# from PIL import Image
# import pytesseract

# # ---------- Load YAML Rules ----------
# def load_doc_type_rules(yml_path):
#     if not os.path.exists(yml_path):
#         print(f"❌ Error: The rules file '{yml_path}' does not exist.")
#         return None
#     with open(yml_path, "r", encoding="utf-8") as f:
#         rules = yaml.safe_load(f)
#     return rules.get("doc_types", None)

# # ---------- Normalize Text for Matching ----------
# def normalize_text(text):
#     return (
#         text.lower()
#         .replace("'", "")
#         .replace("-", " ")
#         .replace("_", " ")
#         .replace(",", " ")
#         .replace(".", " ")
#         .strip()
#     )

# # ---------- Extract Text with OCR (Fallback) ----------
# def extract_text_with_ocr(page):
#     text = page.get_text("text").strip()
#     if text:
#         return text
#     # Use OCR if no text found
#     pix = page.get_pixmap(dpi=400)
#     img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
#     img = img.convert("L")
#     from PIL import ImageEnhance
#     enhancer = ImageEnhance.Contrast(img)
#     img = enhancer.enhance(2.5)
#     threshold = 110
#     img = img.point(lambda x: 0 if x < threshold else 255, '1')
#     text = pytesseract.image_to_string(img, config='--psm 6')
#     return text.strip() if text.strip() else "[EMPTY]"

# # ---------- Identify Document Type ----------
# def identify_doc_type(text, rules):
#     text_clean = normalize_text(text)
#     for doc_type, props in rules.items():
#         if not props or "match_keywords" not in props:
#             print(f"⚠️ Skipping invalid rule: {doc_type}")
#             continue
#         for kw in props["match_keywords"]:
#             kw_clean = normalize_text(kw)
#             if kw_clean in text_clean:
#                 return doc_type
#     return None

# # ---------- Split PDF by Detected Document Type ----------
# def split_pdf_by_doc_type(pdf_path, yml_path, output_dir):
#     os.makedirs(output_dir, exist_ok=True)
#     rules = load_doc_type_rules(yml_path)
#     if not rules:
#         print("❌ No rules found or failed to load rules.")
#         return

#     pdf = fitz.open(pdf_path)
#     print(f"📄 Total pages in PDF: {len(pdf)}")
#     grouped_pages = defaultdict(list)
#     unknown_pages = []
#     page_summaries = []

#     for i, page in enumerate(pdf):
#         text = extract_text_with_ocr(page)
#         text_file_path = os.path.join(output_dir, f"page_{i+1}_extracted.txt")
#         with open(text_file_path, "w", encoding="utf-8") as txt_file:
#             txt_file.write(text)

#         preview = text[:200].replace('\n', ' ')
#         doc_type = identify_doc_type(text, rules)
#         doc_type_display = doc_type if doc_type else "unknown"
#         print(f"Page {i}: 📂 {doc_type_display} | Preview: {preview}...")
#         page_summaries.append({
#             "page": i,
#             "type": doc_type_display,
#             "preview": preview
#         })

#         if doc_type:
#             grouped_pages[doc_type].append(i)
#         else:
#             unknown_pages.append(i)

#     print("\n=== 📊 Document Summary ===")
#     for summary in page_summaries:
#         print(f"Page {summary['page']}: {summary['type']} | {summary['preview']}...")

#     # Save grouped PDFs
#     for doc_type, page_indices in grouped_pages.items():
#         new_doc = fitz.open()
#         for idx in page_indices:
#             new_doc.insert_pdf(pdf, from_page=idx, to_page=idx)
#         output_file = os.path.join(output_dir, f"{doc_type}.pdf")
#         new_doc.save(output_file)
#         new_doc.close()
#         print(f"✅ Saved: {output_file} ({len(page_indices)} pages)")

#     if unknown_pages:
#         print(f"\n⚠️ {len(unknown_pages)} pages could not be classified.")
#         unknown_doc = fitz.open()
#         for idx in unknown_pages:
#             unknown_doc.insert_pdf(pdf, from_page=idx, to_page=idx)
#         unknown_doc.save(os.path.join(output_dir, "unclassified.pdf"))
#         unknown_doc.close()

#     print(f"\n📊 Total original pages: {len(pdf)}")
#     print(f"📌 Total classified pages: {sum(len(p) for p in grouped_pages.values())}")
#     print(f"❓ Total unclassified pages: {len(unknown_pages)}")

# # ---------- Main Execution ----------
# if __name__ == "__main__":
#     split_pdf_by_doc_type(
#         pdf_path=r"C:\Users\PawanMagapalli\Downloads\ilovepdf_merged (1).pdf",
#         yml_path=r"rule.yml",  # Snake_case YAML
#         output_dir="output_docs"
#     )


# import fitz  # PyMuPDF
# import yaml
# import os
# from collections import defaultdict
# from PIL import Image, ImageEnhance
# import pytesseract
# import re

# def load_doc_type_rules(yml_path):
#     if not os.path.exists(yml_path):
#         print(f"❌ Error: Rules file '{yml_path}' not found.")
#         return None
#     with open(yml_path, "r", encoding="utf-8") as f:
#         rules = yaml.safe_load(f)
#     return rules.get("doc_types", None)

# def normalize_text(text):
#     # Lowercase, remove special chars, replace spaces/dashes with underscores
#     text = text.lower()
#     # Replace non-alphanumeric with underscore
#     text = re.sub(r"[^a-z0-9]+", "_", text)
#     # Trim leading/trailing underscores
#     text = text.strip("_")
#     return text

# def extract_text_with_ocr(page, save_debug_image_path=None):
#     text = page.get_text("text").strip()
#     if text:
#         return text

#     # No text? Use OCR
#     pix = page.get_pixmap(dpi=400)
#     img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

#     # Preprocessing for OCR
#     img = img.convert("L")  # Grayscale
#     enhancer = ImageEnhance.Contrast(img)
#     img = enhancer.enhance(2.5)

#     threshold = 110
#     img = img.point(lambda x: 0 if x < threshold else 255, '1')

#     if save_debug_image_path:
#         img.save(save_debug_image_path)

#     ocr_text = pytesseract.image_to_string(img, config='--psm 6')
#     return ocr_text.strip() if ocr_text.strip() else "[EMPTY]"

# def identify_doc_type(text, rules):
#     text_norm = normalize_text(text)
#     for doc_type, props in rules.items():
#         if not props or "match_keywords" not in props:
#             continue
#         for kw in props["match_keywords"]:
#             kw_norm = normalize_text(kw)
#             if kw_norm in text_norm:
#                 return doc_type
#     return None

# def split_pdf_by_doc_type(pdf_path, yml_path, output_dir, save_ocr_debug_images=False):
#     os.makedirs(output_dir, exist_ok=True)
#     rules = load_doc_type_rules(yml_path)
#     if not rules:
#         print("❌ No valid classification rules loaded. Exiting.")
#         return

#     pdf = fitz.open(pdf_path)
#     print(f"📄 Total pages in PDF: {len(pdf)}")

#     grouped_pages = defaultdict(list)
#     unknown_pages = []
#     page_summaries = []

#     for i, page in enumerate(pdf):
#         debug_img_path = None
#         if save_ocr_debug_images:
#             debug_img_path = os.path.join(output_dir, f"page_{i+1}_ocr_debug.png")
#         text = extract_text_with_ocr(page, save_debug_image_path=debug_img_path)
        
#         # Save raw extracted text for review
#         with open(os.path.join(output_dir, f"page_{i+1}_extracted.txt"), "w", encoding="utf-8") as f:
#             f.write(text)

#         preview = text[:200].replace('\n', ' ').strip()
#         doc_type = identify_doc_type(text, rules)
#         doc_type_display = doc_type if doc_type else "unknown"
#         print(f"Page {i+1}: Classified as '{doc_type_display}' | Preview: {preview}...")

#         page_summaries.append({"page": i+1, "type": doc_type_display, "preview": preview})

#         if doc_type:
#             grouped_pages[doc_type].append(i)
#         else:
#             unknown_pages.append(i)

#     print("\n=== 📊 Document Classification Summary ===")
#     for summary in page_summaries:
#         print(f"Page {summary['page']}: {summary['type']} | {summary['preview']}")

#     # Save grouped PDFs
#     for doc_type, page_indices in grouped_pages.items():
#         new_doc = fitz.open()
#         for idx in page_indices:
#             new_doc.insert_pdf(pdf, from_page=idx, to_page=idx)
#         output_file = os.path.join(output_dir, f"{doc_type}.pdf")
#         new_doc.save(output_file)
#         new_doc.close()
#         print(f"✅ Saved: {output_file} ({len(page_indices)} pages)")

#     # Save unclassified pages
#     if unknown_pages:
#         print(f"\n⚠️ {len(unknown_pages)} pages could not be classified. Saved as 'unclassified.pdf'.")
#         unknown_doc = fitz.open()
#         for idx in unknown_pages:
#             unknown_doc.insert_pdf(pdf, from_page=idx, to_page=idx)
#         unknown_doc.save(os.path.join(output_dir, "unclassified.pdf"))
#         unknown_doc.close()

#     print(f"\n📄 Total pages: {len(pdf)}")
#     print(f"📂 Classified pages: {sum(len(pages) for pages in grouped_pages.values())}")
#     print(f"❓ Unclassified pages: {len(unknown_pages)}")

# if __name__ == "__main__":
#     split_pdf_by_doc_type(
#         pdf_path=r"C:\Users\PawanMagapalli\Downloads\ilovepdf_merged (1).pdf",
#         yml_path=r"rule.yml",
#         output_dir="output_docs",
#         save_ocr_debug_images=True  # Set to True to save OCR preprocessed images for troubleshooting
#     )

# import fitz  # PyMuPDF
# import yaml
# import os
# from collections import defaultdict
# from PIL import Image
# import pytesseract
# from rapidfuzz import fuzz

# def load_doc_type_rules(yml_path):
#     if not os.path.exists(yml_path):
#         print(f"Error: The rules file '{yml_path}' does not exist.")
#         return None
#     with open(yml_path, "r", encoding="utf-8") as f:
#         rules = yaml.safe_load(f)
#     return rules.get("doc_types", None)

# def extract_text_with_ocr(page):
#     text = page.get_text("text").strip()
#     if text:
#         return text
#     # Fallback to OCR with preprocessing if no text extracted
#     pix = page.get_pixmap(dpi=400)
#     img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
#     img = img.convert("L")
#     from PIL import ImageEnhance
#     enhancer = ImageEnhance.Contrast(img)
#     img = enhancer.enhance(2.5)
#     threshold = 110
#     img = img.point(lambda x: 0 if x < threshold else 255, '1')
#     text = pytesseract.image_to_string(img, config='--psm 6')
#     return text.strip() if text.strip() else "[EMPTY]"

# def identify_doc_type(text, rules, min_matches=5, similarity_threshold=70):
#     """
#     Fuzzy match keywords against the page text.
#     - min_matches: minimum number of keyword matches required to classify.
#     - similarity_threshold: minimum similarity (0-100) to count a match.
#     Returns:
#         doc_type (str) or None,
#         matched_keywords (list),
#         match_count (int)
#     """
#     text_clean = text.lower()
#     best_doc_type = None
#     best_matched_keywords = []
#     best_count = 0

#     for doc_type, props in rules.items():
#         if not props or "match_keywords" not in props or not props["match_keywords"]:
#             continue

#         matched_keywords = []
#         for kw in props["match_keywords"]:
#             kw_lower = kw.lower()
#             similarity = fuzz.partial_ratio(kw_lower, text_clean)
#             if similarity >= similarity_threshold:
#                 matched_keywords.append(kw)

#         count = len(matched_keywords)
#         if count > best_count and count >= min_matches:
#             best_count = count
#             best_doc_type = doc_type
#             best_matched_keywords = matched_keywords

#     if best_doc_type:
#         return best_doc_type, best_matched_keywords, best_count
#     else:
#         return None, [], 0

# def split_pdf_by_doc_type(pdf_path, yml_path, output_dir):
#     os.makedirs(output_dir, exist_ok=True)
#     rules = load_doc_type_rules(yml_path)
#     if not rules:
#         print("No rules found or failed to load rules.")
#         return
#     pdf = fitz.open(pdf_path)
#     print(f"Total pages in PDF: {len(pdf)}")
#     grouped_pages = defaultdict(list)
#     unknown_pages = []
#     page_summaries = []

#     for i, page in enumerate(pdf):
#         print(f"\n--- Processing Page {i} ---")
#         text = extract_text_with_ocr(page)
#         # Save extracted text for review
#         text_file_path = os.path.join(output_dir, f"page_{i+1}_extracted.txt")
#         with open(text_file_path, "w", encoding="utf-8") as txt_file:
#             txt_file.write(text)

#         doc_type, matched_keywords, match_count = identify_doc_type(text, rules)
#         preview = text[:200].replace('\n', ' ')
#         print(f"Classified as: {doc_type if doc_type else 'Unknown'}")
#         print(f"Matched keywords ({match_count}): {matched_keywords}")
#         print(f"Preview: {preview}...")

#         page_summaries.append({
#             "page": i,
#             "type": doc_type if doc_type else 'Unknown',
#             "matched_keywords": matched_keywords,
#             "match_count": match_count,
#             "preview": preview
#         })

#         if doc_type:
#             grouped_pages[doc_type].append(i)
#         else:
#             unknown_pages.append(i)

#     print("\n=== Summary ===")
#     for summary in page_summaries:
#         print(f"Page {summary['page']}: {summary['type']} | Matched keywords ({summary['match_count']}): {summary['matched_keywords']} | {summary['preview']}...")

#     # Save grouped pages to separate PDFs
#     for doc_type, page_indices in grouped_pages.items():
#         new_doc = fitz.open()
#         for idx in page_indices:
#             new_doc.insert_pdf(pdf, from_page=idx, to_page=idx)
#         output_file = os.path.join(output_dir, f"{doc_type}.pdf")
#         new_doc.save(output_file)
#         new_doc.close()
#         print(f"✅ Saved: {output_file} ({len(page_indices)} pages)")

#     # Save unknown pages to a separate PDF if any
#     if unknown_pages:
#         print(f"\n⚠️ {len(unknown_pages)} pages could not be classified.")
#         unknown_doc = fitz.open()
#         for idx in unknown_pages:
#             unknown_doc.insert_pdf(pdf, from_page=idx, to_page=idx)
#         unknown_path = os.path.join(output_dir, "unclassified.pdf")
#         unknown_doc.save(unknown_path)
#         unknown_doc.close()
#         print(f"⚠️ Saved unclassified pages: {unknown_path}")

#     print(f"\nTotal pages: {len(pdf)}")
#     total_classified = sum(len(pages) for pages in grouped_pages.values())
#     print(f"Classified pages: {total_classified}")
#     print(f"Unclassified pages: {len(unknown_pages)}")

# if __name__ == "__main__":
#     split_pdf_by_doc_type(
#         pdf_path=r"C:\Users\PawanMagapalli\Downloads\ilovepdf_merged (1).pdf",
#         yml_path=r"rule.yml",
#         output_dir="output_docs"
#     )


# import fitz  # PyMuPDF
# import yaml
# import os
# from collections import defaultdict
# from PIL import Image
# import pytesseract

# def load_doc_type_rules(yml_path):
#     if not os.path.exists(yml_path):
#         print(f"Error: The rules file '{yml_path}' does not exist.")
#         return None
#     with open(yml_path, "r", encoding="utf-8") as f:
#         rules = yaml.safe_load(f)
#     return rules.get("doc_types", None)

# def extract_text_with_ocr(page):
#     text = page.get_text("text").strip()
#     if text:
#         return text.lower()
#     # Fallback to OCR
#     pix = page.get_pixmap(dpi=400)
#     img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
#     img = img.convert("L")
#     from PIL import ImageEnhance
#     enhancer = ImageEnhance.Contrast(img)
#     img = enhancer.enhance(2.5)
#     threshold = 110
#     img = img.point(lambda x: 0 if x < threshold else 255, '1')
#     text = pytesseract.image_to_string(img, config='--psm 6')
#     return text.lower().strip() if text.strip() else "[empty]"

# def lcs_length(X, Y):
#     """Compute length of Longest Common Subsequence between strings X and Y."""
#     m = len(X)
#     n = len(Y)
#     # Create a DP table to store lengths of LCS
#     L = [[0]*(n+1) for _ in range(m+1)]
#     for i in range(m):
#         for j in range(n):
#             if X[i] == Y[j]:
#                 L[i+1][j+1] = L[i][j] + 1
#             else:
#                 L[i+1][j+1] = max(L[i][j+1], L[i+1][j])
#     return L[m][n]

# def identify_doc_type_lcs(text, rules, min_matches=5, lcs_ratio_threshold=0.9):
#     """
#     Use LCS between keyword and text.
#     - min_matches: minimum number of keywords matched
#     - lcs_ratio_threshold: minimum ratio (LCS_length / keyword_length) to count as match
#     """
#     matched_keywords = defaultdict(list)  # doc_type -> matched keywords

#     for doc_type, props in rules.items():
#         if not props or "match_keywords" not in props or not props["match_keywords"]:
#             continue

#         count = 0
#         for kw in props["match_keywords"]:
#             kw_lower = kw.lower()
#             lcs_len = lcs_length(kw_lower, text)
#             ratio = lcs_len / len(kw_lower) if len(kw_lower) > 0 else 0
#             if ratio >= lcs_ratio_threshold:
#                 matched_keywords[doc_type].append(kw)
#                 count += 1
#         if count >= min_matches:
#             # Return first doc_type that meets criteria
#             return doc_type, matched_keywords[doc_type], count

#     return None, [], 0

# def split_pdf_by_doc_type(pdf_path, yml_path, output_dir):
#     os.makedirs(output_dir, exist_ok=True)
#     rules = load_doc_type_rules(yml_path)
#     if not rules:
#         print("No rules found or failed to load rules.")
#         return
#     pdf = fitz.open(pdf_path)
#     print(f"Total pages in PDF: {len(pdf)}")
#     grouped_pages = defaultdict(list)
#     unknown_pages = []
#     page_summaries = []

#     for i, page in enumerate(pdf):
#         print(f"\n--- Processing Page {i} ---")
#         text = extract_text_with_ocr(page)
#         # Save extracted text for review
#         text_file_path = os.path.join(output_dir, f"page_{i+1}_extracted.txt")
#         with open(text_file_path, "w", encoding="utf-8") as txt_file:
#             txt_file.write(text)

#         doc_type, matched_keywords, match_count = identify_doc_type_lcs(text, rules)
#         preview = text[:200].replace('\n', ' ')
#         print(f"Classified as: {doc_type if doc_type else 'Unknown'}")
#         print(f"Matched keywords ({match_count}): {matched_keywords}")
#         print(f"Preview: {preview}...")

#         page_summaries.append({
#             "page": i,
#             "type": doc_type if doc_type else 'Unknown',
#             "matched_keywords": matched_keywords,
#             "match_count": match_count,
#             "preview": preview
#         })

#         if doc_type:
#             grouped_pages[doc_type].append(i)
#         else:
#             unknown_pages.append(i)

#     print("\n=== Summary ===")
#     for summary in page_summaries:
#         print(f"Page {summary['page']}: {summary['type']} | Matched keywords ({summary['match_count']}): {summary['matched_keywords']} | {summary['preview']}...")

#     # Save grouped pages to separate PDFs
#     for doc_type, page_indices in grouped_pages.items():
#         new_doc = fitz.open()
#         for idx in page_indices:
#             new_doc.insert_pdf(pdf, from_page=idx, to_page=idx)
#         output_file = os.path.join(output_dir, f"{doc_type}.pdf")
#         new_doc.save(output_file)
#         new_doc.close()
#         print(f"✅ Saved: {output_file} ({len(page_indices)} pages)")

#     # Save unknown pages to a separate PDF if any
#     if unknown_pages:
#         print(f"\n⚠️ {len(unknown_pages)} pages could not be classified.")
#         unknown_doc = fitz.open()
#         for idx in unknown_pages:
#             unknown_doc.insert_pdf(pdf, from_page=idx, to_page=idx)
#         unknown_path = os.path.join(output_dir, "unclassified.pdf")
#         unknown_doc.save(unknown_path)
#         unknown_doc.close()
#         print(f"⚠️ Saved unclassified pages: {unknown_path}")

#     print(f"\nTotal pages: {len(pdf)}")
#     total_classified = sum(len(pages) for pages in grouped_pages.values())
#     print(f"Classified pages: {total_classified}")
#     print(f"Unclassified pages: {len(unknown_pages)}")

# if __name__ == "__main__":
#     split_pdf_by_doc_type(
#         pdf_path=r"C:\Users\PawanMagapalli\Downloads\ilovepdf_merged (1).pdf",
#         yml_path=r"rule.yml",
#         output_dir="output_docs"
#     )

 ### lcs approach

# import fitz  # PyMuPDF
# import yaml
# import os
# from collections import defaultdict
# from PIL import Image, ImageEnhance
# import pytesseract

# def load_doc_type_rules(yml_path):
#     if not os.path.exists(yml_path):
#         print(f"Error: The rules file '{yml_path}' does not exist.")
#         return None
#     with open(yml_path, "r", encoding="utf-8") as f:
#         rules = yaml.safe_load(f)
#     return rules.get("doc_types", None)

# def extract_text_with_ocr(page):
#     text = page.get_text("text").strip()
#     if text:
#         return text.lower()
#     # Fallback to OCR
#     pix = page.get_pixmap(dpi=400)
#     img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
#     img = img.convert("L")
#     enhancer = ImageEnhance.Contrast(img)
#     img = enhancer.enhance(2.5)
#     threshold = 110
#     img = img.point(lambda x: 0 if x < threshold else 255, '1')
#     text = pytesseract.image_to_string(img, config='--psm 6')
#     return text.lower().strip() if text.strip() else "[empty]"

# def lcs_length(X, Y):
#     """Compute length of Longest Common Subsequence between strings X and Y."""
#     m = len(X)
#     n = len(Y)
#     L = [[0]*(n+1) for _ in range(m+1)]
#     for i in range(m):
#         for j in range(n):
#             if X[i] == Y[j]:
#                 L[i+1][j+1] = L[i][j] + 1
#             else:
#                 L[i+1][j+1] = max(L[i][j+1], L[i+1][j])
#     return L[m][n]

# def identify_doc_type_lcs(text, rules, min_matches=5, lcs_ratio_threshold=0.9):
#     matched_keywords = defaultdict(list)
#     for doc_type, props in rules.items():
#         if not props or "match_keywords" not in props or not props["match_keywords"]:
#             continue
#         count = 0
#         for kw in props["match_keywords"]:
#             kw_lower = kw.lower()
#             lcs_len = lcs_length(kw_lower, text)
#             ratio = lcs_len / len(kw_lower) if len(kw_lower) > 0 else 0
#             if ratio >= lcs_ratio_threshold:
#                 matched_keywords[doc_type].append(kw)
#                 count += 1
#         if count >= min_matches:
#             return doc_type, matched_keywords[doc_type], count
#     return None, [], 0

# def split_pdf_by_doc_type(pdf_path, yml_path, output_dir):
#     os.makedirs(output_dir, exist_ok=True)
#     rules = load_doc_type_rules(yml_path)
#     if not rules:
#         print("No rules found or failed to load rules.")
#         return
#     pdf = fitz.open(pdf_path)
#     print(f"Total pages in PDF: {len(pdf)}")
#     grouped_pages = defaultdict(list)
#     unknown_pages = []
#     page_summaries = []

#     for i, page in enumerate(pdf):
#         print(f"\n--- Processing Page {i} ---")
#         text = extract_text_with_ocr(page)
#         text_file_path = os.path.join(output_dir, f"page_{i+1}_extracted.txt")
#         with open(text_file_path, "w", encoding="utf-8") as txt_file:
#             txt_file.write(text)
#         doc_type, matched_keywords, match_count = identify_doc_type_lcs(text, rules)
#         preview = text[:200].replace('\n', ' ')
#         print(f"Classified as: {doc_type if doc_type else 'Unknown'}")
#         print(f"Matched keywords ({match_count}): {matched_keywords}")
#         print(f"Preview: {preview}...")

#         page_summaries.append({
#             "page": i,
#             "type": doc_type if doc_type else 'Unknown',
#             "matched_keywords": matched_keywords,
#             "match_count": match_count,
#             "preview": preview
#         })
#         if doc_type:
#             grouped_pages[doc_type].append(i)
#         else:
#             unknown_pages.append(i)

#     print("\n=== Summary ===")
#     for summary in page_summaries:
#         print(f"Page {summary['page']}: {summary['type']} | Matched keywords ({summary['match_count']}): {summary['matched_keywords']} | {summary['preview']}...")

#     for doc_type, page_indices in grouped_pages.items():
#         new_doc = fitz.open()
#         for idx in page_indices:
#             new_doc.insert_pdf(pdf, from_page=idx, to_page=idx)
#         output_file = os.path.join(output_dir, f"{doc_type}.pdf")
#         new_doc.save(output_file)
#         new_doc.close()
#         print(f"✅ Saved: {output_file} ({len(page_indices)} pages)")

#     if unknown_pages:
#         print(f"\n⚠️ {len(unknown_pages)} pages could not be classified.")
#         unknown_doc = fitz.open()
#         for idx in unknown_pages:
#             unknown_doc.insert_pdf(pdf, from_page=idx, to_page=idx)
#         unknown_path = os.path.join(output_dir, "unclassified.pdf")
#         unknown_doc.save(unknown_path)
#         unknown_doc.close()
#         print(f"⚠️ Saved unclassified pages: {unknown_path}")

#     print(f"\nTotal pages: {len(pdf)}")
#     total_classified = sum(len(pages) for pages in grouped_pages.values())
#     print(f"Classified pages: {total_classified}")
#     print(f"Unclassified pages: {len(unknown_pages)}")

# if __name__ == "__main__":
#     split_pdf_by_doc_type(
#         pdf_path=r"C:\Users\PawanMagapalli\Downloads\ilovepdf_merged (1).pdf",
#         yml_path=r"rule.yml",
#         output_dir="output_docs"
#     )

import fitz  # PyMuPDF
import yaml
import os
from collections import defaultdict
from PIL import Image, ImageEnhance
import pytesseract
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np

# 1. Load classification rules from YAML
def load_doc_type_rules(yml_path):
    if not os.path.exists(yml_path):
        print(f"Error: The rules file '{yml_path}' does not exist.")
        return None
    with open(yml_path, "r", encoding="utf-8") as f:
        rules = yaml.safe_load(f)
    return rules.get("doc_types", None)

# 2. Extract text from a PDF page, fallback to OCR if needed
def extract_text_with_ocr(page):
    text = page.get_text("text").strip()
    if text:
        return text.lower()
    # OCR fallback
    pix = page.get_pixmap(dpi=400)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    img = img.convert("L")
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(2.5)
    threshold = 110
    img = img.point(lambda x: 0 if x < threshold else 255, '1')
    text = pytesseract.image_to_string(img, config='--psm 6')
    return text.lower().strip() if text.strip() else "[empty]"

# 3. Prepare TF-IDF vectors for keywords in rules
def prepare_keyword_vectors(rules):
    doc_type_keywords = {}
    for doc_type, props in rules.items():
        kws = props.get("match_keywords", [])
        # Join all keywords for a doc type into a single string (or keep list)
        combined_text = " ".join(kws).lower()
        doc_type_keywords[doc_type] = combined_text
    return doc_type_keywords

# 4. Classify text using cosine similarity between page text and keywords
def classify_page_text(text, doc_type_keywords, threshold=0.2):
    """
    Returns:
      - doc_type with highest similarity above threshold,
      - similarity score,
      - None if no similarity above threshold
    """
    if text.strip() == "[empty]" or not text.strip():
        return None, 0.0

    corpus = [text] + list(doc_type_keywords.values())
    vectorizer = TfidfVectorizer(stop_words='english').fit(corpus)
    tfidf_matrix = vectorizer.transform(corpus)
    page_vec = tfidf_matrix[0]
    keywords_vecs = tfidf_matrix[1:]

    similarities = cosine_similarity(page_vec, keywords_vecs).flatten()
    best_idx = np.argmax(similarities)
    best_score = similarities[best_idx]
    if best_score >= threshold:
        doc_type = list(doc_type_keywords.keys())[best_idx]
        return doc_type, best_score
    return None, 0.0

# 5. Main function: classify pages, split PDF and save results
def split_pdf_by_doc_type_nlp(pdf_path, yml_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    rules = load_doc_type_rules(yml_path)
    if not rules:
        print("No rules found or failed to load rules.")
        return

    doc_type_keywords = prepare_keyword_vectors(rules)
    pdf = fitz.open(pdf_path)
    print(f"Total pages in PDF: {len(pdf)}")

    grouped_pages = defaultdict(list)
    unknown_pages = []
    page_summaries = []

    for i, page in enumerate(pdf):
        print(f"\n--- Processing Page {i+1} ---")
        text = extract_text_with_ocr(page)
        text_file_path = os.path.join(output_dir, f"page_{i+1}_extracted.txt")
        with open(text_file_path, "w", encoding="utf-8") as f:
            f.write(text)

        doc_type, score = classify_page_text(text, doc_type_keywords)
        preview = text[:200].replace('\n', ' ')

        print(f"Classified as: {doc_type if doc_type else 'Unknown'} (Score: {score:.3f})")
        page_summaries.append({
            "page": i+1,
            "type": doc_type if doc_type else 'Unknown',
            "score": score,
            "preview": preview
        })

        if doc_type:
            grouped_pages[doc_type].append(i)
        else:
            unknown_pages.append(i)

    # Summary output
    print("\n=== Classification Summary ===")
    for summary in page_summaries:
        print(f"Page {summary['page']}: {summary['type']} (Score: {summary['score']:.3f}) | Preview: {summary['preview']}...")

    # Save split PDFs
    for doc_type, pages in grouped_pages.items():
        new_pdf = fitz.open()
        for p in pages:
            new_pdf.insert_pdf(pdf, from_page=p, to_page=p)
        output_file = os.path.join(output_dir, f"{doc_type}.pdf")
        new_pdf.save(output_file)
        new_pdf.close()
        print(f"✅ Saved classified PDF: {output_file} ({len(pages)} pages)")

    # Save unclassified pages
    if unknown_pages:
        unknown_pdf = fitz.open()
        for p in unknown_pages:
            unknown_pdf.insert_pdf(pdf, from_page=p, to_page=p)
        unknown_path = os.path.join(output_dir, "unclassified.pdf")
        unknown_pdf.save(unknown_path)
        unknown_pdf.close()
        print(f"⚠️ Saved unclassified PDF: {unknown_path} ({len(unknown_pages)} pages)")

    print(f"\nTotal pages: {len(pdf)}")
    print(f"Classified pages: {sum(len(p) for p in grouped_pages.values())}")
    print(f"Unclassified pages: {len(unknown_pages)}")

if __name__ == "__main__":
    split_pdf_by_doc_type_nlp(
        pdf_path=r"C:\Users\PawanMagapalli\Downloads\test\test_1_jessey.pdf"
        yml_path=r"rule.yml",
        output_dir="output_docs"
    )
