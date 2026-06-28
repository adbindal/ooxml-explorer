
## [2026-06-28] Ingestion Defect: <w:cantSplit> (docx)
- **Judge Analysis**: The GENERATED document contains multiple inaccuracies. The 'definition' incorrectly expands the scope of 'cantSplit' to include paragraphs, whereas 'w:cantSplit' is specific to table rows. Consequently, the 'parents' array incorrectly includes 'pPr' (Paragraph Properties) in addition to 'trPr'. Furthermore, the GENERATED document incorrectly omits the required 'w:val' attribute, which is an essential part of the 'cantSplit' element's definition. The 'citation' also points to an incorrect section for this element's primary usage. The GOLDEN document is accurate regarding the element's scope, attributes, parents, and citation.
- **Golden Citation**: `ECMA-376 Part 1, Section 17.4.6`
- **Generated Citation**: `ECMA-376 Part 1, Section 17.4.6`
- **Golden Definition**: Specifies that the contents of the current table row shall not be split across a page break. If the row cannot fit entirely within the remaining space on the current page, the entire row is moved to the start of the next page.
- **Generated Definition**: Specifies that the contents of the current table row shall not be split across a page break. If the row cannot fit entirely within the remaining space on the current page, the entire row is moved to the start of the next page.
- **Golden Attributes**: `["w:val"]`
- **Generated Attributes**: `["w:val"]`
---
