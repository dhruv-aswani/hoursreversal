# Reversal Mapping Session Status & Summary

This document summarizes the current state, logic, and output files of the **CyberRani Reversal Mapping** task. It is stored directly in the workspace directory so that we can easily pick up work exactly where we left off.

---

## 1. File & Batch Status
All target allocations have been successfully executed, formatted, and validated.

* **Source WIP**: `CyberRani_WIP_May2026.xlsx`
* **Batch 1 Output**: `reversal_mapping_batch_1.xlsx` (123 allocations validated)
* **Batch 2 Output**: `reversal_mapping_batch_2.xlsx` (126 allocations validated)
* **Reversal Script**: `reversal_mapping.py`

---

## 2. Technical Features implemented

### A. Automatic Multi-Batch Exclusions
* The script automatically scans all previously generated batch files matching the pattern `reversal_mapping_batch_*.xlsx` in sequential order.
* It parses the allocated hours from these batch files and subtracts them from the source rows before making new allocations, ensuring **zero double allocations** or over-allocations across runs.
* It completely ignores the deleted `clean_final_reversal_mapping.xlsx` file and dynamically scales based on whatever batch files exist in the folder.

### B. Fallback Hierarchy
We implemented smart, hierarchical fallback rules to resolve shortages in specific designations:
* `Senior Manager` ➔ falls back to `Manager`
* `Senior Analyst` ➔ falls back to `Associate Consultant` ➔ falls back to `Consultant`
* `Associate Consultant` ➔ falls back to `Consultant` (when Associate Consultant hours are exhausted)

### C. Exact Visual Formatting (Matching Your Reference Sheet)
The generated batch files perfectly match the formatting of the reference worksheets:
* **Headers**: Solid yellow `#FFFF00` fill, bold Calibri font, center-aligned, with thin borders on all sides.
* **Vertically Merged Cells**: Vertically merges Columns A (`Target Employee`), B (`Designation`), and C (`Target ANSR`) for each employee block.
* **Left-Top Alignment**: Vertically merged cells are aligned **Left** and **Top** for clean readability.
* **Standard Data Cells**: Standard Calibri font, center-aligned, with thin borders.
* **Accounting Number Format**: Custom number formatting `_ * #,##0_ ;_ * \-#,##0_ ;_ * "-"??_ ;_ @_ ` is set for all numeric columns (ANSR/Target ANSR).
* **Merged Total Rows**: Columns A to G are merged for the Total row, displaying `'Total'` center-aligned. The ANSR sum is located in Column H.
* **Gray Fills on Total rows**: Column A and Column H on Total rows contain solid light gray `#D9D9D9` fills; others are transparent with top/bottom thin borders to form a clean outer block.
* **Column Widths**: Exact widths set (Col A=20, Col B=20, Col C=16, Col D=22, Col E=20, Col F=17, Col G=10, Col H=15).

---

## 3. Data Integrity & Validation (PASS ✅)
We ran a combined validation script `validate_both_batches.py` which cross-references both files against the source.

* **Total Combined allocations validated**: 249 rows
* **Combined Errors**: 0
* **Combined Warnings**: 0

Everything is 100% accurate, preventing any double-allocation and ensuring exact rate proportions.

---

## 4. Instructions for Next Run
When you are ready to run the next batch of targets:
1. Open `reversal_mapping.py`.
2. Edit the `targets` list in the script with the new target employee names, their designations, and target values.
3. Run the script:
   ```bash
   python3 reversal_mapping.py
   ```
4. The script will automatically load allocations from `reversal_mapping_batch_1.xlsx` and `reversal_mapping_batch_2.xlsx`, subtract those hours, run the new targets on the remaining pool, and generate `reversal_mapping_batch_3.xlsx` with matching formatting and proper fallbacks!
