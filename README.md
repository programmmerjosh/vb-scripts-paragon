# 📊 Workflow Coordinator Scripts – Paragon

**CREDIT:** *Joshua van Niekerk and ChatGPT ~ Team Effort*

---

## 📘 Description

This project consists of a set of custom VBA scripts designed to assist with daily planning and decision-making in the **Workflow Coordinator role at Paragon**.

The scripts were collaboratively developed using ChatGPT to streamline and automate complex Excel-based processes that are part of the coordinator’s daily operations.

---

## ⚙️ What the Scripts Do

### **1. `MergeMySheets`**

- Collects relevant data from multiple predefined worksheets (e.g. `s1`, `s2`, ..., `s8`)
- Combines them into a single new worksheet named **`special`**
- Once merged, it calls the main script: `FilterDataAndCreateSummary`

---

### **2. `FilterDataAndCreateSummary`** *(Main method)*

Performs a comprehensive set of tasks to clean, analyze, and present the data:

#### ✅ Step-by-step breakdown:
1. **Export desired columns**  
   → Creates a new worksheet called `FilteredData` with only relevant fields.

2. **Match `OUTER` values based on `CORP_CD`**  
   → Uses a reference sheet (`outerskey`) to classify and map data.

3. **Highlight high insert counts**  
   → Flags work orders in **red** where `INSERT_CNT > 4`.

4. **Highlight critical outers**  
   → Flags specific work orders in **orange** for outers we always need to order (even with zero inserts).

5. **Highlight remakes**  
   → Flags rows in **yellow** if there's a remake count (`REM_MC_CNT` present).

6. **Generate a summary table**  
   → Aggregates total counts by `OUTER` and maps stock locations.

7. **Compare against previous list** *(if available)*  
   → If a `previous` worksheet exists, compares it to highlight **new entries** in **blue**.

8. **Create enclosed work order list**  
   → Identifies any `WORK_UNIT_CD`s that were in the `previous` list but not in the current one — these are added to the bottom of the report.

9. **Cleanup**  
   → Deletes both `special` and `previous` worksheets at the end of the run to avoid clutter.

---

## 🧠 Notes
- Designed for Excel and maintained in VBA (`.bas` file or directly inside the workbook).
- All outputs are visually formatted for print-readiness (e.g., borders, colors, headers, merged cells).
- Date/time stamp added to the header for quick reference.

---

## 🏁 Getting Started

To run the process:
1. Ensure the source sheets (`s1` to `s8`), `outerskey`, and optional `previous` sheet are present.
2. Run the macro: `MergeMySheets`
3. Let the automation handle the rest 🚀
