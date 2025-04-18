import pandas as pd
from docx import Document

def update_metrics_in_word(excel_path, word_template_path, output_path):
    # Load Excel and normalize column names
    df = pd.read_excel("D://Automation//tests//test.xlsx")
    df = df.dropna(how='all')
    df.columns = df.columns.str.strip().str.lower()

    print("Excel Columns:", df.columns.tolist())

    try:
        app_count = df['application name'].dropna().shape[0]
        total_issues = df['total issues reviewed'].sum()
        actual_issues = df['issue'].sum() + df['probably not an issue'].sum()
        not_an_issue = df['not an issue'].sum()
    except KeyError as e:
        print(f"ERROR: Column not found in Excel: {e}")
        return

    # Map actual Word metric labels to calculated values
    label_to_logic = {
        "total applications reviewed": app_count,
        "total no.of findings triaged": total_issues,
        "total issues found": actual_issues,
        "total false positive": not_an_issue
    }

    print("Calculated Metrics:", label_to_logic)

    # Load Word document
    doc = Document("D://Automation//tests//metric.docx")

    # Find the target table with 'Metric' column
    target_table = None
    for table in doc.tables:
        if any(cell.text.strip().lower() == "metric" for cell in table.rows[0].cells):
            target_table = table
            break

    if not target_table:
        print("ERROR: No table with 'Metric' column found.")
        return

    # Identify column indices
    header_cells = target_table.rows[0].cells
    metric_idx = count_idx = None
    for i, cell in enumerate(header_cells):
        col_name = cell.text.strip().lower()
        if col_name == "metric":
            metric_idx = i
        elif col_name == "count":
            count_idx = i

    if metric_idx is None or count_idx is None:
        print("ERROR: Couldn't identify both 'Metric' and 'Count' columns.")
        return

    # Update table based on exact metric labels
    for row in target_table.rows[1:]:
        metric_text = row.cells[metric_idx].text.strip().lower()
        if metric_text in label_to_logic:
            value = label_to_logic[metric_text]
            row.cells[count_idx].text = str(int(value) if pd.notnull(value) else 0)
            print(f"✅ Updated '{metric_text}' with value: {value}")
        else:
            print(f"⚠️ Skipped unmatched metric: '{metric_text}'")

    doc.save("D://Automation//tests//out.docx")
    print(f"\n✅ Word document updated and saved as: {output_path}")


# === Example usage ===
update_metrics_in_word("test.xlsx", "metric.docx", "out.docx")
