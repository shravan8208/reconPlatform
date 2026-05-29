"""
Operation Registry: Maps all 18 technical operations to plain-English labels,
categories, descriptions, and metadata for non-technical users.
"""

OPERATION_REGISTRY = {
    # =========================================================================
    # Category: Match & Reconcile
    # =========================================================================
    "vlookup": {
        "user_label": "Look Up Matching Values",
        "category": "Match & Reconcile",
        "icon": "search",
        "description": "Find values in one file that match values in another file",
        "when_to_use": "When you need to pull data from a reference table based on a common field (like account number, ID, etc.)",
        "difficulty": "standard",
    },
    "advance_vlookup": {
        "user_label": "Conditional Lookup",
        "category": "Match & Reconcile",
        "icon": "search",
        "description": "Look up values only when a condition is met (IF-THEN lookup)",
        "when_to_use": "When you need to match data but only for certain rows (e.g., only Equity types)",
        "difficulty": "advanced",
    },
    "sumifs": {
        "user_label": "Sum Values by Category",
        "category": "Match & Reconcile",
        "icon": "bar_chart",
        "description": "Add up values grouped by one or more conditions",
        "when_to_use": "When you need totals by category, department, account, etc.",
        "difficulty": "standard",
    },

    # =========================================================================
    # Category: Combine Data
    # =========================================================================
    "import": {
        "user_label": "Copy Columns Between Files",
        "category": "Combine Data",
        "icon": "file_copy",
        "description": "Copy selected columns from one file into another, mapping column names",
        "when_to_use": "When you need to bring specific columns from a source file into a target file",
        "difficulty": "standard",
    },
    "append": {
        "user_label": "Stack Rows from Two Files",
        "category": "Combine Data",
        "icon": "playlist_add",
        "description": "Add rows from one file to the bottom of another file",
        "when_to_use": "When you have the same data split across multiple files and want to combine them",
        "difficulty": "easy",
    },
    "header_map": {
        "user_label": "Rename Column Headers",
        "category": "Combine Data",
        "icon": "edit",
        "description": "Rename column headers to match a standard naming convention",
        "when_to_use": "When two files use different names for the same columns",
        "difficulty": "easy",
    },
    "copy_paste": {
        "user_label": "Copy Column Values",
        "category": "Combine Data",
        "icon": "content_copy",
        "description": "Copy values from one column into one or more other columns",
        "when_to_use": "When you need to duplicate a column's data into other columns",
        "difficulty": "easy",
    },

    # =========================================================================
    # Category: Clean & Transform
    # =========================================================================
    "formula": {
        "user_label": "Add Calculated Column",
        "category": "Clean & Transform",
        "icon": "calculate",
        "description": "Create a new column using a formula (works like Excel formulas)",
        "when_to_use": "When you need to calculate values based on other columns (e.g., Amount = Qty x Price)",
        "difficulty": "standard",
    },
    "replace": {
        "user_label": "Find & Replace Values",
        "category": "Clean & Transform",
        "icon": "find_replace",
        "description": "Find specific values and replace them with new ones",
        "when_to_use": "When you need to fix or standardize values (e.g., replace 'NY' with 'New York')",
        "difficulty": "easy",
    },
    "conditional": {
        "user_label": "Write Value When Condition Met",
        "category": "Clean & Transform",
        "icon": "rule",
        "description": "Set a column's value based on a rule (IF this THEN that)",
        "when_to_use": "When you need to flag or categorize rows based on conditions",
        "difficulty": "standard",
    },
    "forward_fill": {
        "user_label": "Fill Down Empty Cells",
        "category": "Clean & Transform",
        "icon": "vertical_align_bottom",
        "description": "Fill blank cells with the value from the cell above them",
        "when_to_use": "When your data has merged cells or grouped rows with blank values",
        "difficulty": "easy",
    },
    "convert_values": {
        "user_label": "Convert Formulas to Fixed Values",
        "category": "Clean & Transform",
        "icon": "lock",
        "description": "Replace formulas with their calculated values (freeze the results)",
        "when_to_use": "After using formulas, when you want to lock in the results",
        "difficulty": "easy",
    },
    "write_cell": {
        "user_label": "Write a Value to Cells",
        "category": "Clean & Transform",
        "icon": "edit_note",
        "description": "Write a specific value to a cell or fill an entire column with the same value",
        "when_to_use": "When you need to set a fixed value (e.g., write today's date, a status label, etc.)",
        "difficulty": "easy",
    },

    # =========================================================================
    # Category: Format & Report
    # =========================================================================
    "pivot": {
        "user_label": "Pivot Table Builder",
        "category": "Format & Report",
        "icon": "pivot_table_chart",
        "description": "Build a pivot table from source data — choose row fields, column fields, value fields and aggregation functions, then write the result to an output sheet",
        "when_to_use": "When you need to summarise and cross-tabulate data — totals by category, averages by month, counts by region, etc. The pivot result is written as clean data to a new sheet and can be re-run on any new data.",
        "difficulty": "standard",
    },
    "split_export": {
        "user_label": "Split Export",
        "category": "Format & Report",
        "icon": "call_split",
        "description": "Slice a source sheet and export to one file/sheet, multiple sheets, multiple files, or two-level file+sheet splits",
        "when_to_use": "When you need to distribute data into separate files or sheets by category — e.g. one file per client, one sheet per month",
        "difficulty": "standard",
    },
    "sheet_updater": {
        "user_label": "Sheet Updater",
        "category": "Format & Report",
        "icon": "system_update_alt",
        "description": "Push values from a master sheet into a multi-sheet target workbook — match sheets by category name, find the right row by particulars, and write the value",
        "when_to_use": "When you have a master list of (category, label, value) and need to update those values in matching sheets of an output file without overwriting anything else",
        "difficulty": "standard",
    },
    "format": {
        "user_label": "Format & Beautify",
        "category": "Format & Report",
        "icon": "palette",
        "description": "Apply bold, italic, colors, borders, and other formatting — statically to ranges or conditionally based on data values",
        "when_to_use": "After data is ready — apply colors, borders, bold headers, conditional highlights (e.g. red for negatives, green for positives) for final report presentation",
        "difficulty": "standard",
    },

    # =========================================================================
    # Category: Organize
    # =========================================================================
    "filter": {
        "user_label": "Keep Only Matching Rows",
        "category": "Organize",
        "icon": "filter_list",
        "description": "Remove rows that don't match your criteria",
        "when_to_use": "When you want to keep only certain rows (e.g., only 'Active' accounts)",
        "difficulty": "easy",
    },
    "sort": {
        "user_label": "Sort Rows",
        "category": "Organize",
        "icon": "sort",
        "description": "Sort your data by one or more columns",
        "when_to_use": "When you want to arrange rows in a specific order",
        "difficulty": "easy",
    },
    "insert": {
        "user_label": "Insert Blank Rows or Columns",
        "category": "Organize",
        "icon": "add_box",
        "description": "Add empty rows or columns at a specific position",
        "when_to_use": "When you need to make space for new data",
        "difficulty": "easy",
    },
    "delete": {
        "user_label": "Remove Rows or Columns",
        "category": "Organize",
        "icon": "delete",
        "description": "Delete rows, columns, or clear data from columns",
        "when_to_use": "When you need to remove unwanted data",
        "difficulty": "easy",
    },
    "merge_cells": {
        "user_label": "Merge Cells Together",
        "category": "Organize",
        "icon": "merge_type",
        "description": "Combine multiple cells into one (or split merged cells)",
        "when_to_use": "When you need to format cells for reporting",
        "difficulty": "easy",
    },
    "delete_sheets": {
        "user_label": "Delete Sheets",
        "category": "Organize",
        "icon": "tab_unselected",
        "description": "Remove one or more sheets from a workbook — by name, by name pattern, keep-only, or delete empty sheets",
        "when_to_use": "When you need to clean up a workbook by removing temp sheets, helper sheets, or sheets that match a naming pattern",
        "difficulty": "easy",
    },

    # =========================================================================
    # Category: Load Files
    # =========================================================================
    "normalize_import": {
        "user_label": "Normalize Headers & Import",
        "category": "Load Files",
        "icon": "sync_alt",
        "description": "Map messy input headers to standard names, then append the clean data to a working file. Each input class can have its own mapping.",
        "when_to_use": "After collecting input files — translate different header names (e.g. 'Txn Dt', 'Date', 'Value Date') into one standard name ('transaction_date').",
        "difficulty": "standard",
    },
    "input_source": {
        "user_label": "Input Source / File Collector",
        "category": "Load Files",
        "icon": "folder_open",
        "description": "Collect files from single paths, folders, filename rules, or dynamic date patterns. Assign named classes (e.g. 'Bank Statement') so later steps can reference them by name.",
        "when_to_use": "Use this as your first step to define and validate all input files before any processing begins.",
        "difficulty": "easy",
    },
    "add_files": {
        "user_label": "Add Files to Workspace",
        "category": "Load Files",
        "icon": "folder_open",
        "description": "Register files so they can be used in workflow steps",
        "when_to_use": "Start here - add the files you want to work with",
        "difficulty": "easy",
    },
    "data_input": {
        "user_label": "Preview File Data",
        "category": "Load Files",
        "icon": "preview",
        "description": "Load and preview the contents of a file",
        "when_to_use": "When you want to check what's in a file before processing it",
        "difficulty": "easy",
    },
}

# Category display order and descriptions
CATEGORIES = {
    "Load Files": {
        "icon": "folder_open",
        "description": "Add files to your workspace",
        "order": 0,
    },
    "Match & Reconcile": {
        "icon": "compare_arrows",
        "description": "Compare two files and find matching or missing items",
        "order": 1,
    },
    "Combine Data": {
        "icon": "merge",
        "description": "Bring data from multiple files together",
        "order": 2,
    },
    "Clean & Transform": {
        "icon": "build",
        "description": "Fix, replace, or calculate values in your data",
        "order": 3,
    },
    "Organize": {
        "icon": "sort",
        "description": "Filter, sort, or restructure your data",
        "order": 4,
    },
    "Format & Report": {
        "icon": "palette",
        "description": "Apply formatting and styling for final reports",
        "order": 5,
    },
}


def get_operations_by_category():
    """Return operations grouped by category, sorted by category order."""
    categories = {}
    for op_type, meta in OPERATION_REGISTRY.items():
        cat = meta["category"]
        if cat not in categories:
            categories[cat] = []
        categories[cat].append({"type": op_type, **meta})

    # Sort by category order
    sorted_cats = {}
    for cat_name in sorted(categories.keys(), key=lambda c: CATEGORIES.get(c, {}).get("order", 99)):
        sorted_cats[cat_name] = categories[cat_name]
    return sorted_cats


def get_user_label(op_type):
    """Get the plain-English label for an operation type."""
    return OPERATION_REGISTRY.get(op_type, {}).get("user_label", op_type)


def get_category_info(category_name):
    """Get display metadata for a category."""
    return CATEGORIES.get(category_name, {"icon": "help", "description": "", "order": 99})


def get_operation_info(op_type):
    """Get full metadata for an operation."""
    return OPERATION_REGISTRY.get(op_type, {})
