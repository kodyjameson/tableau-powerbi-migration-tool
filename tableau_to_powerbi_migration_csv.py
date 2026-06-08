import os
import re
import csv
import zipfile
import shutil
import xml.etree.ElementTree as ET
from collections import defaultdict, Counter

OUTPUT_FOLDER = "tableau_migration_csv_output"

BAD_WORDS = {
    "sum", "none", "usr", "qk", "nk", "ok", "pcto",
    "Parameters", "Measure Names", "Multiple Values"
}

RISK_PATTERNS = {
    "WINDOW_": "High",
    "INDEX()": "High",
    "LOOKUP(": "High",
    "PREVIOUS_VALUE": "High",
    "{FIXED": "High",
    "{INCLUDE": "High",
    "{EXCLUDE": "High",
    "RANK(": "Medium",
    "DATEPARSE": "Medium",
    "REGEXP": "Medium",
    "ZN(": "Medium",
    "TOTAL(": "Medium"
}


def clean(value):
    if not value:
        return ""

    value = value.replace("[", "").replace("]", "").replace('"', "").strip()
    value = value.split(".")[-1]

    parts = value.split(":")
    useful = []

    for part in parts:
        if part and part not in BAD_WORDS:
            useful.append(part)

    if useful:
        value = useful[-1]

    if value in BAD_WORDS:
        return ""

    return value.strip()


def safe_folder_name(path):
    name = os.path.basename(path)
    name = name.replace(".twbx", "").replace(".twb", "")
    name = re.sub(r"[^A-Za-z0-9_-]+", "_", name)
    return name.strip("_") or "tableau_workbook"


def extract_twb(input_file):
    if input_file.lower().endswith(".twb"):
        return input_file

    if not input_file.lower().endswith(".twbx"):
        raise Exception("Input must be a .twb or .twbx file.")

    extract_folder = "extracted_twbx_temp"

    if os.path.exists(extract_folder):
        shutil.rmtree(extract_folder)

    os.makedirs(extract_folder)

    with zipfile.ZipFile(input_file, "r") as z:
        z.extractall(extract_folder)

    for root_dir, dirs, files in os.walk(extract_folder):
        for filename in files:
            if filename.lower().endswith(".twb"):
                return os.path.join(root_dir, filename)

    raise Exception("No .twb file found inside the .twbx package.")


def write_csv(folder, filename, headers, rows):
    path = os.path.join(folder, filename)

    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)


def complexity_label(score):
    if score <= 10:
        return "Low"
    if score <= 25:
        return "Medium"
    return "High"


def main():
    input_file = input(
        "Drag the Tableau .twb or .twbx file here, then press Enter:\n"
    ).strip().strip("'").strip('"').replace("\\ ", " ")

    if not os.path.exists(input_file):
        print("File not found:")
        print(input_file)
        return

    workbook_name = safe_folder_name(input_file)
    output_folder = os.path.join(OUTPUT_FOLDER, workbook_name)

    if os.path.exists(output_folder):
        shutil.rmtree(output_folder)

    os.makedirs(output_folder)

    twb_path = extract_twb(input_file)

    tree = ET.parse(twb_path)
    root = tree.getroot()

    dashboards = set()
    worksheets = set()
    datasources = set()

    all_fields = []
    field_source = {}

    calc_lookup = {}
    calc_rows = []

    dashboard_map = []
    roadmap = []

    parameters = []
    custom_sql = set()
    risks = set()

    # Data sources, fields, calculations, parameters
    for datasource in root.findall(".//datasource"):
        ds_name = datasource.get("caption") or datasource.get("name", "")

        if ds_name:
            datasources.add(ds_name)

        for column in datasource.findall(".//column"):
            field_name = clean(column.get("caption", "")) or clean(column.get("name", ""))
            raw_name = clean(column.get("name", ""))

            dtype = column.get("datatype", "")
            role = column.get("role", "")
            value = column.get("value", "")
            param_type = column.get("param-domain-type", "")

            if field_name:
                all_fields.append([ds_name, field_name, dtype, role])
                field_source[field_name] = ds_name

            if param_type:
                range_node = column.find("range")

                parameters.append([
                    field_name,
                    dtype,
                    value,
                    range_node.get("min", "") if range_node is not None else "",
                    range_node.get("max", "") if range_node is not None else "",
                    range_node.get("granularity", "") if range_node is not None else ""
                ])

            calc = column.find("calculation")

            if calc is not None:
                formula = calc.get("formula", "")
                calc_name = field_name or raw_name

                calc_lookup[raw_name] = (calc_name, formula, ds_name)
                calc_lookup[calc_name] = (calc_name, formula, ds_name)

                calc_rows.append([ds_name, calc_name, formula])

    # Custom SQL
    for relation in root.findall(".//relation"):
        if relation.get("type") == "text" and relation.text:
            custom_sql.add(relation.text.strip())

    # Dashboards and worksheets
    for dashboard in root.findall(".//dashboard"):
        dashboard_name = dashboard.get("name", "")

        if dashboard_name:
            dashboards.add(dashboard_name)

        seen = set()

        for zone in dashboard.findall(".//zone"):
            worksheet_name = zone.get("name", "")

            if worksheet_name and worksheet_name not in seen:
                seen.add(worksheet_name)
                worksheets.add(worksheet_name)
                dashboard_map.append([dashboard_name, worksheet_name])

    # Worksheet dependencies
    for worksheet in root.findall(".//worksheet"):
        worksheet_name = worksheet.get("name", "")

        if worksheet_name:
            worksheets.add(worksheet_name)

        found = set()

        for elem in worksheet.iter():
            for attr_value in elem.attrib.values():
                matches = re.findall(r"\[[^\]]+\]", attr_value)

                for match in matches:
                    name = clean(match)

                    if not name:
                        continue

                    if name in found:
                        continue

                    if name.startswith("Calculation_") and name not in calc_lookup:
                        continue

                    found.add(name)

                    if name in calc_lookup:
                        calc_name, formula, ds_name = calc_lookup[name]
                        roadmap.append([
                            worksheet_name,
                            "Calculation",
                            calc_name,
                            ds_name,
                            formula
                        ])
                    else:
                        roadmap.append([
                            worksheet_name,
                            "Field",
                            name,
                            field_source.get(name, ""),
                            ""
                        ])

    # Full roadmap with dashboard
    dash_to_ws = defaultdict(set)

    for dash, ws_name in dashboard_map:
        dash_to_ws[dash].add(ws_name)

    ws_to_deps = defaultdict(list)

    for ws_name, dep_type, dep_name, ds_name, formula in roadmap:
        ws_to_deps[ws_name].append((dep_type, dep_name, ds_name, formula))

    full_roadmap = []

    for dash, ws_set in dash_to_ws.items():
        for ws_name in sorted(ws_set):
            for dep_type, dep_name, ds_name, formula in ws_to_deps.get(ws_name, []):
                full_roadmap.append([
                    dash,
                    ws_name,
                    dep_type,
                    dep_name,
                    ds_name,
                    formula
                ])

    # Migration summary
    summary_rows = []

    for dash, ws_set in dash_to_ws.items():
        field_set = set()
        calc_set = set()
        source_set = set()

        for ws_name in ws_set:
            for dep_type, dep_name, ds_name, formula in ws_to_deps.get(ws_name, []):
                if ds_name:
                    source_set.add(ds_name)

                if dep_type == "Field":
                    field_set.add(dep_name)

                if dep_type == "Calculation":
                    calc_set.add(dep_name)

        score = len(ws_set) + len(field_set) + (len(calc_set) * 3)

        summary_rows.append([
            dash,
            len(ws_set),
            len(field_set),
            len(calc_set),
            len(source_set),
            score,
            complexity_label(score)
        ])

    # Calculation dependencies
    calc_dep_rows = []

    for ds_name, calc_name, formula in calc_rows:
        matches = re.findall(r"\[[^\]]+\]", formula)

        used = sorted(set(clean(x) for x in matches if clean(x)))

        uses_calcs = []
        uses_fields = []

        for item in used:
            if item in calc_lookup or item.startswith("Calculation_"):
                uses_calcs.append(item)
            else:
                uses_fields.append(item)

        calc_dep_rows.append([
            calc_name,
            ds_name,
            ", ".join(uses_fields),
            ", ".join(uses_calcs),
            formula
        ])

    # Field and calculation usage
    field_counter = Counter()
    calc_counter = Counter()

    for ws_name, dep_type, dep_name, ds_name, formula in roadmap:
        if dep_type == "Field":
            field_counter[dep_name] += 1

        if dep_type == "Calculation":
            calc_counter[dep_name] += 1

    field_usage_rows = []

    for field, count in field_counter.most_common():
        field_usage_rows.append([
            field,
            field_source.get(field, ""),
            count
        ])

    calc_usage_rows = []

    for calc_name, count in calc_counter.most_common():
        calc_usage_rows.append([
            calc_name,
            count
        ])

    # Worksheet complexity
    worksheet_complexity_rows = []

    for ws_name, deps in sorted(ws_to_deps.items()):
        fields = set()
        calcs = set()

        for dep_type, dep_name, ds_name, formula in deps:
            if dep_type == "Field":
                fields.add(dep_name)

            if dep_type == "Calculation":
                calcs.add(dep_name)

        score = len(fields) + (len(calcs) * 3)

        worksheet_complexity_rows.append([
            ws_name,
            len(fields),
            len(calcs),
            score,
            complexity_label(score)
        ])

    # Data source impact
    source_dash = defaultdict(set)
    source_ws = defaultdict(set)
    source_calc = defaultdict(set)

    for dash, ws_name, dep_type, dep_name, ds_name, formula in full_roadmap:
        if ds_name:
            source_dash[ds_name].add(dash)
            source_ws[ds_name].add(ws_name)

            if dep_type == "Calculation":
                source_calc[ds_name].add(dep_name)

    source_impact_rows = []

    for ds_name in sorted(datasources):
        source_impact_rows.append([
            ds_name,
            len(source_dash[ds_name]),
            len(source_ws[ds_name]),
            len(source_calc[ds_name])
        ])

    # Unused field check
    used_fields = set(field_counter.keys())

    unused_rows = []

    seen_field_rows = set()

    for ds_name, field_name, dtype, role in all_fields:
        key = (ds_name, field_name, dtype, role)

        if key in seen_field_rows:
            continue

        seen_field_rows.add(key)

        unused_rows.append([
            ds_name,
            field_name,
            dtype,
            role,
            "Yes" if field_name in used_fields else "No"
        ])

    # Migration checklist
    checklist_rows = []

    for dash, ws_set in dash_to_ws.items():
        for ws_name in sorted(ws_set):
            deps = ws_to_deps.get(ws_name, [])

            field_count = len([x for x in deps if x[0] == "Field"])
            calc_count = len([x for x in deps if x[0] == "Calculation"])
            score = field_count + (calc_count * 3)

            checklist_rows.append([
                dash,
                ws_name,
                field_count,
                calc_count,
                complexity_label(score),
                "Not Started",
                "",
                ""
            ])

    # Risks
    for ds_name, calc_name, formula in calc_rows:
        upper_formula = formula.upper()

        for pattern, severity in RISK_PATTERNS.items():
            if pattern in upper_formula:
                risks.add((pattern, severity, calc_name, formula))

    risk_pattern_counts = Counter()
    risk_severity_counts = Counter()

    for pattern, severity, calc_name, formula in risks:
        risk_pattern_counts[pattern] += 1
        risk_severity_counts[severity] += 1

    risk_summary_rows = []

    for severity, count in sorted(risk_severity_counts.items()):
        risk_summary_rows.append(["Severity", severity, count])

    for pattern, count in sorted(risk_pattern_counts.items()):
        risk_summary_rows.append(["Pattern", pattern, count])

    # Power BI build order and translation hints
    build_order_rows = [
        [1, "Confirm data sources", "Identify and connect the required source tables/files."],
        [2, "Load source fields", "Bring in all required fields used by the Tableau workbook."],
        [3, "Recreate parameters", "Create slicers or disconnected tables for Tableau parameters."],
        [4, "Create base measures", "Build simple measures first, such as SUM(Sales)."],
        [5, "Create calculated measures", "Recreate Tableau calculated fields in DAX."],
        [6, "Validate dependencies", "Check that calculated fields have their required base fields."],
        [7, "Build visuals", "Recreate each worksheet as a Power BI visual."],
        [8, "Build dashboard page", "Arrange visuals to match the Tableau dashboard layout."],
        [9, "Validate totals and filters", "Compare Tableau and Power BI numbers."],
        [10, "Document gaps", "Log unsupported Tableau logic or manual rebuild needs."]
    ]

    translation_hint_rows = [
        ["SUM([Field])", "SUM(Table[Field])", "Basic aggregation."],
        ["AVG([Field])", "AVERAGE(Table[Field])", "Basic average."],
        ["COUNTD([Field])", "DISTINCTCOUNT(Table[Field])", "Distinct count."],
        ["IF ... THEN ... ELSE ... END", "IF(condition, true_result, false_result)", "May require nested IF or SWITCH."],
        ["CASE ... WHEN ... THEN ... END", "SWITCH()", "Usually maps well to SWITCH."],
        ["ISNULL([Field])", "ISBLANK(Table[Field])", "Null check."],
        ["ZN([Field])", "COALESCE(Table[Field], 0)", "Replace null with zero."],
        ["YEAR([Date])", "YEAR(Table[Date])", "Date function."],
        ["MONTH([Date])", "MONTH(Table[Date])", "Date function."],
        ["WINDOW_SUM", "DAX measure with filter context", "Manual review likely."],
        ["WINDOW_MAX", "DAX measure with filter context", "Manual review likely."],
        ["INDEX()", "RANKX or custom sort logic", "Manual review likely."],
        ["TOTAL()", "CALCULATE with ALL/ALLSELECTED", "Depends on Tableau view context."],
        ["{FIXED ...}", "CALCULATE with ALLEXCEPT/REMOVEFILTERS", "Manual review likely."]
    ]

    workbook_summary_rows = [
        ["Workbook", os.path.basename(input_file)],
        ["Dashboards", len(dashboards)],
        ["Worksheets", len(worksheets)],
        ["Data Sources", len(datasources)],
        ["Fields", len(set(x[1] for x in all_fields))],
        ["Calculated Fields", len(calc_rows)],
        ["Parameters", len(parameters)],
        ["Custom SQL Queries", len(custom_sql)],
        ["Migration Risk Items", len(risks)],
        ["Roadmap Rows", len(full_roadmap)]
    ]

    # Write outputs
    write_csv(
        output_folder,
        "00_workbook_summary.csv",
        ["Metric", "Value"],
        workbook_summary_rows
    )

    write_csv(
        output_folder,
        "01_dashboard_summary.csv",
        ["Dashboard", "Worksheet Count", "Field Count", "Calculation Count", "Data Source Count", "Complexity Score", "Complexity"],
        summary_rows
    )

    write_csv(
        output_folder,
        "02_dashboard_map.csv",
        ["Dashboard", "Worksheet"],
        dashboard_map
    )

    write_csv(
        output_folder,
        "03_migration_roadmap.csv",
        ["Dashboard", "Worksheet", "Dependency Type", "Dependency Name", "Data Source", "Formula"],
        full_roadmap
    )

    write_csv(
        output_folder,
        "04_calculation_dictionary.csv",
        ["Data Source", "Calculation Name", "Formula"],
        calc_rows
    )

    write_csv(
        output_folder,
        "05_calculation_dependencies.csv",
        ["Calculation", "Data Source", "Uses Fields", "Uses Calculations", "Formula"],
        calc_dep_rows
    )

    write_csv(
        output_folder,
        "06_worksheet_complexity.csv",
        ["Worksheet", "Field Count", "Calculation Count", "Score", "Complexity"],
        worksheet_complexity_rows
    )

    write_csv(
        output_folder,
        "07_field_usage_frequency.csv",
        ["Field", "Data Source", "Worksheet Usage Count"],
        field_usage_rows
    )

    write_csv(
        output_folder,
        "08_calculation_usage_frequency.csv",
        ["Calculation", "Worksheet Usage Count"],
        calc_usage_rows
    )

    write_csv(
        output_folder,
        "09_data_source_impact.csv",
        ["Data Source", "Dashboard Count", "Worksheet Count", "Calculation Count"],
        source_impact_rows
    )

    write_csv(
        output_folder,
        "10_data_sources.csv",
        ["Data Source"],
        [[x] for x in sorted(datasources)]
    )

    write_csv(
        output_folder,
        "11_custom_sql.csv",
        ["SQL"],
        [[x] for x in sorted(custom_sql)]
    )

    write_csv(
        output_folder,
        "12_parameters.csv",
        ["Parameter", "Data Type", "Default Value", "Min", "Max", "Step"],
        parameters
    )

    write_csv(
        output_folder,
        "13_unused_field_check.csv",
        ["Data Source", "Field", "Data Type", "Role", "Used In Workbook"],
        unused_rows
    )

    write_csv(
        output_folder,
        "14_migration_risks.csv",
        ["Risk Pattern", "Severity", "Calculation", "Formula"],
        [list(x) for x in sorted(risks)]
    )

    write_csv(
        output_folder,
        "15_risk_summary.csv",
        ["Summary Type", "Item", "Count"],
        risk_summary_rows
    )

    write_csv(
        output_folder,
        "16_powerbi_build_order.csv",
        ["Step", "Build Item", "Purpose"],
        build_order_rows
    )

    write_csv(
        output_folder,
        "17_powerbi_translation_hints.csv",
        ["Tableau Pattern", "Power BI / DAX Direction", "Notes"],
        translation_hint_rows
    )

    write_csv(
        output_folder,
        "18_migration_checklist.csv",
        ["Dashboard", "Worksheet", "Field Count", "Calculation Count", "Complexity", "Status", "Assigned To", "Notes"],
        checklist_rows
    )

    print("")
    print("CSV migration package created:")
    print(output_folder)
    print("")
    print("Open the folder and review the CSV files in Excel.")


if __name__ == "__main__":
    main()
