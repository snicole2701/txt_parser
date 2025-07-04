import pandas as pd

def parse_12_month_txt(filepath):
    # Read raw lines into a DataFrame
    with open(filepath, 'r', encoding='utf-8') as file:
        lines = [line.rstrip() for line in file.readlines()]
    
    df = pd.DataFrame(lines, columns=["raw"])

    # Ensure each line is exactly 182 characters long
    df["raw"] = df["raw"].apply(lambda x: x.ljust(182))

    # Remove unwanted lines
    df = df[~df["raw"].str.contains("ANALYSIS REPORT", case=False, na=False)]
    df = df[~df["raw"].str.strip().eq("")]

    # Keep only the first "MONTHS" line
    months_mask = df["raw"].str.startswith("MONTHS")
    first_months_index = months_mask.idxmax() if months_mask.any() else None
    if first_months_index is not None:
        df = df.loc[~(months_mask & (df.index != first_months_index))]

    # Insert placeholder column for Employee Number
    df.insert(0, "Employee Number", "")

    # Split lines into columns using your custom logic
    split_columns = []
    for line in df["raw"]:
        last_col = line[-17:]
        remaining = line[:-17]

        cols_13 = []
        for _ in range(11):
            cols_13.insert(0, remaining[-13:])
            remaining = remaining[:-13]

        id_segment = remaining[-10:]
        label = remaining[:-10].strip()

        split_columns.append([label, id_segment] + cols_13 + [last_col])

    # Replace with split columns as a new DataFrame
    df = df.reset_index(drop=True)
    split_df = pd.DataFrame(split_columns)
    split_df.insert(0, "Employee Number", df["Employee Number"])

    # Copy "Employee Number" into first data cell and promote second row to header
    if split_df.columns[0] == "Employee Number":
        split_df.iloc[0, 0] = "Employee Number"

    new_header = split_df.iloc[0]
    split_df = split_df[1:].copy()
    split_df.columns = new_header

    # ▶ Fill empty fields with "$" in rows where Description starts with "*"
    for i in range(len(split_df)):
        if str(split_df.iloc[i, 1]).strip().startswith("*"):
            for j in range(split_df.shape[1]):
                if str(split_df.iat[i, j]).strip() == "":
                    split_df.iat[i, j] = "$"


    # Propagate Employee Number from logic based on column 6 and column 2
    for i in range(len(split_df)):
        col_5 = str(split_df.iloc[i, 5]).strip()
        col_1 = str(split_df.iloc[i, 1]).strip()

        if col_5 == "" and col_1 != "*FP FLAG *":
            current_id = col_1[:10]

        if current_id:
            split_df.iat[i, 0] = current_id
            
    # Remove rows where 6th field is blank and 2nd field is not "*FP FLAG *"
    condition = (split_df.iloc[:, 5].astype(str).str.strip() == "") & \
                (split_df.iloc[:, 1].astype(str).str.strip() != "*FP FLAG *")

    split_df = split_df[~condition].reset_index(drop=True)

    # Insert blank column in third
    split_df.insert(2, "CATEGORY TAG", "")
    
        # Identify rows with "*FP FLAG *" to avoid tagging past them
    fp_flag_indices = split_df[split_df.iloc[:, 1].astype(str).str.strip() == "*FP FLAG *"].index.tolist()

    # Column positions
    desc_col = 1  # Description
    tag_col = 2   # Temp/Category Tag
    emp_col = 0   # Employee Number

    # Start bottom-up tagging
    category = None
    previous_emp = None
    max_time_block = 500

    for i in reversed(range(len(split_df))):
        emp = split_df.iat[i, emp_col]
        desc = str(split_df.iat[i, desc_col]).strip()

        # On employee change, walk upward and tag "Hours" until FP FLAG or end
        if previous_emp and emp != previous_emp:
            j = i
            steps = 0
            while j >= 0 and steps < max_time_block:
                prior_emp = split_df.iat[j, emp_col]
                prior_desc = str(split_df.iat[j, desc_col]).strip()
                if prior_desc == "*FP FLAG *":
                    break
                if prior_emp == emp and not prior_desc.startswith("*"):
                    split_df.iat[j, tag_col] = "TIME"
                j -= 1
                steps += 1
            category = None
        previous_emp = emp

        # Skip over "*FP FLAG *" rows cleanly
        if i in fp_flag_indices:
            category = None
            continue

        # If row is a new "*..." marker, tag it and capture the category
        if desc.startswith("*"):
            cleaned = desc.replace("*", "").replace("TOT:", "").strip()
            category = cleaned
            split_df.iat[i, tag_col] = cleaned
        elif category and not desc.startswith("*"):
            split_df.iat[i, tag_col] = category
            
    # Rename column 2 and convert all headers to Title Case
    split_df.columns = [col.strip().title() for col in split_df.columns]

    # Keep all existing column names, but rename only the second one to "Description"
    split_df.columns = [
        "Description" if i == 1 else col
        for i, col in enumerate(split_df.columns)
    ]

    # Replace all occurrences of '.00' (as string) with '0' across the entire DataFrame
    split_df = split_df.apply(lambda col: col.map(lambda val: "0" if str(val).strip() == ".00" else val))

    # ▶ Clean up all "$" markers back to empty strings
    split_df = split_df.apply(lambda col: col.map(lambda val: "" if str(val).strip() == "$" else val))

    return split_df.reset_index(drop=True)

if __name__ == "__main__":
    import os

    folder = os.path.dirname(__file__)
    filepath = os.path.join(folder, "Co01 12M Report.txt")
    df = parse_12_month_txt(filepath)

    df.to_csv(os.path.join(folder, "outcome.csv"), index=False)
    df.to_excel(os.path.join(folder, "outcome.xlsx"), index=False)

    print("✅ Parsing complete. Files saved as outcome.csv and outcome.xlsx.")