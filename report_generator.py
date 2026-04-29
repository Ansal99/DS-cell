import pandas as pd


def generate_report(file_path):
    # Step 1: Read file without assuming header
    df_raw = pd.read_excel(file_path, header=None, engine='openpyxl')

    header_row = None
    for i in range(len(df_raw)):
        row = df_raw.iloc[i].astype(str).str.strip().tolist()
        if 'DPSU' in row and 'Equipment_Name' in row:
            header_row = i
            break

    if header_row is None:
        raise Exception('Header row not found. Check Excel format.')

    # Step 2: Read again with the correct header row
    df = pd.read_excel(file_path, header=header_row, engine='openpyxl')
    df.columns = df.columns.str.strip()

    grouped = df.groupby(['DPSU', 'Equipment_Name'])
    report_data = {}

    for (dpsu, equipment), group in grouped:
        total_codified = group['Received_Date'].notna().sum() if 'Received_Date' in group.columns else 0
        fwd_dca        = group['Forward_Date'].notna().sum()  if 'Forward_Date'  in group.columns else 0
        nsn_allotted   = group['NSN'].notna().sum()           if 'NSN'           in group.columns else 0
        returned       = group['Return_Date'].notna().sum()   if 'Return_Date'   in group.columns else 0

        if dpsu not in report_data:
            report_data[dpsu] = []

        report_data[dpsu].append({
            'Equipment':      equipment,
            'Total_Codified': int(total_codified),
            'Fwd_DCA':        int(fwd_dca),
            'NSN':            int(nsn_allotted),
            'Returned':       int(returned)
        })

    return report_data