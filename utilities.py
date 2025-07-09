import pandas as pd
import string

# Utility: Clean and normalize patient/facility identifiers
def clean_id(val):
    if pd.isna(val):
        return ''
    val = str(val).strip().lower().replace(' ', '').lstrip('0')
    return val

def process_emr_data(df, dfbaseline, emr_df):
    # Remove rows with any blank fields in mapping
    emr_df = emr_df[(emr_df != '').all(axis=1)]
    
    # Select and deduplicate necessary columns from emr_df
    emr_subset = emr_df[['Name on NMRS', 'LGA', 'STATE', 'Name on Lamis']].drop_duplicates(subset='Name on NMRS')

    # Merge once using FacilityName <-> Name on NMRS
    df = df.merge(
        emr_subset,
        how='left',
        left_on='FacilityName',
        right_on='Name on NMRS',
        suffixes=('', '_emr')
    )

    # Fill missing LGA and State from EMR
    df['LGA'] = df['LGA'].fillna(df['LGA_emr'])
    df['State'] = df['State'].fillna(df['STATE'])

    # Replace FacilityName if different
    df.loc[df['Name on Lamis'] != df['FacilityName'], 'FacilityName'] = df['Name on Lamis']

    # Drop extra columns
    df.drop(['Name on NMRS', 'LGA_emr', 'STATE', 'Name on Lamis'], axis=1, inplace=True)


    # Normalize hospital numbers and unique IDs
    df['PatientHospitalNo1'] = df['PatientHospitalNo'].apply(clean_id)
    df['PatientUniqueID1'] = df['PEPID'].apply(clean_id)
    dfbaseline['Hospital Number1'] = dfbaseline['Hospital Number'].apply(clean_id)
    dfbaseline['Unique ID1'] = dfbaseline['Unique ID'].apply(clean_id)

    # Create consistent unique identifiers for both datasets
    dfbaseline['unique identifiers'] = (
        dfbaseline["LGA"].astype(str).str.lower().str.strip().str.replace(' ', '') +
        dfbaseline["Facility"].astype(str).str.lower().str.strip().str.replace(' ', '') +
        dfbaseline["Hospital Number1"] +
        dfbaseline["Unique ID1"]
    )

    df['unique identifiers'] = (
        df["LGA"].astype(str).str.lower().str.strip().str.replace(' ', '') +
        df["FacilityName"].astype(str).str.lower().str.strip().str.replace(' ', '') +
        df["PatientHospitalNo1"] +
        df["PatientUniqueID1"]
    )

    # Drop duplicates from baseline data
    dfbaseline = dfbaseline.drop_duplicates(subset=['unique identifiers'], keep=False)

    # Handle duplicate IDs in current df by appending _A, _B, etc.
    #dup_mask = df.duplicated(['unique identifiers'], keep=False)
    #alphabet = dict(enumerate(string.ascii_uppercase))
    #df.loc[dup_mask, 'unique identifiers'] += '_' + df[dup_mask].groupby(['unique identifiers']).cumcount().map(alphabet)

    # Merge into df
    df = df.merge(
        dfbaseline[['unique identifiers', 'Date of TPT Start (yyyy-mm-dd)', 'TPT Type']],
        on='unique identifiers',
        how='left',
        suffixes=('', '_baseline')
    )
    df.to_excel('df.xlsx')

    # Fill missing TPT values
    df['Date of TPT Start (yyyy-mm-dd)'] = pd.to_datetime(df['Date of TPT Start (yyyy-mm-dd)'], errors='coerce', dayfirst=True)
    df['First_TPT_Pickupdate'] = pd.to_datetime(df['First_TPT_Pickupdate'], errors='coerce', dayfirst=True)
    df['First_TPT_Pickupdate'] = df['First_TPT_Pickupdate'].fillna(df['Date of TPT Start (yyyy-mm-dd)'])
    df['Current_TPT_Received'] = df['Current_TPT_Received'].fillna(df['TPT Type'])

    return df

"""
# Utility: Parse individual date
def parse_date(date):
    
    # Handle if date is a list or array: try first element or return NaT if empty
    if isinstance(date, (list, tuple, np.ndarray)):
        if len(date) > 0:
            date = date[0]
        else:
            return pd.NaT

    if pd.isna(date):  # Handle NaN values
        return pd.NaT
    
    if isinstance(date, pd.Timestamp):  # If already a datetime object
        return date if 1900 <= date.year <= 2099 else pd.NaT
    
    if isinstance(date, (int, float)):  # Handle Excel serial numbers
        try:
            return pd.to_datetime(date, origin='1899-12-30', unit='D')
        except Exception:
            return pd.NaT

    date_formats = ["%Y-%m-%d", "%d/%m/%Y", "%m-%d-%Y", "%d-%m-%Y", "%Y.%m.%d", "%Y-%b-%d"]

    for fmt in date_formats:
        try:
            return pd.to_datetime(date, format=fmt)
        except (ValueError, TypeError):
            continue

    try:
        return parser.parse(str(date), fuzzy=True, ignoretz=True)
    except (parser.ParserError, ValueError, TypeError):
        return pd.NaT
"""

"""
# Utility: clean dates and numbers
def clean_dataframe(df):
    for col in DATE_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(parse_date)
    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    return df
"""