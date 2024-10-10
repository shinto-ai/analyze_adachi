import pandas as pd
import numpy as np
import os

def process_excel_file(file_path):
    """Process a single Excel file and return the sum of all non-empty sheets."""
    xl = pd.ExcelFile(file_path)
    total_df = None

    for sheet_name in xl.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        
        # Skip empty sheets
        if df.empty or df.isnull().all().all():
            continue
        
        # Remove index row and column
        df = df.iloc[1:, 1:]
        
        # Convert to numeric, replacing non-numeric values with 0
        df = df.apply(pd.to_numeric, errors='coerce').fillna(0)
        
        if total_df is None:
            total_df = df
        else:
            total_df += df

    return total_df

def create_cross_tabulation(df):
    """Create a cross-tabulation table with row and column totals."""
    row_totals = df.sum(axis=1)
    col_totals = df.sum(axis=0)
    grand_total = df.sum().sum()

    # Add row totals
    df['Row Total'] = row_totals

    # Add column totals
    col_totals_with_grand = pd.concat([col_totals, pd.Series({'Row Total': grand_total})])
    df.loc['Column Total'] = col_totals_with_grand

    return df

def calculate_expected_frequencies(observed_df):
    """Calculate expected frequencies for each cell."""
    total = observed_df.loc['Column Total', 'Row Total']
    row_totals = observed_df['Row Total'].drop('Column Total')
    col_totals = observed_df.loc['Column Total'].drop('Row Total')

    expected_df = observed_df.copy()
    for i in row_totals.index:
        for j in col_totals.index:
            expected_df.loc[i, j] = (row_totals[i] * col_totals[j]) / total

    return expected_df

def calculate_chi_square(observed_df, expected_df):
    """Calculate chi-square value."""
    observed = observed_df.drop('Column Total').drop('Row Total', axis=1)
    expected = expected_df.drop('Column Total').drop('Row Total', axis=1)
    return np.sum((observed - expected)**2 / expected)

def main():
    file_paths = ['HU_data.xlsx', 'Otani_data.xlsx']
    results = {}

    for file_path in file_paths:
        group_name = os.path.splitext(file_path)[0].split('_')[0]
        df = process_excel_file(file_path)
        cross_tab = create_cross_tabulation(df)
        expected_freq = calculate_expected_frequencies(cross_tab)
        chi_square = calculate_chi_square(cross_tab, expected_freq)

        results[group_name] = {
            'cross_tab': cross_tab,
            'expected_freq': expected_freq,
            'chi_square': chi_square
        }

        # Save results to Excel files
        cross_tab.to_excel(f'{group_name}_cross_tabulation.xlsx')
        expected_freq.to_excel(f'{group_name}_expected_frequencies.xlsx')

    print("Chi-square values:")
    for group, data in results.items():
        print(f"{group}: {data['chi_square']}")

if __name__ == "__main__":
    main()