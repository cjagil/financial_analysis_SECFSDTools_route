import pandas as pd

def process_input_data(input_file):
    # Load the full dataset
    df = pd.read_excel(input_file)
    
    # Separate data for each statement type using unique column names
    balance_sheet_df = df[['fy', 'form', 'name', 'Assets', 'Liabilities', 'Equity']].dropna(subset=['name'])
    income_statement_df = df[['fy', 'form', 'name', 'Revenues', 'OperatingIncomeLoss', 'IncomeLossFromContinuingOperationsBeforeIncomeTaxExpenseBenefit', 'AllIncomeTaxExpenseBenefit']].dropna(subset=['name'])
    cash_flow_df = df[['fy', 'form', 'name', 'NetCashProvidedByUsedInOperatingActivities', 'PaymentsToAcquirePropertyPlantAndEquipment', 'DepreciationDepletionAndAmortization']].dropna(subset=['name'])

    # Rename columns to avoid conflicts during merging
    balance_sheet_df = balance_sheet_df.rename(columns={'name': 'bs_name'})
    income_statement_df = income_statement_df.rename(columns={'name': 'is_name'})
    cash_flow_df = cash_flow_df.rename(columns={'name': 'cf_name'})

    # Aggregate data to ensure no duplicates before merging
    balance_sheet_df = balance_sheet_df.groupby(['fy', 'form', 'bs_name']).sum().reset_index()
    income_statement_df = income_statement_df.groupby(['fy', 'form', 'is_name']).sum().reset_index()
    cash_flow_df = cash_flow_df.groupby(['fy', 'form', 'cf_name']).sum().reset_index()

    # Merge the three dataframes on fiscal year ('fy'), form ('form'), and company name columns
    merged_df = pd.merge(balance_sheet_df, income_statement_df, how='inner', left_on=['fy', 'form', 'bs_name'], right_on=['fy', 'form', 'is_name'])
    merged_df = pd.merge(merged_df, cash_flow_df, how='inner', left_on=['fy', 'form', 'bs_name'], right_on=['fy', 'form', 'cf_name'])

    # Filter the data for Apple Inc.
    apple_df = merged_df[(merged_df['bs_name'] == 'APPLE INC') | (merged_df['is_name'] == 'APPLE INC') | (merged_df['cf_name'] == 'APPLE INC')]

    # Drop duplicates if any remain, keeping only one entry per fiscal year
    apple_df = apple_df.drop_duplicates(subset=['fy', 'form'])

    # Debug: Display the merged and deduplicated data for Apple
    print("Merged and deduplicated data for Apple Inc:\n", apple_df.head())

    return apple_df


def dcf_analysis(input_file, output_file, discount_rate, terminal_growth_rate, projection_years=20):
    # Process input data to merge financial statements
    apple_df = process_input_data(input_file)

    # Define possible labels for each required metric
    label_variations = {
        'cfo': ['NetCashProvidedByUsedInOperatingActivities'],
        'capex': ['PaymentsToAcquirePropertyPlantAndEquipment']
    }

    # Helper function to get financial values based on possible labels
    def get_financial_value(df, possible_labels):
        max_abs_value = pd.Series([0] * len(df))  # Default if none of the labels exist
        found_label = None
        
        for label in possible_labels:
            if label in df.columns:
                current_abs_value = df[label].abs()
                # Check if current label's values have a higher absolute value
                if current_abs_value.sum() > max_abs_value.sum():
                    max_abs_value = df[label].fillna(0)
                    found_label = label

        print(f"Using label: {found_label} for calculation.")
        return max_abs_value

    # Extract values using label variations
    cfo = get_financial_value(apple_df, label_variations['cfo'])
    capex = get_financial_value(apple_df, label_variations['capex'])

    # Debug: Print extracted values to verify
    print("\nExtracted CFO:\n", cfo)
    print("Extracted CapEx:\n", capex)

    # Calculate Free Cash Flow to the Firm (FCFF) for each year
    apple_df['FCFF'] = cfo - capex

    # Debug: Print calculated FCFF for each year
    print("\nCalculated FCFF for each year:\n", apple_df[['fy', 'FCFF']])

    # Select all historical years of FCFF
    historical_years = apple_df['fy'].tolist()
    historical_fcffs = apple_df['FCFF'].tolist()

    print("\nHistorical Years:\n", historical_years)
    print("Historical FCFFs:\n", historical_fcffs)

    # Initialize lists to store final output
    terminal_values = []
    net_present_values = []

    # Loop over each historical year to perform DCF analysis
    for i in range(len(historical_fcffs)):
        current_fcff = historical_fcffs[i]
        print(f"\nYear: {historical_years[i]}, Current FCFF: {current_fcff}")

        # Project future FCFF for the specified number of years using the growth rate
        projected_fcffs = [current_fcff * (1 + terminal_growth_rate) ** j for j in range(1, projection_years + 1)]
        print(f"Projected FCFFs for Year {historical_years[i]}:\n", projected_fcffs)

        # Calculate the present value of projected FCFFs for the given historical year
        discount_factors = [(1 + discount_rate) ** j for j in range(1, projection_years + 1)]
        discounted_fcffs = [fcff / discount_factors[j] for j, fcff in enumerate(projected_fcffs)]
        print(f"Discounted FCFFs for Year {historical_years[i]}:\n", discounted_fcffs)

        # Calculate the terminal value and its present value for the current year
        terminal_value = projected_fcffs[-1] * (1 + terminal_growth_rate) / (discount_rate - terminal_growth_rate)
        terminal_value_pv = terminal_value / discount_factors[-1]
        print(f"Terminal Value (Present Value) for Year {historical_years[i]}: {terminal_value_pv}")

        # Sum of discounted FCFFs and terminal value gives the net present value for the current year
        net_present_value = sum(discounted_fcffs) + terminal_value_pv
        print(f"Net Present Value for Year {historical_years[i]}: {net_present_value}")

        # Store results for the current year
        terminal_values.append(terminal_value_pv)
        net_present_values.append(net_present_value)

    # Create a final output DataFrame, converting values to millions of USD
    output_df = pd.DataFrame({
        'Year': historical_years,
        'Free Cash Flow to Firm (in millions)': [fcff / 1_000_000 for fcff in historical_fcffs],
        'Terminal Value (in millions)': [tv / 1_000_000 for tv in terminal_values],
        'Net Present Value (in millions)': [npv / 1_000_000 for npv in net_present_values]
    })

    # Save results to a new Excel file
    with pd.ExcelWriter(output_file) as writer:
        output_df.to_excel(writer, sheet_name='DCF Analysis', index=False)

    print(f"\nDCF analysis complete. Results saved to {output_file}")


if __name__ == '__main__':
    input_file = '10k_financials.xlsx'
    output_file = 'apple_dcf_analysis.xlsx'
    discount_rate = 0.1  # Example discount rate (10%)
    terminal_growth_rate = 0.02  # Example terminal growth rate (2%)
    projection_years = 20  # Number of years to project cash flows

    dcf_analysis(input_file, output_file, discount_rate, terminal_growth_rate, projection_years)
