import os
import pandas as pd
from secfsdstools.d_container.databagmodel import JoinedDataBag
from secfsdstools.e_collector.zipcollecting import ZipCollector
from secfsdstools.u_usecases.bulk_loading import default_postloadfilter
from secfsdstools.f_standardize.bs_standardize import BalanceSheetStandardizer
from secfsdstools.f_standardize.is_standardize import IncomeStatementStandardizer
from secfsdstools.f_standardize.cf_standardize import CashFlowStandardizer

def get_combined_financials(ticker, years, output_file):
    # Collect the ZIP files based on the specified years
    zip_names = [f"{year}q{q}.zip" for year in years for q in range(1, 5)]
    collector = ZipCollector.get_zip_by_names(names=zip_names,
                                              forms_filter=["10-K"],
                                              post_load_filter=default_postloadfilter)
    
    # Collect and join the data
    joined_bag: JoinedDataBag = collector.collect().join()
    print(f"Number of loaded reports for {ticker}: ", len(joined_bag.sub_df))

    # Initialize standardizers for balance sheet, income statement, and cash flow statement
    bs_standardizer = BalanceSheetStandardizer()
    is_standardizer = IncomeStatementStandardizer()
    cf_standardizer = CashFlowStandardizer()

    # Standardize each financial statement
    standardized_bs_df = joined_bag.present(bs_standardizer)
    standardized_is_df = joined_bag.present(is_standardizer)
    standardized_cf_df = joined_bag.present(cf_standardizer)

    # Merge the standardized DataFrames on 'cik', 'fy' (fiscal year), and 'form' columns
    merged_df = pd.merge(standardized_bs_df, standardized_is_df, on=['cik', 'fy', 'form'], suffixes=('_bs', '_is'))
    merged_df = pd.merge(merged_df, standardized_cf_df, on=['cik', 'fy', 'form'], suffixes=('', '_cf'))

    # Drop redundant or unnecessary columns for simplicity
    merged_df = merged_df.drop(columns=['coreg', 'uom', 'coreg_cf', 'uom_cf'], errors='ignore')

    # Write the merged DataFrame to an Excel file
    merged_df.to_excel(output_file, index=False)
    print(f"Financial data has been saved to {output_file}")

if __name__ == '__main__':
    # Example usage
    ticker = 'AAPL'
    years = [2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023]
    output_file = '10k_financials.xlsx'
    get_combined_financials(ticker, years, output_file)
