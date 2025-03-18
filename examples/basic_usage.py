# Example of basic usage of the excelMCP library

import pandas as pd
from excelMCP import ExcelReader, ExcelWriter, DataProcessor

def create_sample_data():
    """Create a sample DataFrame for demonstration purposes."""
    data = {
        'Product': ['Widget A', 'Widget B', 'Widget C', 'Widget A', 'Widget B'],
        'Region': ['North', 'South', 'North', 'East', 'North'],
        'Sales': [1000, 1500, 800, 1200, 950],
        'Quantity': [100, 120, 80, 110, 95]
    }
    return pd.DataFrame(data)

def main():
    # Create sample data and save to Excel
    df = create_sample_data()
    
    # Create a new Excel file with the sample data
    writer = ExcelWriter('sample_data.xlsx')
    writer.dataframe_to_excel(df, sheet_name='SalesData')
    writer.save()
    
    print("Sample Excel file created: sample_data.xlsx")
    
    # Read the Excel file
    reader = ExcelReader('sample_data.xlsx')
    sheet_names = reader.get_sheet_names()
    print(f"Sheets in workbook: {sheet_names}")
    
    # Read data into DataFrame
    sales_data = reader.read_sheet_to_dataframe('SalesData')
    print("\nRead data from Excel:")
    print(sales_data)
    
    # Process the data
    processor = DataProcessor(sales_data)
    
    # Filter data for Widget A
    widget_a_data = processor.filter_by_value('Product', 'Widget A')
    print("\nFiltered data for Widget A:")
    print(widget_a_data)
    
    # Group by Product and calculate aggregates
    agg_data = processor.group_and_aggregate(
        'Product', 
        {'Sales': ['sum', 'mean'], 'Quantity': ['sum', 'mean']}
    )
    print("\nAggregated data by Product:")
    print(agg_data)
    
    # Write the aggregated data to a new sheet
    writer = ExcelWriter('processed_data.xlsx')
    writer.dataframe_to_excel(agg_data, sheet_name='AggregateByProduct')
    writer.save()
    
    print("\nProcessed data saved to: processed_data.xlsx")

if __name__ == "__main__":
    main()
