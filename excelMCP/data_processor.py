# Data processor module

import pandas as pd
import numpy as np

class DataProcessor:
    """Class for processing and transforming Excel data."""
    
    def __init__(self, data=None):
        """Initialize with optional DataFrame."""
        self.data = data
    
    def set_data(self, data):
        """Set the DataFrame to process."""
        if not isinstance(data, pd.DataFrame):
            raise TypeError("Data must be a pandas DataFrame")
        self.data = data
    
    def clean_data(self, columns=None):
        """Clean data by removing NaN values and whitespace from strings."""
        if self.data is None:
            raise ValueError("No data set. Call set_data() first.")
        
        clean_df = self.data.copy()
        
        # Define which columns to process
        cols_to_clean = columns if columns else clean_df.columns
        
        # Clean string columns
        for col in cols_to_clean:
            if clean_df[col].dtype == 'object':
                # Remove leading/trailing whitespace
                clean_df[col] = clean_df[col].str.strip()
                
        return clean_df
    
    def filter_by_value(self, column, value, comparison='equals'):
        """Filter DataFrame by a condition on a column value."""
        if self.data is None:
            raise ValueError("No data set. Call set_data() first.")
        
        if column not in self.data.columns:
            raise ValueError(f"Column '{column}' not found in data")
        
        if comparison == 'equals':
            return self.data[self.data[column] == value]
        elif comparison == 'contains':
            return self.data[self.data[column].str.contains(value, na=False)]
        elif comparison == 'greater_than':
            return self.data[self.data[column] > value]
        elif comparison == 'less_than':
            return self.data[self.data[column] < value]
        else:
            raise ValueError(f"Unknown comparison type: {comparison}")
    
    def group_and_aggregate(self, group_by, agg_dict):
        """Group data and apply aggregation functions.
        
        Args:
            group_by (str or list): Column(s) to group by
            agg_dict (dict): Dictionary mapping columns to aggregation functions
                Example: {'sales': 'sum', 'quantity': ['min', 'max', 'mean']}
        """
        if self.data is None:
            raise ValueError("No data set. Call set_data() first.")
        
        return self.data.groupby(group_by).agg(agg_dict).reset_index()
