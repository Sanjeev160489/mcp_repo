# Tests for ExcelReader class

import unittest
import pandas as pd
import os
from excelMCP import ExcelReader

class TestExcelReader(unittest.TestCase):
    
    def setUp(self):
        # Create a sample DataFrame and save to Excel for testing
        self.test_file = 'test_data.xlsx'
        self.test_data = pd.DataFrame({
            'Column1': [1, 2, 3, 4, 5],
            'Column2': ['A', 'B', 'C', 'D', 'E']
        })
        self.test_data.to_excel(self.test_file, sheet_name='TestSheet', index=False)
        
    def tearDown(self):
        # Clean up test file
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
    
    def test_initialization(self):
        reader = ExcelReader()
        self.assertIsNone(reader.file_path)
        self.assertIsNone(reader.workbook)
        self.assertIsNone(reader.data)
        
        reader2 = ExcelReader(self.test_file)
        self.assertEqual(reader2.file_path, self.test_file)
    
    def test_load_file(self):
        reader = ExcelReader()
        self.assertTrue(reader.load_file(self.test_file))
        self.assertEqual(reader.file_path, self.test_file)
        self.assertIsNotNone(reader.workbook)
    
    def test_get_sheet_names(self):
        reader = ExcelReader(self.test_file)
        reader.load_file()
        sheet_names = reader.get_sheet_names()
        self.assertIn('TestSheet', sheet_names)
    
    def test_read_sheet_to_dataframe(self):
        reader = ExcelReader(self.test_file)
        df = reader.read_sheet_to_dataframe(sheet_name='TestSheet')
        self.assertIsNotNone(df)
        self.assertEqual(len(df), 5)  # 5 rows
        self.assertEqual(len(df.columns), 2)  # 2 columns
        
        # Check column names
        self.assertIn('Column1', df.columns)
        self.assertIn('Column2', df.columns)
        
        # Check data values
        self.assertEqual(df['Column1'].iloc[0], 1)
        self.assertEqual(df['Column2'].iloc[4], 'E')

if __name__ == '__main__':
    unittest.main()
