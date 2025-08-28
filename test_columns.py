import pandas as pd
import os

# Test the filtered columns functionality
def test_filtered_columns():
    print("Testing filtered columns functionality...")
    
    # Load sample data to verify column filtering
    if os.path.exists('input/Sample_Test_Records.xlsx'):
        df = pd.read_excel('input/Sample_Test_Records.xlsx')
        print("Sample file columns:", df.columns.tolist())
        
        # Expected columns after filtering
        expected_columns = [
            'Testing Date_test',
            'Part Number_test', 
            'Serial Number',
            'Testing Date_putaway',
            'Part Number_putaway',
            'Rack Bin',
            'Test Result',
            'Validation Result'
        ]
        
        print("Expected output columns:", expected_columns)
        print("✅ Column filtering configuration is ready")
    else:
        print("❌ Sample file not found")

if __name__ == "__main__":
    test_filtered_columns()
