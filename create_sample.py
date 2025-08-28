import pandas as pd
import os

# Create a sample test records file for demonstration
sample_data = {
    'Serial Number': ['SN001', 'SN002', 'SN003', 'SN004', 'SN005'],
    'Part Number': ['PN-001', 'PN-002', 'PN-003', 'PN-004', 'PN-005'],
    'Test Result': ['PASS', 'FAIL', 'PASS', 'MAYBE', 'PASS'],
    'Rack Bin': ['A1P', 'B2F', 'C3P', 'Testwip', 'D4P'],
    'Testing Date': ['2025-08-01', '2025-08-02', '2025-08-03', '2025-08-04', '2025-08-05']
}

df = pd.DataFrame(sample_data)

# Ensure input directory exists
os.makedirs('input', exist_ok=True)

# Save to Excel
df.to_excel('input/Sample_Test_Records.xlsx', index=False)
print("Sample test records file created successfully!")
