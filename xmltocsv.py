import pandas as pd


# This looks for any node named 'Order' regardless of namespace
# Change 'Order' to 'Row' if the file uses Rows
df = pd.read_xml(
    "amazon_report_20251218_012809.xml", 
    xpath=".//*[local-name()='Order']", 
    compression="gzip" # keep this if the file is still zipped
)

# Save to CSV
df.to_csv("output_xml_1.csv", index=False)