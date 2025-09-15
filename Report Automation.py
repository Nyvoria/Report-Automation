import pandas as pd
from openai import OpenAI

#read the Excel file
df = pd.read_excel(r"C:\Users\vishw\OneDrive\Desktop\Infor Research\10 Sept\26- Market Research Reports - Food Beverages.xlsx")

#print column names to confirmm
print("Cleaned columns:", df.columns.tolist())

#extract the 'report title' column
titles = df["Report Title"].dropna().tolist()

#show first few titles
print("Report titles found:")
for t in titles[:5]:
    print("-",t)

#Initialize OpenAI client (replace with your API key)
client = OpenAI(api_key="your_api_key_here")

# Pick one samle row (first where Summary is not empty)
sample_row = df.dropna(subset=["Summar"]).iloc[0]

sample_title = sample_row["Report Title"]
sample_toc = sample_row["Table of Contents"]
sample_tables = sample_row["List of Tables:"]
sample_figures = sample_row["List of Fugures"]
sample_companies = sample_row["Companies Mentioned"]
sample_summary = sample_row["Summary"]

# Pick one empty row (first where summary is blank)
target_row = df[df["Summary"].isna()].iloc[0]
target_title = target_row["Report Title"]

# Build prompt
prompt = f"""
Here is a sample report:

Title: {sample_title}
Table of Contents: {sample_toc}
List of Tables: {sample_tables}
List of Figures: {sample_figures}
Companies Mentioned: {sample_companies}
Summary: {sample_summary}

Now, write a report in the same structure for this new title:
{target_title}

Please return the output in this exact format:

Table of Contents: ...
List of Tables: ...
List of Figures: ...
Companies Mentioned: ...
Summary: ...
"""

# Send prompt to OpenAI
response = client.chat.completion.create(
    model="gpt-40-mini",
    message=[{"role":"user","content": prompt}],
)

# Print response
print("AI Response:\n")
print(response.choices[0].message.content)
