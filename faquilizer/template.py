NOTEBOOK_TEMPLATE = """
# Paste the entire contents of this output into a new Jupyter Notebook.
# This template provides a complete, runnable workflow.

# %%
# Cell 1: Instructions & Setup (Markdown)
## Faquillizer: Your AI-Powered List Processor
### This notebook takes a list of URLs, fetches their titles, and uses AI to generate insights.

# %%
# Cell 2: Imports & Job Initialization (Code)
# pip install pipulate google-generativeai requests beautifulsoup4 pandas openpyxl
from pipulate import pip
import google.generativeai as genai
import requests
from bs4 import BeautifulSoup
import pandas as pd
import getpass
import json

# Each run of this notebook is a "job" with a unique name
job = "faq-session-01"

# %%
# Cell 3: Google AI Authentication (Code)
# This cell handles your Google AI API key.
# It will ask for your key once, then store it for this job.
API_KEY_STEP = "api_key"
api_key = pip.get(job, API_KEY_STEP)

if not api_key:
    try:
        # Use getpass for secure input in a notebook
        api_key = getpass.getpass("Enter your Google AI API Key (get one at https://aistudio.google.com/app/apikey): ")
        pip.set(job, API_KEY_STEP, api_key)
        print("✅ API Key received and stored for this session.")
    except Exception as e:
        print(f"Could not get API key: {e}")

if api_key:
    genai.configure(api_key=api_key)
    print("✅ Google AI configured successfully.")

# %%
# Cell 4: List Input (Code)
## Paste your list of URLs between the triple quotes below.
URL_LIST_STEP = "url_list"
EASILY_PASTED_LIST = '''
https://www.google.com
https://www.github.com
https://www.mikelev.in
'''.split("\n")[1:-1]

pip.set(job, URL_LIST_STEP, EASILY_PASTED_LIST)
urls_to_process = pip.get(job, URL_LIST_STEP, [])
print(f"✅ Found {len(urls_to_process)} URLs to process.")

# %%
# Cell 5: Processing Loop (Code)
## This cell fetches the title for each URL.
### If you restart the kernel and run it again, it will only process the remaining URLs.
RAW_DATA_STEP = "raw_data"
processed_data = pip.get(job, RAW_DATA_STEP, [])
processed_urls = {item['url'] for item in processed_data}

print(f"🔄 Starting processing... {len(processed_urls)} URLs already complete.")

for url in urls_to_process:
    if url in processed_urls:
        continue # Skip already processed URLs
    try:
        print(f"Fetching {url}...")
        response = requests.get(url, timeout=10)
        soup = BeautifulSoup(response.content, 'html.parser')
        title = soup.title.string if soup.title else "No Title Found"
        processed_data.append({'url': url, 'title': title.strip()})
        pip.set(job, RAW_DATA_STEP, processed_data) # Save progress after each item!
        processed_urls.add(url)
    except Exception as e:
        print(f"❌ Failed to process {url}: {e}")

print("✅ Raw data processing complete.")

# %%
# Cell 6: AI Augmentation (Optional but Powerful) (Code)
AI_INSIGHTS_STEP = "ai_insights"
ai_insights = pip.get(job, AI_INSIGHTS_STEP, [])
processed_titles = {item['title'] for item in ai_insights}

print("🧠 Generating AI insights...")
model = genai.GenerativeModel('gemini-2.5-flash')

for item in processed_data:
    if item['title'] in processed_titles:
        continue
    try:
        prompt = f"Based on the title '{item['title']}', what is the likely primary topic of this page? Be concise."
        response = model.generate_content(prompt)
        ai_insights.append({'title': item['title'], 'topic': response.text.strip()})
        pip.set(job, AI_INSIGHTS_STEP, ai_insights)
    except Exception as e:
        print(f"❌ AI insight failed for '{item['title']}': {e}")

print("✅ AI insights generated.")

# %%
# Cell 7: DataFrame Display (Code)
## Merge raw data with AI insights and display as a styled table.
df_raw = pd.DataFrame(processed_data)
df_ai = pd.DataFrame(ai_insights)

df_final = pd.merge(df_raw, df_ai, on="title", how="left")

# --- Styling Pandas DataFrames ---
styled_df = df_final.style.set_properties(**{
    'text-align': 'left',
    'white-space': 'pre-wrap',
}).set_table_styles([
    {'selector': 'th', 'props': [('text-align', 'left'), ('font-weight', 'bold')]},
    {'selector': 'tr:nth-of-type(odd)', 'props': [('background-color', '#f9f9f9')]},
    {'selector': 'tr:nth-of-type(even)', 'props': [('background-color', 'white')]}
]).hide(axis="index")

display(styled_df)
pip.set(job, "final_dataframe", df_final.to_json())

# %%
# Cell 8: Export to Excel (Code)
## Export the final DataFrame to a formatted Excel file.
EXPORT_FILE_STEP = "export_file_path"
output_filename = f"{job}_output.xlsx"

try:
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name='FAQ_Data')
        # Auto-fit column widths
        worksheet = writer.sheets['FAQ_Data']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    pip.set(job, EXPORT_FILE_STEP, output_filename)
    print(f"✅ Success! Data exported to '{output_filename}'")
except Exception as e:
    print(f"❌ Failed to export to Excel: {e}")
"""
