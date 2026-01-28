# -*- coding: utf-8 -*-
import json
import pandas as pd
from google import genai
from google.genai import types

import json
import os
import sys
import io 
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
def create_boq_json():
    df = pd.read_excel(r"C:\Users\PS\AppData\Roaming\pyRevit\Extensions\mytools.extension\mytools.tab\quantities.panel\auto-boq.pushbutton\boq.xlsx")

    boq_items = []
    for _, row in df.iterrows():
        boq_items.append({
            "code": str(row["Code"]),
            "title": row["Title"],
            "unit": row["Unit"],
            "category": row["Category"],
            "keywords": [k.strip() for k in row["Keywords"].split("،")]
        })
    return boq_items

def create_structural_elements_json():
    with open(r"C:\Users\PS\AppData\Roaming\pyRevit\Extensions\mytools.extension\mytools.tab\quantities.panel\auto-boq.pushbutton\temp.json") as f:
        return json.load(f)

def call_gemini_to_excel():
    # 1. Initialize Client
    client = genai.Client(api_key='your api')

    # 2. The Prompt (Hardcoded as requested)
    boq_mapping_prompt = f"""
You are a construction quantity surveyor. I have two JSON inputs:
1) BOQ items (item_code, description, unit)
2) Structural model elements (name, category, material, volume_m3)

Task: Match each BOQ item to the relevant structural elements.
Matching rules:
- Match based on construction meaning, not exact text
- Use element category, material, and description
- Do NOT invent elements
- If no match exists, return an empty list

Output rules:
- Return ONLY valid JSON
- No explanations, no comments

Output format:
[{{"item_code": "string", "title": "", "unit": "", "quantity": number, "matched_elements": [{{"element_name": "string", "volume_m3": number}}] }}]

DATA:
{json.dumps(create_structural_elements_json(), ensure_ascii=False)}

{json.dumps(create_boq_json(), ensure_ascii=False)}
    """
    print("Sending prompt to Gemini...",boq_mapping_prompt)
    # 3. Get JSON Response
    response = client.models.generate_content(
        model='gemini-2.5-flash',
        contents=boq_mapping_prompt,
        config=types.GenerateContentConfig(
            response_mime_type='application/json',
            thinking_config=types.ThinkingConfig(thinking_budget=1024)
        )
    )
    print("Response", response.text);
    # 4. Parse JSON and Flatten for Excel
    try:
        data = json.loads(response.text)
        
        # Flattening the nested list so it looks good in Excel
        flattened_data = []
        for item in data:
            flattened_data.append({
                "BOQ Code": item.get('item_code'),
                "Title": item.get('title'),
                "Unit": item.get('unit'),
                "Total Item Qty": item.get('quantity'),
            })

        # 5. Create DataFrame and Save to Excel
        df = pd.DataFrame(flattened_data)
        file_name = r"C:\Users\PS\AppData\Roaming\pyRevit\Extensions\mytools.extension\mytools.tab\quantities.panel\auto-boq.pushbutton\BOQ_Mapping_Results.xlsx"
        df.to_excel(file_name, index=False)
        
        print("Success! Data exported to {file_name}")
        print(df.head()) # Preview in console

    except Exception as e:
        print("Error parsing or saving: {e}")

if __name__ == "__main__":
    call_gemini_to_excel()