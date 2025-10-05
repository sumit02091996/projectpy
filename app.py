

import streamlit as st
import pandas as pd
import openai
from openai import OpenAI

client = OpenAI()

response = client.chat.completions.create(
    model="gpt-4o-mini",
    messages=[
        {"role": "system", "content": "You are a helpful assistant."},
        {"role": "user", "content": "Hello, how are you?"}
    ]
)

print(response.choices[0].message.content)

openai.api_key = "YOUR_OPENAI_API_KEY"

st.title("Excel AI Query Assistant")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        excel_data = pd.read_excel(uploaded_file, sheet_name=None)
        sheet_names = list(excel_data.keys())
        st.success(f"Loaded Excel file with sheets: {sheet_names}")

        sheet_choice = st.selectbox("Select sheet", sheet_names)
        df = excel_data[sheet_choice]
        st.dataframe(df)

        user_query = st.text_area("Enter your query or formula")

        if st.button("Run Query") and user_query.strip() != "":
            prompt = f"""
Given the pandas DataFrame df with data:

{df.head().to_csv(index=False)}

Perform this user query on df:

{user_query}

Write python pandas code to do this and assign the result to a variable named result.
Only output the code.
"""
            try:
                response = openai.ChatCompletion.create(
                    model="gpt-4",
                    messages=[
                        {"role": "system", "content": "You are a helpful assistant that writes python pandas code."},
                        {"role": "user", "content": prompt}
                    ],
                    max_tokens=400,
                    temperature=0
                )
                code = response['choices'][0]['message']['content']
                st.code(code, language='python')

                local_vars = {"df": df.copy(), "pd": pd}
                exec(code, {}, local_vars)

                if "result" in local_vars:
                    st.subheader("Query Result:")
                    st.dataframe(local_vars["result"])
                else:
                    st.info("No 'result' variable found in code output.")

            except Exception as e:
                st.error(f"Error executing code: {e}")
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
else:
    st.info("Upload an Excel file to start")

	

