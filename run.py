import os
import pandas as pd
import re
import streamlit as st
from datetime import datetime
from spellchecker import SpellChecker
import language_tool_python
from io import BytesIO

# 获取当前时间，格式为：YYYY-MM-DD_HH-MM-SS
def get_current_time():
    return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

# 去除多余的空格
def remove_extra_spaces(text):
    return re.sub(r'\s+', ' ', text).strip()

# 根据标点前后的字符调整标点符号
def adjust_punctuation(text):
    punctuation_map = {
        '，': ',',  '。': '.',  '；': ';',  '：': ':',  '！': '!',  '？': '?', 
        '（': '(',  '）': ')',  '【': '[',  '】': ']',  '《': '<',  '》': '>', 
        '“': '"',  '”': '"',  '‘': "'",  '’': "'",  '……': '...',  '·': '.'
    }
    
    def replace_punctuation(match):
        punctuation = match.group(0)
        prev_char = match.string[match.start() - 1] if match.start() > 0 else ''
        if prev_char and (ord(prev_char) > 128):
            return punctuation
        return punctuation_map.get(punctuation, punctuation)

    pattern = r'[，。；：！？（）【】《》“”‘’……·]'
    return re.sub(pattern, replace_punctuation, text)

# 标点符号后添加空格（不包括引号）
def add_space_after_punctuation(text):
    def add_space(match):
        punctuation = match.group(0)
        next_char = match.string[match.end()] if match.end() < len(match.string) else ''
        if punctuation not in ['“', '”', '‘', '’'] and next_char != ' ':
            return punctuation + ' '
        return punctuation

    pattern = r'[，。；：！？（）【】《》“”‘’……·]'
    return re.sub(pattern, add_space, text)

# 拼写检查和修正
def correct_spelling(text):
    spell = SpellChecker()
    words = text.split()
    corrected_words = [spell.correction(word) if word not in spell else word for word in words]
    return ' '.join(corrected_words)

# 语法检查和修正
def correct_grammar(text):
    tool = language_tool_python.LanguageTool('en-US')
    matches = tool.check(text)
    return language_tool_python.utils.correct(text, matches)

# 处理DataFrame中的标点符号、空格、去除前后空白并调整标点符号
def process_dataframe(df):
    for col in df.columns:
        df[col] = df[col].apply(lambda x: remove_extra_spaces(str(x)) if isinstance(x, str) else x)
        df[col] = df[col].apply(lambda x: adjust_punctuation(str(x)) if isinstance(x, str) else x)
        df[col] = df[col].apply(lambda x: add_space_after_punctuation(str(x)) if isinstance(x, str) else x)
        df[col] = df[col].apply(lambda x: correct_spelling(str(x)) if isinstance(x, str) else x)
        df[col] = df[col].apply(lambda x: correct_grammar(str(x)) if isinstance(x, str) else x)
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
    return df

# Convert DataFrame to Excel and return as a BytesIO object
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# Streamlit App
def main():
    st.title("Excel File Processing and Correction")
    
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        
        st.write("Original Data:")
        st.write(df.head())  # Display the first few rows of the uploaded file
        
        # Process the DataFrame
        processed_df = process_dataframe(df)
        
        st.write("Processed Data:")
        st.write(processed_df.head())  # Display the processed data
        
        # Convert the processed dataframe to an Excel file for download
        processed_excel = convert_df_to_excel(processed_df)
        st.download_button(label="Download Processed Excel", data=processed_excel, file_name="processed_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
