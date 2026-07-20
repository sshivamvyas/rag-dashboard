import os, json

TARGET = r'C:\Users\adm\.openclaw-autoclaw\agents\task-2-rag\workspace\rag-study\DELIVERY\streamlit_app.py'

# Build the complete source code as a list of lines
code_lines = []

def L(s=''):
    code_lines.append(s)

L('# RAG Engineering Study - Streamlit Dashboard')
L('# Tracks: Retrieval Engineering | Graph-Based Retrieval | Context & Memory Optimization')
L('# Dataset: LegalBench-RAG (CUAD, MAUD, ContractNLI, PrivacyQA)')
L('# Run: streamlit run streamlit_app.py')
L('')
L('import json, os')
L('import pandas as pd')
L('import numpy as np')
L('import streamlit as st')
L('')
L("st.set_page_config(page_title='RAG Engineering Study Dashboard', page_icon='\\U0001f4ca', layout='wide')")
L('')

print(f'Generated {len(code_lines)} lines')
with open(TARGET, 'w', encoding='utf-8') as f:
    f.write('\\n'.join(code_lines))
print(f'Wrote to {TARGET}')
