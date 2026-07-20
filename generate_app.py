#!/usr/bin/env python3
"""Generate streamlit_app.py to avoid encoding corruption."""
import os, json

TARGET = r"C:\Users\adm\.openclaw-autoclaw\agents\task-2-rag\workspace\rag-study\DELIVERY\streamlit_app.py"

sample = {
  "track1_baseline": {"faithfulness": 0.617, "precision": 0.987, "recall": 0.803, "relevance": 0.733, "hit_rate@5": 0.667, "latency_ms": 1667},
  "track1_final": {"faithfulness": 0.633, "precision": 1.000, "recall": 0.817, "relevance": 0.720, "hit_rate@5": 0.933, "latency_ms": 1857},
  "track2_final": {"faithfulness": 0.600, "precision": 0.960, "recall": 0.810, "relevance": 0.580, "hit_rate@5": 0.433, "latency_ms": 2660},
  "track3_final": {"faithfulness": 0.727, "precision": 0.933, "recall": 0.803, "relevance": 0.467, "hit_rate@5": 0.667, "latency_ms": 1145},
}

code = r'''# RAG Engineering Study — Streamlit Dashboard
# Tracks: Retrieval Engineering | Graph-Based Retrieval | Context & Memory Optimization
# Dataset: LegalBench-RAG (CUAD, MAUD, ContractNLI, PrivacyQA)
# Run: streamlit run streamlit_app.py

import json, os
import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(page_title="RAG Engineering Study Dashboard", page_icon="\U0001f4ca", layout="wide")

# ---- DATA LOADING ----

def _load_json(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        st.warning(f"Could not load {os.path.basename(path)}: {e}")
        return None

def load_results(data_dir="results/"):
    patterns = {
        "track1_baseline": "track1_baseline.json",
        "track1_chunking": "track1_chunking.json",
        "track1_retrieval": "track1_retrieval.json",
        "track1_reranking": "track1_reranking.json",
        "track1_transform": "track1_transform.json",
        "track1_final": "track1_final.json",
        "track2_chunking": "track2_chunking.json",
        "track2_graph_search": "track2_graph_search.json",
        "track2_communities": "track2_communities.json",
        "track2_routed": "track2_routed.json",
        "track2_final": "track2_final.json",
        "track3_prompt_versioning": "track3_prompt_versioning.json",
        "track3_compression": "track3_compression.json",
        "track3_memory": "track3_memory.json",
        "track3_caching": "track3_caching.json",
        "track3_fullcontext": "track3_fullcontext.json",
        "track3_final": "track3_final.json",
    }
    return {k: _load_json(os.path.join(data_dir, fn)) for k, fn in patterns.items()}

SAMPLE = json.loads("""''' + json.dumps(sample, ensure_ascii=False) + r'''""")

def get_data(data_dir="results/"):
    results = load_results(data_dir)
    if any(v is not None for v in results.values()):
        return results, False
    st.info("No result files found — showing demo data.")
    return SAMPLE, True
'''

# Instead of all that complexity, just write the whole thing via Python
with open(TARGET, 'w', encoding='utf-8') as f:
    f.write(code)
print(f"Wrote {len(code)} bytes to {TARGET}")
