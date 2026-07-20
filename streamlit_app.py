# RAG Engineering Study — Streamlit Dashboard
# Tracks: Retrieval Engineering | Graph-Based Retrieval | Context & Memory Optimization
# Dataset: LegalBench-RAG (CUAD, MAUD, ContractNLI, PrivacyQA)
# Run: streamlit run streamlit_app.py

import streamlit as st
import pandas as pd
import numpy as np
import json, os

st.set_page_config(page_title="RAG Engineering Study", layout="wide")

METRICS = ["faithfulness", "precision", "recall", "relevance", "hit_rate@5"]
METRIC_LABELS = ["Faithfulness", "Precision", "Recall", "Relevance", "Hit Rate@5"]

SAMPLE = {
    "baseline": {"faithfulness": 0.617, "precision": 0.987, "recall": 0.803, "relevance": 0.733, "hit_rate@5": 0.667, "latency": 1667},
    "t1_best": {"faithfulness": 0.633, "precision": 1.000, "recall": 0.817, "relevance": 0.720, "hit_rate@5": 0.933, "latency": 1857},
    "t2_best": {"faithfulness": 0.600, "precision": 0.960, "recall": 0.810, "relevance": 0.580, "hit_rate@5": 0.433, "latency": 2660},
    "t3_best": {"faithfulness": 0.727, "precision": 0.933, "recall": 0.803, "relevance": 0.467, "hit_rate@5": 0.667, "latency": 1145},
    "t2_extra": {"nodes": 17, "edges": 1, "fallback_pct": 53, "entities": 180},
    "t3_cache": {"hit_rate": 33.3, "speedup": 615, "faith_delta": -0.015},
}

def load_results(data_dir):
    if not os.path.isdir(data_dir):
        return None
    files = [f for f in os.listdir(data_dir) if f.endswith('.json') and f.startswith('track')]
    if not files:
        return None
    data = {}
    for fname in files:
        try:
            with open(os.path.join(data_dir, fname), encoding='utf-8') as f:
                data[fname.replace('.json', '')] = json.load(f)
        except Exception:
            pass
    return data if data else None

def extract_avg(data, key, field='avg'):
    if data and key in data:
        d = data[key]
        if isinstance(d, dict):
            if field in d:
                return d[field]
            if any(m in d for m in METRICS):
                return d
    return None

# Sidebar
with st.sidebar:
    st.markdown("## RAG Engineering Study")
    st.markdown("LegalBench-RAG | 3 Tracks | Zero-Cost")
    st.markdown("---")
    page = st.radio("Navigation", [
        "Executive Summary", "Track 1: Retrieval", "Track 2: Graph",
        "Track 3: Context & Memory", "3-Track Comparison"
    ])
    with st.expander("Data Source"):
        data_dir = st.text_input("Results directory", value="results")
    st.markdown("---")
    st.caption("FLAN-T5-XL (3B) + all-MiniLM-L6-v2")
    st.caption("Kaggle T4 GPU | Cost: $0.00")

data = load_results(data_dir)
using_sample = data is None

if using_sample:
    baseline, t1_best, t2_best, t3_best = SAMPLE["baseline"], SAMPLE["t1_best"], SAMPLE["t2_best"], SAMPLE["t3_best"]
    t2_extra, t3_cache = SAMPLE["t2_extra"], SAMPLE["t3_cache"]
else:
    baseline = extract_avg(data, 'track1_baseline') or SAMPLE["baseline"]
    t1_raw = extract_avg(data, 'track1_final', 't1_avg') or extract_avg(data, 'track1_final') or SAMPLE["t1_best"]
    t1_best = t1_raw
    t2_raw = extract_avg(data, 'track2_final', 'rt_avg') or extract_avg(data, 'track2_routed', 'router') or SAMPLE["t2_best"]
    t2_best = t2_raw
    t3_raw = extract_avg(data, 'track3_final', 'overall') or SAMPLE["t3_best"]
    t3_best = t3_raw
    gs = extract_avg(data, 'track2_graph_search') or {}
    rt = extract_avg(data, 'track2_routed') or {}
    t2_extra = {"nodes": gs.get('graph_nodes', 17), "edges": gs.get('graph_edges', 1),
                "fallback_pct": rt.get('fallbacks', 15)/30*100, "entities": gs.get('entity_count', 180)}
    cs = extract_avg(data, 'track3_caching') or {}
    t3_cache = {"hit_rate": cs.get('hit_rate', 0.333)*100, "speedup": cs.get('speedup_x', 615),
                "faith_delta": cs.get('faith_delta', -0.015)}

latency_data = {"Config": ["Baseline", "Track 1", "Track 2", "Track 3"],
    "Latency (ms)": [baseline.get("latency",0), t1_best.get("latency",0),
                     t2_best.get("latency",0), t3_best.get("latency",0)]}

if page == "Executive Summary":
    st.title("Executive Summary")
    c1, c2, c3 = st.columns(3)
    with c1: st.metric("Best Faithfulness", f"{t3_best.get('faithfulness',0):.3f}", "Track 3")
    with c2: st.metric("Best Latency", f"{t3_best.get('latency',0):.0f} ms", "-522ms vs baseline")
    with c3: st.metric("Best Hit Rate@5", f"{t1_best.get('hit_rate@5',0):.3f}", "Track 1")

    if using_sample: st.info("Using sample data — place JSON result files in 'results/' directory.")

    st.subheader("Metric Comparison Across Tracks")
    df = pd.DataFrame({ml: [baseline.get(m,0), t1_best.get(m,0), t2_best.get(m,0), t3_best.get(m,0)]
        for ml, m in zip(METRIC_LABELS, METRICS)}, index=["Baseline", "T1 Retr", "T2 Graph", "T3 Ctx"])
    st.dataframe(df.style.highlight_max(color="#d4edda", axis=0).highlight_min(color="#f8d7da", axis=0),
                 use_container_width=True)
    st.bar_chart(df.T)
    st.bar_chart(pd.DataFrame(latency_data).set_index("Config"))

    st.subheader("Key Insights")
    cols = st.columns(2)
    with cols[0]:
        st.success("**Track 1** wins hit_rate: hybrid retrieval +40% h@5")
        st.success("**Track 3** wins faithfulness: structured prompt + generative compression")
    with cols[1]:
        st.warning("**Track 2**: entity extraction at 3B scale insufficient (17 nodes)")
        st.info("**All tracks**: $0 cost, FLAN-T5-XL + MiniLM on Kaggle T4")

elif page == "Track 1: Retrieval":
    st.title("Track 1: Retrieval Engineering")
    t1w = sum(1 for m in METRICS if t1_best.get(m,0) > baseline.get(m,0)+0.01)
    t1l = sum(1 for m in METRICS if t1_best.get(m,0) < baseline.get(m,0)-0.01)
    st.subheader(f"Final: Track 1 wins {t1w}/5 | Baseline wins {t1l}/5 | Ties {5-t1w-t1l}/5")
    rows = []
    for i,m in enumerate(METRICS):
        b, t = baseline.get(m,0), t1_best.get(m,0)
        rows.append({"Metric": METRIC_LABELS[i], "Baseline": f"{b:.3f}", "Track 1": f"{t:.3f}", "Delta": f"{t-b:+.3f}"})
    rows.append({"Metric": "Latency (ms)", "Baseline": f"{baseline.get('latency',0):.0f}",
                 "Track 1": f"{t1_best.get('latency',0):.0f}",
                 "Delta": f"{t1_best.get('latency',0)-baseline.get('latency',0):+.0f}"})
    st.dataframe(pd.DataFrame(rows).set_index("Metric"), use_container_width=True)
    st.info("Hybrid retrieval (dense + keyword RRF) delivers +40% hit_rate@5. Reranking adds +86.5% latency.")

    with st.expander("S3: Retrieval Mode"):
        ret = {"dense": [0.617, 0.667, 1684], "keyword": [0.603, 0.933, 1293], "hybrid": [0.613, 0.967, 982]}
        rd = pd.DataFrame(ret, index=["Faithfulness", "Hit Rate@5", "Latency (ms)"]).T
        st.dataframe(rd.style.highlight_max(color="#d4edda", subset=["Faithfulness", "Hit Rate@5"])
                     .highlight_min(color="#d4edda", subset=["Latency (ms)"]), use_container_width=True)
        st.bar_chart(rd[["Faithfulness", "Hit Rate@5"]])

    with st.expander("S4: Reranking"):
        st.dataframe(pd.DataFrame({"No Rerank": [0.613, 0.942, 983], "Rerank (10>3)": [0.633, 1.000, 1834]},
                     index=["Faithfulness", "Precision", "Latency (ms)"]), use_container_width=True)

elif page == "Track 2: Graph":
    st.title("Track 2: Graph-Based Retrieval")
    t2w = sum(1 for m in METRICS if t2_best.get(m,0) > baseline.get(m,0)+0.01)
    t2l = sum(1 for m in METRICS if t2_best.get(m,0) < baseline.get(m,0)-0.01)
    st.subheader(f"Final: Track 2 wins {t2w}/5 | Baseline wins {t2l}/5 | Ties {5-t2w-t2l}/5")
    rows = []
    for i,m in enumerate(METRICS):
        b, t = baseline.get(m,0), t2_best.get(m,0)
        rows.append({"Metric": METRIC_LABELS[i], "Baseline": f"{b:.3f}", "Track 2": f"{t:.3f}", "Delta": f"{t-b:+.3f}"})
    rows.append({"Metric": "Latency (ms)", "Baseline": f"{baseline.get('latency',0):.0f}",
                 "Track 2": f"{t2_best.get('latency',0):.0f}",
                 "Delta": f"{t2_best.get('latency',0)-baseline.get('latency',0):+.0f}"})
    st.dataframe(pd.DataFrame(rows).set_index("Metric"), use_container_width=True)
    st.warning("Graph too sparse (17 nodes, 1 edge). 53% fallback rate. FLAN-T5-XL entity extraction at 3B insufficient.")
    c1, c2, c3 = st.columns(3)
    with c1: st.metric("Entities", str(t2_extra.get("entities", 180)))
    with c2: st.metric("Graph Nodes", str(t2_extra.get("nodes", 0)))
    with c3: st.metric("Fallback", f"{t2_extra.get('fallback_pct', 0):.0f}%")
    gd = pd.DataFrame({"Graph-1hop": [t2_best.get('faithfulness',0), t2_best.get('hit_rate@5',0)],
                        "Dense": [baseline.get('faithfulness',0), baseline.get('hit_rate@5',0)]},
                       index=["Faithfulness", "Hit Rate@5"])
    st.bar_chart(gd)

elif page == "Track 3: Context & Memory":
    st.title("Track 3: Context & Memory Optimization")
    t3w = sum(1 for m in METRICS if t3_best.get(m,0) > baseline.get(m,0)+0.01)
    t3l = sum(1 for m in METRICS if t3_best.get(m,0) < baseline.get(m,0)-0.01)
    st.subheader(f"Final: Track 3 wins {t3w}/5 | Baseline wins {t3l}/5 | Ties {5-t3w-t3l}/5")
    rows = []
    for i,m in enumerate(METRICS):
        b, t = baseline.get(m,0), t3_best.get(m,0)
        rows.append({"Metric": METRIC_LABELS[i], "Baseline": f"{b:.3f}", "Track 3": f"{t:.3f}", "Delta": f"{t-b:+.3f}"})
    rows.append({"Metric": "Latency (ms)", "Baseline": f"{baseline.get('latency',0):.0f}",
                 "Track 3": f"{t3_best.get('latency',0):.0f}",
                 "Delta": f"{t3_best.get('latency',0)-baseline.get('latency',0):+.0f}"})
    st.dataframe(pd.DataFrame(rows).set_index("Metric"), use_container_width=True)
    st.success("Generative compression: 91% token reduction (922>83 tok), faith 0.690 (+0.045). Structured prompt +0.045.")
    c1, c2, c3 = st.columns(3)
    with c1: st.metric("Best Faith", f"{t3_best.get('faithfulness',0):.3f}")
    with c2: st.metric("Latency", f"{t3_best.get('latency',0):.0f} ms", "-522ms")
    with c3: st.metric("Cache Hit Rate", f"{t3_cache.get('hit_rate',0):.0f}%")

    with st.expander("S2: Prompt Versioning"):
        pd_data = pd.DataFrame({"Faithfulness": [0.600, 0.645, 0.620, 0.640],
            "Latency (ms)": [1304, 1886, 5542, 7360]}, index=["baseline", "structured", "citation", "cot"])
        st.dataframe(pd_data.style.highlight_max(color="#d4edda", subset=["Faithfulness"])
                     .highlight_min(color="#d4edda", subset=["Latency (ms)"]), use_container_width=True)
        st.info("Structured prompt wins: 'You are a legal expert. Answer concisely.'")

    with st.expander("S3: Compression"):
        cd = pd.DataFrame({"Faithfulness": [0.645, 0.605, 0.690], "Avg Tokens": [922, 188, 83],
            "Latency (ms)": [1863, 1313, 1100]}, index=["none", "extractive", "generative"])
        st.dataframe(cd.style.highlight_max(color="#d4edda", subset=["Faithfulness"])
                     .highlight_min(color="#d4edda", subset=["Avg Tokens", "Latency (ms)"]), use_container_width=True)
        st.success(f"Generative: 91% token reduction, faith 0.690 (+0.045 vs none)")

    with st.expander("S4: Memory"):
        md = pd.DataFrame({"Faithfulness": [0.600, 0.600, 0.600], "Relevance": [0.620, 0.500, 0.220],
            "Latency (ms)": [2061, 3188, 814]}, index=["none", "buffer", "summary"])
        st.dataframe(md, use_container_width=True)

    with st.expander("S5: Semantic Cache"):
        c1, c2 = st.columns(2)
        with c1:
            st.metric("Hit Rate", f"{t3_cache.get('hit_rate',0):.0f}%")
            st.metric("Speedup", f"{t3_cache.get('speedup',0):.0f}x")
        with c2:
            st.metric("Hits/Misses", "10/30")
            st.metric("Faith Delta", f"{t3_cache.get('faith_delta',0):+.3f}")
        st.info("33.3% hit rate, 0ms hit latency, minimal faithfulness cost")

    with st.expander("S6: Full-Context"):
        st.error("NOT FEASIBLE: 53,044 chunks = 5.3M tokens vs 1,024 context = 5,180x over capacity")

elif page == "3-Track Comparison":
    st.title("3-Track Comparison")
    rows = []
    for i,m in enumerate(METRICS):
        rows.append({"Metric": METRIC_LABELS[i], "Baseline": f"{baseline.get(m,0):.3f}",
            "T1 Retrieval": f"{t1_best.get(m,0):.3f}", "T2 Graph": f"{t2_best.get(m,0):.3f}",
            "T3 Context": f"{t3_best.get(m,0):.3f}"})
    rows.append({"Metric": "Latency (ms)", "Baseline": f"{baseline.get('latency',0):.0f}",
        "T1 Retrieval": f"{t1_best.get('latency',0):.0f}",
        "T2 Graph": f"{t2_best.get('latency',0):.0f}", "T3 Context": f"{t3_best.get('latency',0):.0f}"})
    st.dataframe(pd.DataFrame(rows).set_index("Metric"), use_container_width=True)

    st.subheader("Chart Comparison")
    chart = pd.DataFrame({ml: [baseline.get(m,0), t1_best.get(m,0), t2_best.get(m,0), t3_best.get(m,0)]
        for ml,m in zip(METRIC_LABELS, METRICS)}, index=["Baseline", "T1", "T2", "T3"])
    st.bar_chart(chart.T)
    st.bar_chart(pd.DataFrame(latency_data).set_index("Config"))

    w1 = sum(1 for m in METRICS if t1_best.get(m,0) > baseline.get(m,0)+0.01)
    w2 = sum(1 for m in METRICS if t2_best.get(m,0) > baseline.get(m,0)+0.01)
    w3 = sum(1 for m in METRICS if t3_best.get(m,0) > baseline.get(m,0)+0.01)
    st.dataframe(pd.DataFrame({"Wins": [f"{w1}/5", f"{w2}/5", f"{w3}/5"],
        "Best Metric": ["Hit Rate@5 (+0.267)", "None (baseline wins 4/5)", "Faithfulness (+0.110)"]},
        index=["T1 Retrieval", "T2 Graph", "T3 Context"]), use_container_width=True)

    st.markdown("---")
    st.success("**For hit_rate**: Track 1 — hybrid retrieval + reranking")
    st.success("**For faithfulness**: Track 3 — structured prompt + generative compression")
    st.info("**For latency**: Track 3 — generative compression, 31% less than baseline")
    st.warning("**Graph retrieval not recommended at 3B** — entity extraction insufficient")
