# RAG Engineering Study Dashboard

Interactive dashboard for the RAG Engineering Comparative Study (3 tracks) on LegalBench-RAG.

## Quick Start

```bash
pip install streamlit pandas numpy
streamlit run streamlit_app.py
```

The dashboard loads JSON result files from the `results/` folder. If no files are found, it displays sample data from the actual Kaggle run.

## Data

Place the 18 JSON result files from your notebook run into `results/`:

```
results/
  track1_baseline.json      track2_chunking.json      track3_prompt_versioning.json
  track1_chunking.json       track2_entities.json       track3_compression.json
  track1_retrieval.json      track2_graph_search.json   track3_memory.json
  track1_reranking.json      track2_communities.json    track3_caching.json
  track1_transform.json      track2_routed.json         track3_fullcontext.json
  track1_final.json          track2_final.json          track3_final.json
```

## Deploy on Streamlit Cloud

1. Push this repo to GitHub
2. Go to https://share.streamlit.io
3. Connect your GitHub repo
4. Set `streamlit_app.py` as the entry point
5. Deploy — no secrets or config needed

## Tracks

- **Track 1**: Retrieval Engineering — chunking, dense/hybrid/keyword retrieval, reranking, query transforms
- **Track 2**: Graph-Based Retrieval — entity extraction, graph construction, community search, query routing
- **Track 3**: Context & Memory — prompt versioning, context compression, memory strategies, semantic caching

## Tech

- **LLM**: Google FLAN-T5-XL (3B) — local, no API cost
- **Embedding**: all-MiniLM-L6-v2 (384d)
- **Reranker**: cross-encoder/ms-marco-MiniLM-L-6-v2
- **Graph**: NetworkX entity co-occurrence
- **All zero cost**: Kaggle T4 GPU free tier
