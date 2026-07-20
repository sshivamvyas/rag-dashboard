import os

TARGET = r"C:\Users\adm\.openclaw-autoclaw\agents\task-2-rag\workspace\rag-study\DELIVERY\streamlit_app.py"
code_lines = []
def L(s=""):
    code_lines.append(s)

L("# === PAGE RENDERERS ===")
L("")
L("def page_executive(data, _sample):")
L("    st.header('Executive Summary')")
L("    bl = data.get('track1_baseline') or data.get('track1_final') or {}")
L("    t1 = data.get('track1_final') or {}")
L("    t2 = data.get('track2_final') or {}")
L("    t3 = data.get('track3_final') or {}")
L("    configs = {'Baseline': bl, 'Track 1 (Retrieval)': t1, 'Track 2 (Graph)': t2, 'Track 3 (Context & Mem)': t3}")
L("")
L("    st.subheader('Best-in-Class Metrics')")
L("    c1, c2, c3_ = st.columns(3)")
L("    best_f = max(configs.items(), key=lambda kv: kv[1].get('faithfulness', 0) if kv[1] else 0)")
L("    best_l = min(configs.items(), key=lambda kv: kv[1].get('latency_ms', 999999) if kv[1] else 999999)")
L("    best_h = max(configs.items(), key=lambda kv: kv[1].get('hit_rate@5', 0) if kv[1] else 0)")
L("    with c1: render_dark_card('Best Faithfulness', str(best_f[1].get('faithfulness','---')), best_f[0])")
L("    with c2: render_dark_card('Lowest Latency', f\"{int(best_l[1].get('latency_ms',0))} ms\", best_l[0])")
L("    with c3_: render_dark_card('Best Hit Rate@5', str(best_h[1].get('hit_rate@5','---')), best_h[0])")
L("")
L("    st.subheader('Final Comparison - All Configurations')")
L("    df = build_comparison_df(configs)")
L("    styled = _highlight_best_worst(df.set_index('Configuration').reset_index(), METRICS_BASIC, ['latency_ms'])")
L("    st.dataframe(styled, use_container_width=True, hide_index=True)")
L("")
L("    st.subheader('Metric Comparison Across Configurations')")
L("    render_metric_bar_chart(df, METRICS_BASIC, 'Per-metric performance across all configurations')")
L("")
L("    st.success('Key takeaway: Track 3 achieves highest faithfulness (0.727) and lowest latency (1145ms). Track 1 excels in Precision (1.000) and Hit Rate@5 (0.933). The Graph approach (Track 2) underperforms due to sparse entity extraction.')")
L("")

with open(TARGET, "a", encoding="utf-8") as f:
    f.write("\n".join(code_lines))

print(f"Part 3a appended {len(code_lines)} lines, file now {os.path.getsize(TARGET)} bytes")
