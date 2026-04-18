import os

import streamlit as st

from format_thesis import run_formatter


st.set_page_config(
    page_title="湖南大学硕士论文格式助手",
    page_icon="🎓",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@400;500;700;800&family=Space+Grotesk:wght@500;700&display=swap');

    :root {
        --bg-main: #f4efe7;
        --paper: rgba(255, 252, 246, 0.96);
        --ink: #1f2933;
        --muted: #5b6770;
        --accent: #b5542f;
        --accent-dark: #8f3e20;
        --accent-light: rgba(181, 84, 47, 0.10);
        --line: rgba(31, 41, 51, 0.09);
        --shadow: 0 24px 64px rgba(64, 45, 24, 0.13);
        --shadow-sm: 0 8px 24px rgba(64, 45, 24, 0.08);
    }

    html, body, [class*="css"] {
        font-family: "Noto Sans SC", sans-serif;
        color: var(--ink);
    }

    .stApp {
        background:
            radial-gradient(ellipse at top left, rgba(210, 115, 62, 0.20) 0%, transparent 42%),
            radial-gradient(ellipse at top right, rgba(196, 164, 132, 0.24) 0%, transparent 36%),
            radial-gradient(ellipse at bottom center, rgba(181, 84, 47, 0.06) 0%, transparent 60%),
            linear-gradient(180deg, #f9f4ec 0%, var(--bg-main) 100%);
        min-height: 100vh;
    }

    .block-container {
        max-width: 880px;
        padding-top: 2.4rem;
        padding-bottom: 4rem;
    }

    /* ── cards ── */
    .hero-card, .panel-card, .log-card, .badge-bar {
        background: var(--paper);
        border: 1px solid var(--line);
        box-shadow: var(--shadow);
        border-radius: 28px;
    }

    .hero-card {
        padding: 2.4rem 2.4rem 1.8rem;
        margin-bottom: 1.4rem;
        position: relative;
        overflow: hidden;
    }

    .hero-card::before {
        content: "";
        position: absolute;
        top: -60px; right: -60px;
        width: 260px; height: 260px;
        border-radius: 50%;
        background: radial-gradient(circle, rgba(181,84,47,0.10) 0%, transparent 70%);
        pointer-events: none;
    }

    /* ── eyebrow / title ── */
    .eyebrow {
        font-family: "Space Grotesk", sans-serif;
        font-size: 0.82rem;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        color: var(--accent);
        margin-bottom: 0.65rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }

    .eyebrow::before {
        content: "";
        display: inline-block;
        width: 22px; height: 2px;
        background: var(--accent);
        border-radius: 2px;
    }

    .hero-title {
        font-size: 2.55rem;
        line-height: 1.08;
        font-weight: 800;
        margin: 0 0 0.8rem;
        letter-spacing: -0.02em;
    }

    .hero-title span {
        background: linear-gradient(135deg, #bc5c35 0%, #8f3e20 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }

    .hero-subtitle {
        font-size: 1.02rem;
        color: var(--muted);
        line-height: 1.80;
        margin-bottom: 1.3rem;
        max-width: 560px;
    }

    .hero-tags {
        display: flex;
        flex-wrap: wrap;
        gap: 0.55rem;
    }

    .hero-tag {
        border: 1px solid rgba(181, 84, 47, 0.20);
        background: rgba(181, 84, 47, 0.08);
        color: var(--accent-dark);
        border-radius: 999px;
        padding: 0.38rem 0.9rem;
        font-size: 0.88rem;
        font-weight: 500;
    }

    /* ── panel ── */
    .panel-card {
        padding: 1.6rem 1.6rem 1.3rem;
        margin-top: 1.2rem;
        box-shadow: var(--shadow-sm);
    }

    .section-title {
        font-size: 1.05rem;
        font-weight: 700;
        margin-bottom: 0.85rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }

    .section-title::before {
        content: "";
        display: inline-block;
        width: 4px; height: 18px;
        background: linear-gradient(to bottom, #bc5c35, #8f3e20);
        border-radius: 4px;
    }

    /* ── tip ── */
    .tip-box {
        border-left: 3px solid var(--accent);
        background: var(--accent-light);
        border-radius: 14px;
        padding: 1rem 1.1rem;
        margin-top: 1rem;
        color: var(--muted);
        line-height: 1.75;
        font-size: 0.95rem;
    }

    /* ── steps ── */
    .step-grid {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 0.9rem;
        margin-top: 0.8rem;
    }

    .step-card {
        background: rgba(255, 255, 255, 0.68);
        border: 1px solid var(--line);
        border-radius: 20px;
        padding: 1.1rem;
        transition: box-shadow 0.2s;
    }

    .step-card:hover {
        box-shadow: var(--shadow-sm);
    }

    .step-no {
        font-family: "Space Grotesk", sans-serif;
        color: var(--accent);
        font-size: 0.80rem;
        font-weight: 700;
        letter-spacing: 0.08em;
        margin-bottom: 0.4rem;
    }

    .step-name {
        font-weight: 700;
        font-size: 0.97rem;
        margin-bottom: 0.3rem;
    }

    .step-desc {
        color: var(--muted);
        font-size: 0.90rem;
        line-height: 1.65;
    }

    /* ── spec list ── */
    .spec-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 0.6rem;
        margin-top: 0.8rem;
    }

    .spec-item {
        display: flex;
        align-items: flex-start;
        gap: 0.5rem;
        font-size: 0.92rem;
        color: var(--muted);
        line-height: 1.6;
    }

    .spec-dot {
        width: 7px; height: 7px;
        border-radius: 50%;
        background: var(--accent);
        margin-top: 0.42em;
        flex-shrink: 0;
    }

    /* ── upload / inputs ── */
    .stFileUploader > div, .stTextInput > div > div {
        background: rgba(255, 255, 255, 0.75) !important;
        border-radius: 16px !important;
    }

    /* ── buttons ── */
    .stButton > button, .stDownloadButton > button {
        border-radius: 999px !important;
        border: none !important;
        background: linear-gradient(135deg, #bc5c35 0%, #8f3e20 100%) !important;
        color: white !important;
        font-weight: 700 !important;
        font-family: "Noto Sans SC", sans-serif !important;
        padding: 0.75rem 1.6rem !important;
        font-size: 1rem !important;
        box-shadow: 0 12px 30px rgba(143, 62, 32, 0.26) !important;
        transition: all 0.2s !important;
    }

    .stButton > button:hover, .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #c9673f 0%, #7f3418 100%) !important;
        box-shadow: 0 16px 36px rgba(143, 62, 32, 0.32) !important;
        transform: translateY(-1px) !important;
    }

    /* ── log ── */
    .log-card {
        padding: 1.2rem 1.2rem 0.8rem;
        margin-top: 1.2rem;
        box-shadow: var(--shadow-sm);
    }

    .log-line {
        font-family: "Space Grotesk", monospace;
        font-size: 0.90rem;
        padding: 0.36rem 0;
        border-bottom: 1px dashed rgba(31, 41, 51, 0.07);
        color: var(--muted);
    }

    /* ── footer ── */
    .footer {
        text-align: center;
        color: rgba(91, 103, 112, 0.55);
        font-size: 0.82rem;
        margin-top: 2.5rem;
        line-height: 1.9;
    }

    /* ── responsive ── */
    @media (max-width: 760px) {
        .hero-title { font-size: 1.9rem; }
        .step-grid { grid-template-columns: 1fr; }
        .spec-grid { grid-template-columns: 1fr; }
        .hero-card, .panel-card, .log-card { border-radius: 22px; padding: 1.4rem; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── Hero ──
st.markdown(
    """
    <div class="hero-card">
        <div class="eyebrow">HNU Thesis Formatter</div>
        <div class="hero-title">湖南大学<br><span>硕士论文格式助手</span></div>
        <div class="hero-subtitle">
            上传 Word 原稿，自动按湖南大学硕士学位论文规范整理页面、标题、
            表格、目录字段与页眉页脚。几秒钟完成，下载即用。
        </div>
        <div class="hero-tags">
            <div class="hero-tag">📄 上传 .docx</div>
            <div class="hero-tag">⚡ 自动格式化</div>
            <div class="hero-tag">📥 即刻下载</div>
            <div class="hero-tag">🆓 完全免费</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Upload panel ──
st.markdown('<div class="panel-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">开始处理</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("上传 Word 文件（.docx）", type=["docx"], label_visibility="visible")
title = st.text_input(
    "论文标题（可选，留空则自动识别）",
    placeholder="例：Supply Chain Optimization in Manufacturing Enterprises",
)

if uploaded_file is not None:
    st.caption(f"已选文件：{uploaded_file.name}")

st.markdown(
    """
    <div class="tip-box">
        💡 处理完成后，用 Word 打开下载的文件，按 <strong>Ctrl + A</strong>，
        再按 <strong>F9</strong>，即可刷新目录页码和图表编号。
    </div>
    """,
    unsafe_allow_html=True,
)

start = st.button("开始格式化", type="primary", disabled=uploaded_file is None)
st.markdown("</div>", unsafe_allow_html=True)

# ── Steps ──
st.markdown(
    """
    <div class="panel-card">
        <div class="section-title">使用流程</div>
        <div class="step-grid">
            <div class="step-card">
                <div class="step-no">STEP 01</div>
                <div class="step-name">上传原稿</div>
                <div class="step-desc">选择本地 .docx 文件，支持直接从 Word 导出的格式。</div>
            </div>
            <div class="step-card">
                <div class="step-no">STEP 02</div>
                <div class="step-name">自动处理</div>
                <div class="step-desc">统一段落、标题、三线表、目录字段、奇偶页眉页脚。</div>
            </div>
            <div class="step-card">
                <div class="step-no">STEP 03</div>
                <div class="step-name">下载检查</div>
                <div class="step-desc">下载后在 Word 中刷新目录页码，最后人工核对一遍。</div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── What it covers ──
st.markdown(
    """
    <div class="panel-card">
        <div class="section-title">自动覆盖的格式项</div>
        <div class="spec-grid">
            <div class="spec-item"><div class="spec-dot"></div>A4 页面 + 版心边距</div>
            <div class="spec-item"><div class="spec-dot"></div>正文 Times New Roman 12pt，1.25 倍行距</div>
            <div class="spec-item"><div class="spec-dot"></div>一至四级标题字体与对齐</div>
            <div class="spec-item"><div class="spec-dot"></div>图题 / 表题样式与编号</div>
            <div class="spec-item"><div class="spec-dot"></div>三线表边框（1.5 pt / 0.75 pt）</div>
            <div class="spec-item"><div class="spec-dot"></div>自动目录 / 图表目录字段</div>
            <div class="spec-item"><div class="spec-dot"></div>奇偶页眉（论文题目 / 章节名）</div>
            <div class="spec-item"><div class="spec-dot"></div>居中页码页脚</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Processing ──
if start:
    if uploaded_file is None:
        st.error("请先上传 .docx 文件。")
    else:
        logs = []
        with st.spinner("正在格式化，请稍候…"):
            result = run_formatter(
                uploaded_file.getvalue(),
                title=title.strip() or None,
                log_func=logs.append,
            )

        output_name = f"{os.path.splitext(uploaded_file.name)[0]}_格式化版.docx"

        st.success("✅ 格式化完成，点击下方按钮下载。")
        st.download_button(
            "📥 下载格式化文档",
            data=result["output_bytes"],
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        st.info("如果目录页码还不完整，请用 Word 打开后按 Ctrl+A，再按 F9 刷新。")

        st.markdown('<div class="log-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">处理日志</div>', unsafe_allow_html=True)
        for line in logs:
            st.markdown(f'<div class="log-line">▸ {line}</div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

# ── Footer ──
st.markdown(
    """
    <div class="footer">
        湖南大学硕士学位论文格式助手 · 仅供学习参考<br>
        格式依据《湖南大学硕士学位论文规范》英文版
    </div>
    """,
    unsafe_allow_html=True,
)
