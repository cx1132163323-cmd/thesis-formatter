import os

import streamlit as st

from format_thesis import run_formatter


st.set_page_config(
    page_title="硕士论文格式助手",
    page_icon="🎓",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    :root {
        --accent: #b5542f;
        --accent-dark: #8f3e20;
        --muted: #5b6770;
        --line: rgba(31,41,51,0.09);
        --paper: rgba(255,252,246,0.96);
        --shadow: 0 20px 56px rgba(64,45,24,0.13);
        --shadow-sm: 0 6px 20px rgba(64,45,24,0.08);
    }

    html, body, [class*="css"], p, div, span, label {
        font-family: "PingFang SC", "Microsoft YaHei", "Noto Sans CJK SC",
                     "Hiragino Sans GB", Arial, sans-serif !important;
        color: #1f2933;
    }

    .stApp {
        background:
            radial-gradient(ellipse at top left, rgba(210,115,62,0.18) 0%, transparent 40%),
            radial-gradient(ellipse at top right, rgba(196,164,132,0.20) 0%, transparent 35%),
            linear-gradient(180deg, #f9f4ec 0%, #f4efe7 100%);
    }

    .block-container {
        max-width: 860px !important;
        padding-top: 2rem !important;
        padding-bottom: 4rem !important;
    }

    /* ── cards ── */
    .card {
        background: var(--paper);
        border: 1px solid var(--line);
        border-radius: 24px;
        box-shadow: var(--shadow);
        padding: 2rem 2.2rem;
        margin-bottom: 1.2rem;
    }

    .card-sm {
        background: var(--paper);
        border: 1px solid var(--line);
        border-radius: 20px;
        box-shadow: var(--shadow-sm);
        padding: 1.5rem 1.6rem;
        margin-bottom: 1.2rem;
    }

    /* ── hero ── */
    .eyebrow {
        font-size: 0.78rem;
        letter-spacing: 0.14em;
        text-transform: uppercase;
        color: var(--accent);
        margin-bottom: 0.7rem;
        font-weight: 600;
    }

    .hero-title {
        font-size: 2.4rem;
        font-weight: 800;
        line-height: 1.12;
        margin: 0 0 0.8rem;
        letter-spacing: -0.01em;
        color: #1a1a1a;
    }

    .hero-title .hl {
        color: var(--accent-dark);
    }

    .hero-sub {
        font-size: 0.98rem;
        color: var(--muted);
        line-height: 1.80;
        margin-bottom: 1.2rem;
    }

    .tag-row {
        display: flex;
        flex-wrap: wrap;
        gap: 0.5rem;
    }

    .tag {
        border: 1px solid rgba(181,84,47,0.22);
        background: rgba(181,84,47,0.07);
        color: var(--accent-dark);
        border-radius: 999px;
        padding: 0.35rem 0.85rem;
        font-size: 0.85rem;
        font-weight: 500;
    }

    /* ── section title ── */
    .sec-title {
        font-size: 1.02rem;
        font-weight: 700;
        margin-bottom: 0.9rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
        color: #1a1a1a;
    }

    .sec-title::before {
        content: "";
        display: inline-block;
        width: 4px;
        height: 17px;
        background: linear-gradient(to bottom, #bc5c35, #8f3e20);
        border-radius: 3px;
        flex-shrink: 0;
    }

    /* ── tip ── */
    .tip {
        border-left: 3px solid var(--accent);
        background: rgba(181,84,47,0.07);
        border-radius: 12px;
        padding: 0.9rem 1rem;
        margin-top: 0.9rem;
        color: var(--muted);
        font-size: 0.93rem;
        line-height: 1.75;
    }

    /* ── steps ── */
    .step-grid {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 0.85rem;
    }

    .step {
        background: rgba(255,255,255,0.65);
        border: 1px solid var(--line);
        border-radius: 18px;
        padding: 1.05rem;
    }

    .step-num {
        font-size: 0.75rem;
        font-weight: 700;
        letter-spacing: 0.08em;
        color: var(--accent);
        margin-bottom: 0.38rem;
    }

    .step-name {
        font-weight: 700;
        font-size: 0.95rem;
        margin-bottom: 0.28rem;
        color: #1a1a1a;
    }

    .step-desc {
        color: var(--muted);
        font-size: 0.88rem;
        line-height: 1.65;
    }

    /* ── spec grid ── */
    .spec-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 0.55rem 1.2rem;
    }

    .spec-item {
        display: flex;
        align-items: flex-start;
        gap: 0.5rem;
        font-size: 0.90rem;
        color: var(--muted);
        line-height: 1.6;
    }

    .dot {
        width: 6px;
        height: 6px;
        border-radius: 50%;
        background: var(--accent);
        flex-shrink: 0;
        margin-top: 0.45em;
    }

    /* ── log ── */
    .log-line {
        font-size: 0.88rem;
        padding: 0.32rem 0;
        border-bottom: 1px dashed rgba(31,41,51,0.08);
        color: var(--muted);
        font-family: "Courier New", monospace !important;
    }

    /* ── buttons ── */
    .stButton > button,
    .stDownloadButton > button {
        border-radius: 999px !important;
        border: none !important;
        background: linear-gradient(135deg, #bc5c35 0%, #8f3e20 100%) !important;
        color: #fff !important;
        font-weight: 700 !important;
        padding: 0.72rem 1.5rem !important;
        font-size: 0.98rem !important;
        box-shadow: 0 10px 28px rgba(143,62,32,0.26) !important;
    }

    .stButton > button:hover,
    .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #c9673f 0%, #7f3418 100%) !important;
    }

    /* ── footer ── */
    .footer {
        text-align: center;
        color: rgba(91,103,112,0.50);
        font-size: 0.80rem;
        margin-top: 2.5rem;
        line-height: 1.9;
    }

    /* ── responsive ── */
    @media (max-width: 700px) {
        .hero-title { font-size: 1.85rem; }
        .step-grid { grid-template-columns: 1fr; }
        .spec-grid { grid-template-columns: 1fr; }
        .card, .card-sm { padding: 1.3rem; border-radius: 18px; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── Hero ──
st.markdown(
    """
    <div class="card">
        <div class="eyebrow">Master Thesis Formatter</div>
        <div class="hero-title">硕士论文<br><span class="hl">格式助手</span></div>
        <div class="hero-sub">
            上传 Word 原稿，自动整理页面、标题、表格、目录字段与页眉页脚。<br>
            几秒钟完成，下载即用。
        </div>
        <div class="tag-row">
            <span class="tag">📄 上传 .docx</span>
            <span class="tag">⚡ 自动格式化</span>
            <span class="tag">📥 即刻下载</span>
            <span class="tag">🆓 完全免费</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Upload ──
st.markdown('<div class="card-sm">', unsafe_allow_html=True)
st.markdown('<div class="sec-title">开始处理</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("上传 Word 文件（.docx）", type=["docx"], label_visibility="visible")
title = st.text_input(
    "论文标题（可选，留空则自动识别）",
    placeholder="例：Supply Chain Optimization in Manufacturing Enterprises",
)

if uploaded_file is not None:
    st.caption(f"已选文件：{uploaded_file.name}")

st.markdown(
    """
    <div class="tip">
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
    <div class="card-sm">
        <div class="sec-title">使用流程</div>
        <div class="step-grid">
            <div class="step">
                <div class="step-num">STEP 01</div>
                <div class="step-name">上传原稿</div>
                <div class="step-desc">选择本地 .docx 文件，支持直接从 Word 导出的格式。</div>
            </div>
            <div class="step">
                <div class="step-num">STEP 02</div>
                <div class="step-name">自动处理</div>
                <div class="step-desc">统一段落、标题、三线表、目录字段、奇偶页眉页脚。</div>
            </div>
            <div class="step">
                <div class="step-num">STEP 03</div>
                <div class="step-name">下载检查</div>
                <div class="step-desc">下载后在 Word 中刷新目录页码，最后人工核对一遍。</div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Spec ──
st.markdown(
    """
    <div class="card-sm">
        <div class="sec-title">自动处理的格式项</div>
        <div class="spec-grid">
            <div class="spec-item"><div class="dot"></div>A4 页面 + 版心边距</div>
            <div class="spec-item"><div class="dot"></div>正文 Times New Roman 12pt，1.25 倍行距</div>
            <div class="spec-item"><div class="dot"></div>一至四级标题字体与对齐</div>
            <div class="spec-item"><div class="dot"></div>图题 / 表题样式与编号</div>
            <div class="spec-item"><div class="dot"></div>三线表边框（1.5 pt / 0.75 pt）</div>
            <div class="spec-item"><div class="dot"></div>自动目录 / 图表目录字段</div>
            <div class="spec-item"><div class="dot"></div>奇偶页眉（论文题目 / 章节名）</div>
            <div class="spec-item"><div class="dot"></div>居中页码页脚</div>
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

        st.markdown('<div class="card-sm">', unsafe_allow_html=True)
        st.markdown('<div class="sec-title">处理日志</div>', unsafe_allow_html=True)
        for line in logs:
            st.markdown(f'<div class="log-line">▸ {line}</div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

# ── Footer ──
st.markdown(
    """
    <div class="footer">
        硕士论文格式助手 · 仅供学习参考<br>
        格式依据相关学位论文写作规范
    </div>
    """,
    unsafe_allow_html=True,
)
