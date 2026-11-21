import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import os
from io import BytesIO
from groq import Groq
from dotenv import load_dotenv

# =====================================================
# 1. PAGE CONFIG & DARK THEME CSS
# =====================================================
st.set_page_config(
    page_title="Budget vs. Actuals AI",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    /* Paksa latar belakang utama dan sidebar menjadi hitam elegan */
    .stApp {
        background-color: #090d16;
        color: #f8fafc;
    }
    section[data-testid="stSidebar"] {
        background-color: #0f172a;
        border-right: 1px solid #1e293b;
    }
    /* Kartu KPI Dark Mode */
    div[data-testid="metric-container"] {
        background-color: #111827;
        border: 1px solid #1f2937;
        padding: 18px 22px;
        border-radius: 14px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.4);
    }
    div[data-testid="metric-container"] label {
        color: #94a3b8 !important;
        font-size: 14px !important;
    }
    div[data-testid="metric-container"] div[data-testid="stMetricValue"] {
        color: #f8fafc !important;
        font-weight: 700;
    }
    /* Header Banner Dark Mode */
    .main-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e1b4b 50%, #0f172a 100%);
        border: 1px solid #312e81;
        padding: 26px 32px;
        border-radius: 16px;
        margin-bottom: 24px;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.5);
    }
    .main-header h1 {
        color: #ffffff;
        margin: 0;
        font-size: 28px;
        letter-spacing: -0.5px;
    }
    .main-header p {
        color: #94a3b8;
        margin: 6px 0 0 0;
        font-size: 15px;
    }
    /* Styling Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #111827;
        padding: 6px;
        border-radius: 12px;
        border: 1px solid #1f2937;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        color: #94a3b8;
        padding: 8px 18px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1e293b !important;
        color: #38bdf8 !important;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# =====================================================
# 2. LOAD GROQ API KEY
# =====================================================
load_dotenv()
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

try:
    if "GROQ_API_KEY" in st.secrets:
        GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
except Exception:
    pass

if not GROQ_API_KEY:
    st.error("🚨 GROQ_API_KEY belum diatur di Streamlit Secrets!")
    st.stop()

client = Groq(api_key=GROQ_API_KEY)

# Fungsi helper download Excel
def to_excel_bytes(dataframe):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Variance_Analysis")
    return output.getvalue()

# =====================================================
# 3. SIDEBAR
# =====================================================
with st.sidebar:
    st.markdown("### ⚙️ Control Panel")
    st.markdown("---")
    
    selected_model = st.selectbox(
        "🤖 Model AI (Groq)",
        ["openai/gpt-oss-20b", "openai/gpt-oss-120b", "llama-3.3-70b-versatile"],
        index=0
    )
    
    uploaded_file = st.file_uploader(
        "📂 Upload Dataset (.xlsx)",
        type=["xlsx"],
        help="Wajib memiliki kolom: Category, Budget, dan Actual"
    )

    # Tombol unduh template jika user belum punya file
    sample_df = pd.DataFrame({
        "Category": ["Sales Revenue", "COGS", "Marketing", "Salaries", "IT & Software", "Utilities"],
        "Budget": [150000000, 60000000, 15000000, 35000000, 8000000, 5000000],
        "Actual": [145000000, 64000000, 18500000, 35000000, 7200000, 5600000]
    })
    st.download_button(
        label="📥 Download Template Excel",
        data=to_excel_bytes(sample_df),
        file_name="Template_Budget_vs_Actual.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# =====================================================
# 4. HEADER BANNER
# =====================================================
st.markdown("""
<div class="main-header">
    <h1>📊 Budget vs. Actuals AI – Variance Intelligence</h1>
    <p>Analisis Varians Anggaran Real-Time, Visualisasi Interaktif & Asisten Keuangan AI</p>
</div>
""", unsafe_allow_html=True)

# =====================================================
# 5. MAIN CONTENT
# =====================================================
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        required_columns = ["Category", "Budget", "Actual"]

        if not all(col in df.columns for col in required_columns):
            st.error("⚠️ File Excel wajib memiliki kolom: 'Category', 'Budget', dan 'Actual'!")
            st.stop()

        df["Budget"] = pd.to_numeric(df["Budget"], errors="coerce")
        df["Actual"] = pd.to_numeric(df["Actual"], errors="coerce")

        # Filter Kategori di Sidebar
        with st.sidebar:
            st.markdown("---")
            selected_categories = st.multiselect(
                "🔍 Filter Kategori:",
                options=df["Category"].unique().tolist(),
                default=df["Category"].unique().tolist()
            )

        df_filtered = df[df["Category"].isin(selected_categories)].copy()

        # Hitung Variance, Variance %, dan Status
        df_filtered["Variance"] = df_filtered["Actual"] - df_filtered["Budget"]
        df_filtered["Variance %"] = (df_filtered["Variance"] / df_filtered["Budget"].replace(0, float("nan"))) * 100
        df_filtered["Status"] = df_filtered["Variance"].apply(
            lambda x: "🔴 Over Budget" if x > 0 else ("🟢 Under Budget" if x < 0 else "⚪ On Target")
        )

        # KPI METRICS
        total_budget = df_filtered["Budget"].sum()
        total_actual = df_filtered["Actual"].sum()
        total_variance = total_actual - total_budget
        variance_pct_total = (total_variance / total_budget * 100) if total_budget != 0 else 0
        absorption_rate = (total_actual / total_budget * 100) if total_budget != 0 else 0

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("💰 Total Budget", f"{total_budget:,.0f}")
        col2.metric("💳 Total Actual", f"{total_actual:,.0f}")
        col3.metric(
            "📉 Total Variance",
            f"{total_variance:,.0f}",
            delta=f"{variance_pct_total:.1f}%",
            delta_color="inverse"
        )
        col4.metric("📊 Rasio Realisasi", f"{absorption_rate:.1f}%")

        st.markdown("<br>", unsafe_allow_html=True)

        # TABS
        tab1, tab2, tab3 = st.tabs(["📈 Visualisasi Grafik", "📋 Tabel & Export Data", "🤖 AI Financial Analyst"])

        # TAB 1: GRAFIK DARK MODE
        with tab1:
            col_c1, col_c2 = st.columns(2)

            with col_c1:
                fig_grouped = go.Figure(data=[
                    go.Bar(
                        name="Budget",
                        x=df_filtered["Category"],
                        y=df_filtered["Budget"],
                        marker_color="#38bdf8",
                        text=df_filtered["Budget"],
                        texttemplate="%{text:.2s}",
                        textposition="outside"
                    ),
                    go.Bar(
                        name="Actual",
                        x=df_filtered["Category"],
                        y=df_filtered["Actual"],
                        marker_color="#a855f7",
                        text=df_filtered["Actual"],
                        texttemplate="%{text:.2s}",
                        textposition="outside"
                    )
                ])
                fig_grouped.update_layout(
                    title="Perbandingan Budget vs. Actual",
                    barmode="group",
                    template="plotly_dark",
                    paper_bgcolor="#111827",
                    plot_bgcolor="#111827",
                    xaxis_tickangle=-25,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    margin=dict(t=60, b=40, l=20, r=20),
                    height=440
                )
                st.plotly_chart(fig_grouped, use_container_width=True)

            with col_c2:
                colors = ["#f43f5e" if v > 0 else "#10b981" for v in df_filtered["Variance"]]
                fig_var = go.Figure(go.Bar(
                    x=df_filtered["Variance"],
                    y=df_filtered["Category"],
                    orientation="h",
                    marker_color=colors,
                    text=df_filtered["Variance"],
                    texttemplate="%{text:,.0f}",
                    textposition="auto"
                ))
                fig_var.add_vline(x=0, line_width=1.5, line_dash="dash", line_color="#64748b")
                fig_var.update_layout(
                    title="Selisih (Variance) per Kategori",
                    template="plotly_dark",
                    paper_bgcolor="#111827",
                    plot_bgcolor="#111827",
                    yaxis=dict(autorange="reversed"),
                    margin=dict(t=60, b=40, l=20, r=20),
                    height=440
                )
                st.plotly_chart(fig_var, use_container_width=True)

        # TAB 2: TABEL & EXPORT
        with tab2:
            col_t1, col_t2 = st.columns([3, 1])
            with col_t1:
                st.subheader("📋 Rincian Perhitungan Variance")
            with col_t2:
                st.download_button(
                    label="💾 Download Hasil (.xlsx)",
                    data=to_excel_bytes(df_filtered),
                    file_name="Hasil_Variance_Analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            st.dataframe(
                df_filtered.style.format({
                    "Budget": "{:,.2f}",
                    "Actual": "{:,.2f}",
                    "Variance": "{:,.2f}",
                    "Variance %": "{:.2f}%"
                }),
                use_container_width=True,
                height=380
            )

        # TAB 3: AI ANALYSIS & CHAT
        with tab3:
            col_ai1, col_ai2 = st.columns([1.2, 1])

            with col_ai1:
                st.subheader("📑 Executive AI Commentary")
                if st.button("✨ Generate Analisis AI", type="primary"):
                    with st.spinner("AI sedang menganalisis data keuangan..."):
                        response = client.chat.completions.create(
                            messages=[
                                {
                                    "role": "system",
                                    "content": "You are a senior financial analyst. Provide a sharp, structured executive summary, highlight overspending/underspending risks, and give actionable recommendations in Indonesian."
                                },
                                {
                                    "role": "user",
                                    "content": f"Analisis data Budget vs Actual berikut:\n{df_filtered.to_string(index=False)}"
                                }
                            ],
                            model=selected_model,
                        )
                        st.session_state.variance_analysis = response.choices[0].message.content

                if "variance_analysis" in st.session_state:
                    st.markdown(st.session_state.variance_analysis)

            with col_ai2:
                st.subheader("💬 Tanya Jawab AI")
                user_query = st.text_input("🔍 Tanyakan tentang data Anda (misal: 'Kategori mana yang paling berisiko?'):")
                if st.button("Kirim Pertanyaan") and user_query:
                    with st.spinner("Menyiapkan jawaban..."):
                        chat_response = client.chat.completions.create(
                            messages=[
                                {
                                    "role": "system",
                                    "content": "You are an AI financial analyst. Answer concisely and clearly in Indonesian based on the dataset."
                                },
                                {
                                    "role": "user",
                                    "content": f"Data:\n{df_filtered.to_string(index=False)}\n\nPertanyaan: {user_query}"
                                }
                            ],
                            model=selected_model,
                        )
                        st.success(chat_response.choices[0].message.content)

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
else:
    st.info("👈 Unggah file Excel (`.xlsx`) di panel kiri atau klik **Download Template Excel** untuk mencoba.")
