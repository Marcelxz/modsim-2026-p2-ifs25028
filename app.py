import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

# =====================================================
# KONFIGURASI APLIKASI
# =====================================================
st.set_page_config(
    page_title="Dashboard Analisis Kuesioner",
    layout="wide",
    page_icon="📊"
)

# --- PERBAIKAN: MENGGUNAKAN ABSOLUTE PATH ---
# Mendapatkan folder tempat script app.py ini berada
current_dir = os.path.dirname(os.path.abspath(__file__))
# Menggabungkan path folder dengan nama file
DATA_FILE = os.path.join(current_dir, "data_kuesioner.xlsx")
# --------------------------------------------

# =====================================================
# FUNGSI BANTU (HELPER FUNCTIONS)
# =====================================================
def format_persen(value):
    """Format angka ke persentase string"""
    return f"{value:.1f}%"

def format_skor(value):
    """Format angka ke skor (2 desimal)"""
    return f"{value:.2f}"

# Mapping Nilai & Warna
SKOR_MAP = {"SS": 6, "S": 5, "CS": 4, "CTS": 3, "TS": 2, "STS": 1}
URUTAN_SKALA = ["SS", "S", "CS", "CTS", "TS", "STS"]

# Palet Warna Konsisten
COLOR_MAP = {
    "SS": "#1f77b4",   # Biru Tua
    "S": "#00CC96",    # Hijau Tosca
    "CS": "#AB63FA",   # Ungu
    "CTS": "#FFA15A",  # Oranye Muda
    "TS": "#EF553B",   # Merah Muda
    "STS": "#B42020"   # Merah Tua
}

# =====================================================
# MEMUAT DATA (DATA LOADING)
# =====================================================
@st.cache_data
def muat_data():
    """
    Memuat dan mempersiapkan data dari file Excel Kuesioner
    """
    # Cek apakah file ada sebelum mencoba membaca
    if not os.path.exists(DATA_FILE):
        st.error(f"❌ File tidak ditemukan di lokasi: {DATA_FILE}")
        st.stop()

    try:
        # Load Data (Tanpa spesifik sheet_name agar membaca sheet pertama secara default)
        df = pd.read_excel(DATA_FILE, engine="openpyxl")
        
        # Hapus kolom Partisipan jika ada
        if 'Partisipan' in df.columns:
            df = df.drop(columns=['Partisipan'])
            
        # Buat dataframe versi numerik untuk perhitungan skor
        # Hanya mengganti nilai yang ada di map, sisanya biarkan (untuk handling error)
        df_numeric = df.replace(SKOR_MAP)
        
        # Pastikan data numerik benar-benar angka (coerce error to NaN)
        for col in df_numeric.columns:
            df_numeric[col] = pd.to_numeric(df_numeric[col], errors='coerce')

        return df, df_numeric

    except ImportError:
        st.error("❌ Library 'openpyxl' belum terinstall. Silakan jalankan `pip install openpyxl` di terminal.")
        return pd.DataFrame(), pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Gagal memuat data. Detail Error: {str(e)}")
        return pd.DataFrame(), pd.DataFrame()

# Load Data Utama
df, df_numeric = muat_data()

# Validasi data kosong
if df.empty:
    st.warning("⚠️ Data berhasil dimuat tapi kosong, atau terjadi error saat loading.")
    st.stop()

# =====================================================
# SIDEBAR (NAVIGASI)
# =====================================================
st.sidebar.title("📊 Kuesioner Analytics")
st.sidebar.markdown("---")

# Filter Pertanyaan (Opsional)
st.sidebar.subheader("🏷️ Filter Pertanyaan")
all_questions = df.columns.tolist()
selected_questions = st.sidebar.multiselect(
    "Pilih pertanyaan untuk analisis detail",
    all_questions,
    default=all_questions
)

# Terapkan filter kolom jika ada yang dipilih
if selected_questions:
    df_display = df[selected_questions]
    df_num_display = df_numeric[selected_questions]
else:
    df_display = df
    df_num_display = df_numeric

st.sidebar.markdown("---")

# Menu navigasi
menu = st.sidebar.radio(
    "📋 PILIH VISUALISASI",
    [
        "📊 Dashboard Utama",
        "📝 Analisis per Pertanyaan",
        "⭐ Peringkat Skor",
        "🎭 Analisis Sentimen/Kategori",
        "📋 Tabel Data Lengkap"
    ]
)

# =====================================================
# DASHBOARD UTAMA
# =====================================================
if menu == "📊 Dashboard Utama":
    st.title("📊 Dashboard Utama - Overview")
    st.markdown("---")
    
    # Hitung KPI Global
    total_responden = len(df)
    
    # Rata-rata global (mengabaikan NaN)
    rata_rata_global = df_num_display.values.flatten()
    rata_rata_global = rata_rata_global[~pd.isna(rata_rata_global)].mean()
    
    # Mencari Pertanyaan dengan Skor Tertinggi & Terendah
    mean_per_q = df_num_display.mean()
    best_q = mean_per_q.idxmax()
    best_score = mean_per_q.max()
    worst_q = mean_per_q.idxmin()
    worst_score = mean_per_q.min()

    # KPI Cards
    col1, col2 = st.columns(2)

    with col1:
        st.metric(
            label="👥 Total Responden",
            value=f"{total_responden} Orang",
            delta="Data Masuk"
        )

    with col2:
        st.metric(
            label="⭐ Rata-rata Skor Global",
            value=f"{rata_rata_global:.2f} / 6.00",
            delta="Skala 1-6"
        )

    col3, col4 = st.columns(2)

    with col3:
        st.metric(
            label="🏆 Pertanyaan Terbaik",
            value=best_q if pd.notna(best_q) else "-",
            delta=f"Skor: {best_score:.2f}" if pd.notna(best_score) else "-"
        )

    with col4:
        st.metric(
            label="⚠️ Perlu Perhatian",
            value=worst_q if pd.notna(worst_q) else "-",
            delta=f"Skor: {worst_score:.2f}" if pd.notna(worst_score) else "-",
            delta_color="inverse"
        )

    st.markdown("---")
        
    # Visualisasi Utama - Global Distribution
    col_a, col_b = st.columns(2)
    
    # Data Flatten untuk distribusi global
    all_vals = df_display.values.flatten()
    # Filter hanya nilai valid
    all_vals = [v for v in all_vals if v in URUTAN_SKALA]
    
    dist_global = pd.Series(all_vals).value_counts().reindex(URUTAN_SKALA).fillna(0).reset_index()
    dist_global.columns = ["Skala", "Jumlah"]
    
    with col_a:
        st.subheader("📊 Distribusi Jawaban Keseluruhan")
        fig = px.bar(
            dist_global,
            x="Skala",
            y="Jumlah",
            color="Skala",
            color_discrete_map=COLOR_MAP,
            text_auto=True,
            title="Frekuensi Jawaban (Semua Pertanyaan)"
        )
        fig.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    with col_b:
        st.subheader("🥧 Proporsi Jawaban")
        fig = px.pie(
            dist_global,
            names="Skala",
            values="Jumlah",
            color="Skala",
            color_discrete_map=COLOR_MAP,
            hole=0.4,
            title="Persentase Pilihan Jawaban"
        )
        fig.update_traces(textposition='inside', textinfo='percent+label')
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)

# =====================================================
# ANALISIS PER PERTANYAAN
# =====================================================
elif menu == "📝 Analisis per Pertanyaan":
    st.title("📝 Analisis Detail per Pertanyaan")
    st.markdown("---")
    
    st.info("Grafik di bawah memperlihatkan sebaran jawaban untuk setiap pertanyaan secara detail.")

    # Data Processing untuk Stacked Bar
    q_counts = df_display.apply(pd.Series.value_counts).reindex(URUTAN_SKALA).fillna(0).T.reset_index()
    q_counts = q_counts.rename(columns={'index': 'Pertanyaan'})
    q_melt = q_counts.melt(id_vars='Pertanyaan', var_name='Skala', value_name='Jumlah')

    fig = px.bar(
        q_melt,
        x="Pertanyaan",
        y="Jumlah",
        color="Skala",
        barmode="stack",
        category_orders={"Skala": URUTAN_SKALA},
        color_discrete_map=COLOR_MAP,
        text_auto=True,
        title="Distribusi Jawaban per Pertanyaan (Stacked)"
    )
    
    fig.update_layout(
        height=600,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    st.plotly_chart(fig, use_container_width=True)
    
    # Heatmap Distribusi
    st.subheader("Heatmap Kepadatan Jawaban")
    heatmap_data = q_counts.set_index('Pertanyaan')
    
    fig_heat = px.imshow(
        heatmap_data,
        labels=dict(x="Skala", y="Pertanyaan", color="Jumlah"),
        x=URUTAN_SKALA,
        y=heatmap_data.index,
        color_continuous_scale="Blues",
        aspect="auto",
        text_auto=True
    )
    fig_heat.update_layout(height=500)
    st.plotly_chart(fig_heat, use_container_width=True)

# =====================================================
# PERINGKAT SKOR
# =====================================================
elif menu == "⭐ Peringkat Skor":
    st.title("⭐ Peringkat Skor Rata-rata")
    st.markdown("---")
    
    # Hitung rata-rata
    avg_scores = df_num_display.mean().reset_index()
    avg_scores.columns = ["Pertanyaan", "Rata_rata_Skor"]
    avg_scores = avg_scores.sort_values("Rata_rata_Skor", ascending=True)
    
    tab1, tab2 = st.tabs(["📊 Bar Chart", "📉 Analisis Detail"])
    
    with tab1:
        fig = px.bar(
            avg_scores,
            x="Rata_rata_Skor",
            y="Pertanyaan",
            orientation="h",
            color="Rata_rata_Skor",
            color_continuous_scale="Viridis",
            text=avg_scores["Rata_rata_Skor"].apply(format_skor),
            range_x=[0, 7]
        )
        
        fig.update_layout(
            title="Peringkat Rata-rata Skor per Pertanyaan",
            xaxis_title="Skor Rata-rata (Max 6.0)",
            yaxis_title="Pertanyaan",
            height=600
        )
        st.plotly_chart(fig, use_container_width=True)
        
    with tab2:
        st.subheader("Detail Statistik Skor")
        stats_df = df_num_display.describe().T[["mean", "std", "min", "max", "25%", "50%", "75%"]]
        stats_df = stats_df.sort_values("mean", ascending=False)
        st.dataframe(stats_df.style.background_gradient(cmap="Greens", subset=["mean"]), use_container_width=True)

# =====================================================
# ANALISIS KATEGORI / SENTIMEN
# =====================================================
elif menu == "🎭 Analisis Sentimen/Kategori":
    st.title("🎭 Analisis Kategori (Positif/Netral/Negatif)")
    st.markdown("---")
    
    st.markdown("""
    **Kategorisasi:**
    - ✅ **Positif**: Sangat Setuju (SS) & Setuju (S)
    - ➖ **Netral**: Cukup Setuju (CS)
    - ❌ **Negatif**: Cukup Tidak Setuju (CTS), Tidak Setuju (TS), Sangat Tidak Setuju (STS)
    """)
    
    # Hitung per Pertanyaan
    def hitung_kategori(row):
        pos = (row == 'SS').sum() + (row == 'S').sum()
        neu = (row == 'CS').sum()
        neg = (row == 'CTS').sum() + (row == 'TS').sum() + (row == 'STS').sum()
        return pd.Series([pos, neu, neg], index=['Positif', 'Netral', 'Negatif'])

    cat_per_q = df_display.apply(hitung_kategori, axis=0).T.reset_index().rename(columns={'index': 'Pertanyaan'})
    
    # Chart 1: Stacked Bar 100%
    st.subheader("Komposisi Sentimen per Pertanyaan")
    
    # Normalize to percentage
    cat_per_q['Total'] = cat_per_q['Positif'] + cat_per_q['Netral'] + cat_per_q['Negatif']
    # Hindari pembagian dengan nol
    cat_per_q['Total'] = cat_per_q['Total'].replace(0, 1)
    
    cat_per_q['Pct_Positif'] = (cat_per_q['Positif'] / cat_per_q['Total']) * 100
    cat_per_q['Pct_Netral'] = (cat_per_q['Netral'] / cat_per_q['Total']) * 100
    cat_per_q['Pct_Negatif'] = (cat_per_q['Negatif'] / cat_per_q['Total']) * 100
    
    cat_melt = cat_per_q.melt(
        id_vars=['Pertanyaan', 'Total', 'Positif', 'Netral', 'Negatif'], 
        value_vars=['Pct_Positif', 'Pct_Netral', 'Pct_Negatif'],
        var_name='Kategori_Pct', value_name='Persentase'
    )
    
    # Mapping nama kategori agar bersih di legenda
    cat_map_name = {'Pct_Positif': 'Positif', 'Pct_Netral': 'Netral', 'Pct_Negatif': 'Negatif'}
    cat_melt['Kategori'] = cat_melt['Kategori_Pct'].map(cat_map_name)
    
    color_sentimen = {'Positif': '#28B463', 'Netral': '#808B96', 'Negatif': '#C0392B'}

    fig = px.bar(
        cat_melt,
        x="Persentase",
        y="Pertanyaan",
        color="Kategori",
        orientation='h',
        color_discrete_map=color_sentimen,
        text=cat_melt['Persentase'].apply(format_persen)
    )
    
    fig.update_layout(
        title="Persentase Sentimen per Pertanyaan",
        xaxis_title="Persentase (%)",
        height=600,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    st.plotly_chart(fig, use_container_width=True)
    
    # Chart 2: Total Agregat
    st.subheader("Total Agregat Sentimen Keseluruhan")
    total_pos = cat_per_q['Positif'].sum()
    total_neu = cat_per_q['Netral'].sum()
    total_neg = cat_per_q['Negatif'].sum()
    
    df_agregat = pd.DataFrame({
        'Kategori': ['Positif', 'Netral', 'Negatif'],
        'Jumlah': [total_pos, total_neu, total_neg]
    })
    
    fig_agg = px.bar(
        df_agregat,
        x="Kategori",
        y="Jumlah",
        color="Kategori",
        color_discrete_map=color_sentimen,
        text_auto=True
    )
    st.plotly_chart(fig_agg, use_container_width=True)


# =====================================================
# TABEL DATA LENGKAP
# =====================================================
elif menu == "📋 Tabel Data Lengkap":
    st.title("📋 Data Mentah Kuesioner")
    st.markdown("---")
    
    st.write(f"**Total Data:** {len(df)} Responden x {len(df.columns)} Pertanyaan")
    
    # Tampilkan dataframe
    st.dataframe(df_display, use_container_width=True, height=500)
    
    # Download Button
    csv = df_display.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Data CSV",
        data=csv,
        file_name="hasil_kuesioner_processed.csv",
        mime="text/csv"
    )

# =====================================================
# FOOTER
# =====================================================
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray;'>
    <p>📊 Dashboard Analisis Kuesioner | © 2026</p>
    <p>Dikembangkan dengan Streamlit & Plotly</p>
    </div>
    """,
    unsafe_allow_html=True
)

# =====================================================
# STYLE TAMBAHAN
# =====================================================
st.markdown("""
<style>
    .stMetric {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #1f77b4;
    }
    
    .stMetric label {
        font-weight: bold;
        color: #2c3e50;
    }
    
    .stMetric div {
        font-size: 1.5rem;
        color: #1f77b4;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #f0f2f6;
        border-radius: 5px 5px 0px 0px;
        padding: 8px 12px;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #1f77b4;
        color: white;
    }
</style>
""", unsafe_allow_html=True)