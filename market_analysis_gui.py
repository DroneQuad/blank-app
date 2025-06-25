# market_analysis_gui.py
# GUI Web App menggunakan Streamlit

import streamlit as st
import pandas as pd
from datetime import datetime
import io

from Analysisxls_combined import (
    fetch_market_data,
    create_excel_report,
    generate_short_report,
    generate_chart_with_ema,
    send_telegram_message,
    send_chart_to_telegram,
    format_price,
    generate_comment
)

# ================================
# Layout Streamlit
# ================================
st.set_page_config(page_title="Market Analysis Dashboard", layout="wide")
st.title("📈 Market Analysis Dashboard")

# Ambil data pasar
with st.spinner("Fetching market data..."):
    assets_data, macro_drivers = fetch_market_data()
    df = pd.DataFrame(assets_data)
    df['Komentar'] = df.apply(lambda row: generate_comment(row['symbol'], row['bias']), axis=1)

# ================================
# Tampilkan Ringkasan
# ================================
st.subheader("🔹 Ringkasan Analisa")
st.dataframe(df[['symbol', 'price', 'bias', 'Komentar']], use_container_width=True)

# ================================
# Pilih Asset untuk Chart
# ================================
symbol_map = {
    "BTCUSD": "BTC-USD",
    "ETHUSD": "ETH-USD",
    "XAUUSD": "GC=F",
    "USOIL": "CL=F",
    "USTEC": "^NDX"
}

selected = st.selectbox("📊 Pilih asset untuk lihat chart:", list(symbol_map.keys()))
ticker = symbol_map[selected]

chart_buf = generate_chart_with_ema(ticker, selected)
if chart_buf:
    st.image(chart_buf, caption=f"{selected} 15m Chart with EMA5/20")
else:
    st.warning("Chart tidak tersedia.")

# ================================
# Download Excel
# ================================
if st.button("📥 Download Laporan Excel"):
    excel_file = create_excel_report(assets_data, macro_drivers)
    with open(excel_file, "rb") as f:
        st.download_button(
            label="Klik untuk unduh",
            data=f,
            file_name=excel_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ================================
# Kirim ke Telegram Opsional
# ================================
if st.button("📤 Kirim Laporan ke Telegram"):
    message = generate_short_report(assets_data)
    send_telegram_message(message)
    st.success("Pesan terkirim ke Telegram!")
    for sym, tick in symbol_map.items():
        send_chart_to_telegram(tick, sym)
    st.success("Chart juga dikirim ke Telegram!")

# ================================
# Footer
# ================================
st.markdown("---")
st.markdown("🛡️ _Data diperbarui secara real-time. Gunakan manajemen risiko yang bijak_.")
