import streamlit as st
import pandas as pd
import plotly.express as px
import io
import os

st.set_page_config(page_title="Data Studio", page_icon="📊", layout="wide")

st.title("📊 Data Studio")
st.caption("Upload Excel files to visualize · Build and export custom datasets")

tab1, tab2 = st.tabs(["📈 Visualizer", "🗄️ Data Storage"])


# ─────────────────────────────────────────────
# TAB 1 — VISUALIZER
# ─────────────────────────────────────────────
with tab1:
    uploaded = st.file_uploader("Upload Excel or CSV file", type=["xlsx", "xls", "csv"])

    if uploaded:
        # Load file
        if uploaded.name.endswith(".csv"):
            df = pd.read_csv(uploaded)
            sheet_name = None
        else:
            xl = pd.ExcelFile(uploaded)
            sheet_names = xl.sheet_names
            sheet_name = st.selectbox("Select sheet", sheet_names) if len(sheet_names) > 1 else sheet_names[0]
            df = xl.parse(sheet_name)

        st.success(f"Loaded **{len(df)} rows × {len(df.columns)} columns**")

        cols = list(df.columns)

        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            chart_type = st.selectbox("Chart type", ["Bar", "Line", "Pie", "Scatter", "Area", "Histogram", "Box"])
        with col2:
            x_col = st.selectbox("X axis / Label column", cols)
        with col3:
            remaining = [c for c in cols if c != x_col]
            if chart_type in ("Histogram", "Box"):
                y_cols = st.multiselect("Y axis columns", remaining, default=remaining[:2])
            elif chart_type == "Pie":
                y_cols = [st.selectbox("Value column", remaining)]
            else:
                y_cols = st.multiselect("Y axis columns", remaining, default=remaining[:2])

        chart_title = st.text_input("Chart title", value="My Chart")

        if y_cols or chart_type in ("Histogram",):
            fig = None
            try:
                if chart_type == "Bar":
                    fig = px.bar(df, x=x_col, y=y_cols, title=chart_title, barmode="group")
                elif chart_type == "Line":
                    fig = px.line(df, x=x_col, y=y_cols, title=chart_title, markers=True)
                elif chart_type == "Area":
                    fig = px.area(df, x=x_col, y=y_cols, title=chart_title)
                elif chart_type == "Pie":
                    fig = px.pie(df, names=x_col, values=y_cols[0], title=chart_title)
                elif chart_type == "Scatter":
                    fig = px.scatter(df, x=x_col, y=y_cols[0], title=chart_title,
                                     color=y_cols[1] if len(y_cols) > 1 else None)
                elif chart_type == "Histogram":
                    fig = px.histogram(df, x=x_col, title=chart_title)
                elif chart_type == "Box":
                    fig = px.box(df, x=x_col, y=y_cols[0] if y_cols else None, title=chart_title)
            except Exception as e:
                st.error(f"Could not render chart: {e}")

            if fig:
                st.plotly_chart(fig, use_container_width=True)

                # Download chart as HTML (interactive)
                html_buf = io.StringIO()
                fig.write_html(html_buf)
                st.download_button(
                    label="⬇ Download chart (HTML)",
                    data=html_buf.getvalue(),
                    file_name=f"{chart_title.replace(' ', '_')}.html",
                    mime="text/html"
                )

                # Download chart as PNG (static)
                try:
                    img_bytes = fig.to_image(format="png", scale=2)
                    st.download_button(
                        label="⬇ Download chart (PNG)",
                        data=img_bytes,
                        file_name=f"{chart_title.replace(' ', '_')}.png",
                        mime="image/png"
                    )
                except Exception:
                    st.info("Install `kaleido` for PNG export: `pip install kaleido`")
        else:
            st.info("Select at least one Y axis column to render the chart.")

        st.subheader("Data preview")
        st.dataframe(df.head(20), use_container_width=True)

        # Download original data
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        st.download_button(
            "⬇ Download data as XLSX",
            data=out.getvalue(),
            file_name="data_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.info("Upload an Excel (.xlsx / .xls) or CSV file to get started.")


# ─────────────────────────────────────────────
# TAB 2 — DATA STORAGE
# ─────────────────────────────────────────────
with tab2:
    st.subheader("Manual data entry table")

    # Column management
    if "storage_cols" not in st.session_state:
        st.session_state.storage_cols = ["Name", "Value", "Notes"]
    if "storage_data" not in st.session_state:
        st.session_state.storage_data = pd.DataFrame(columns=st.session_state.storage_cols)

    with st.expander("⚙ Manage columns"):
        new_col = st.text_input("New column name", key="new_col_input")
        col_a, col_b = st.columns([1, 3])
        with col_a:
            if st.button("+ Add column") and new_col.strip():
                if new_col.strip() not in st.session_state.storage_cols:
                    st.session_state.storage_cols.append(new_col.strip())
                    st.session_state.storage_data[new_col.strip()] = ""
                    st.rerun()
        with col_b:
            if len(st.session_state.storage_cols) > 1:
                col_to_remove = st.selectbox("Remove column", st.session_state.storage_cols, key="rm_col")
                if st.button("Remove selected column"):
                    st.session_state.storage_cols.remove(col_to_remove)
                    st.session_state.storage_data.drop(columns=[col_to_remove], inplace=True)
                    st.rerun()

    # Editable data table
    edited_df = st.data_editor(
        st.session_state.storage_data,
        num_rows="dynamic",
        use_container_width=True,
        key="storage_editor"
    )
    st.session_state.storage_data = edited_df

    st.caption(f"{len(edited_df)} rows · {len(edited_df.columns)} columns · Click any cell to edit · Use ＋ button at bottom to add rows")

    # Downloads
    dl1, dl2 = st.columns(2)
    with dl1:
        out_xlsx = io.BytesIO()
        with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
            edited_df.to_excel(writer, index=False)
        st.download_button(
            "⬇ Download as XLSX",
            data=out_xlsx.getvalue(),
            file_name="data_storage.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with dl2:
        csv_data = edited_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇ Download as CSV",
            data=csv_data,
            file_name="data_storage.csv",
            mime="text/csv"
        )
