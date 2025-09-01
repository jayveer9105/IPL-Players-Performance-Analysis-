import streamlit as st
import os
from main import (
    load_ipl_data,
    clean_ipl_data,
    analyze_ipl,
    export_report,
    generate_charts,
    embed_charts
)

# ---------------- Search Player Function ---------------- #
def search_player(df, player_name):
    if player_name.strip() == "":
        return None
    result = df[df['Player'].str.strip().str.lower() == player_name.strip().lower()]
    if not result.empty:
        return result.iloc[0]
    return None

# ---------------- Main Streamlit App ---------------- #
def main():
    st.set_page_config(page_title="IPL Player Analyzer", layout="wide")
    st.title("🏏 IPL Player Stats Analysis System")
    st.markdown("---")

    # Sidebar Navigation
    st.sidebar.title("Navigation")
    page = st.sidebar.selectbox("Choose a page", ["Search Player", "Data Cleaning", "Generate Analysis Report", "View Charts"])

    # Load RAW IPL data
    if 'df_raw' not in st.session_state:
        with st.spinner("Loading IPL data..."):
            try:
                df_raw = load_ipl_data("ipl_data.xlsx")   # Update your file path
                st.session_state.df_raw = df_raw
                st.success("Raw IPL data loaded successfully!")
            except Exception as e:
                st.error(f"Error loading IPL data: {str(e)}")
                return

    df_raw = st.session_state.df_raw

    # ---------------- Search Player ---------------- #
    if page == "Search Player":
        st.header("🔍 Search Player Details")
        player_name = st.text_input("Enter Player Name:", placeholder="e.g., Virat Kohli")
        if st.button("Search"):
            # Use cleaned data if available, else raw
            df_to_search = st.session_state.get('df_ipl', df_raw)
            result = search_player(df_to_search, player_name)
            if result is not None:
                st.success("Player found!")

                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("Player Info")
                    st.write(f"**Name:** {result['Player']}")
                    st.write(f"**Role:** {result['Role']}")
                    st.write(f"**Team:** {result['Team']}")
                    st.write(f"**Matches:** {result['Matches']}")
                with col2:
                    st.subheader("Performance Stats")
                    st.write(f"**Runs:** {result['Runs']}")
                    st.write(f"**Batting SR:** {result['Bat_SR']}")
                    if result['Wickets'] > 0:
                        st.write(f"**Wickets:** {result['Wickets']}")
                        st.write(f"**Economy:** {result['Econ']}")

                with st.expander("📋 Show Full Record"):
                    st.dataframe(result.to_frame().T)
            else:
                st.error("Player not found. Please check the name.")

    # ---------------- Data Cleaning ---------------- #
    elif page == "Data Cleaning":
        st.header("🧹 Data Cleaning Overview")
        st.markdown("Inspect and clean missing values & duplicates in the IPL dataset.")

        st.subheader("Missing Values (Before Cleaning)")
        st.dataframe(df_raw.isna().sum())

        st.subheader("Duplicate Rows (Before Cleaning)")
        duplicates = df_raw[df_raw.duplicated()]
        if not duplicates.empty:
            st.dataframe(duplicates)
        else:
            st.write("✅ No duplicate rows found.")

        if st.button("Clean Data Now"):
            cleaned_df = clean_ipl_data(df_raw.copy())
            st.session_state.df_ipl = cleaned_df.copy()
            st.success("Data cleaned successfully!")
            st.subheader("Cleaned Data Preview")
            st.dataframe(cleaned_df.head(70))

            st.info("Cleaning Steps Applied:\n"
                    "- Missing Player/Role/Team filled with 'Unknown'\n"
                    "- Missing numeric stats filled with 0\n"
                    "- Duplicate rows removed")

    # ---------------- Generate Analysis Report ---------------- #
    elif page == "Generate Analysis Report":
        st.header("📊 Generate Analysis Report")

        if 'df_ipl' not in st.session_state:
            st.warning("⚠️ Please clean the data first (use Data Cleaning page).")
            return

        df = st.session_state.df_ipl

        if st.button("Generate Complete Analysis Report"):
            with st.spinner("Generating report..."):
                try:
                    top_runs, top_wickets, summary = analyze_ipl(df)

                    output_file = "ipl_analysis_report.xlsx"
                    export_report(output_file, top_runs, top_wickets, summary)
                    chart_paths = generate_charts(df)
                    embed_charts(output_file, chart_paths)

                    st.success("Analysis report generated successfully!")

                    st.subheader("Summary")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Players", len(df))
                        st.metric("Total Teams", df['Team'].nunique())
                    with col2:
                        st.metric("Top Run Scorer", top_runs.iloc[0]['Player'])
                        st.metric("Runs", int(top_runs.iloc[0]['Runs']))
                    with col3:
                        st.metric("Top Wicket Taker", top_wickets.iloc[0]['Player'])
                        st.metric("Wickets", int(top_wickets.iloc[0]['Wickets']))

                    if os.path.exists(output_file):
                        with open(output_file, "rb") as f:
                            st.download_button(
                                label="📥 Download Analysis Report (Excel)",
                                data=f.read(),
                                file_name=output_file,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                except Exception as e:
                    st.error(f"Error generating report: {str(e)}")

    # ---------------- View Charts ---------------- #
    elif page == "View Charts":
        st.header("📈 IPL Data Visualization Charts")

        if 'df_ipl' not in st.session_state:
            st.warning("⚠️ Please clean the data first (use Data Cleaning page).")
            return

        df = st.session_state.df_ipl
        charts_dir = "charts"
        chart_files = {
            "Runs Distribution": "runs_dist.png",
            "Wickets Distribution": "wickets_dist.png",
            "Total Runs by Team": "team_runs.png",
            "Total Wickets by Team": "team_wickets.png"
        }

        if st.button("Generate Charts"):
            with st.spinner("Generating charts..."):
                try:
                    generate_charts(df, charts_dir)  # overwrite existing charts
                    st.success("Charts generated successfully!")
                except Exception as e:
                    st.error(f"Error generating charts: {str(e)}")

        for chart_name, chart_file in chart_files.items():
            chart_path = os.path.join(charts_dir, chart_file)
            if os.path.exists(chart_path):
                st.subheader(chart_name)
                st.image(chart_path, use_column_width=True)
            else:
                st.info(f"{chart_name} not found. Click 'Generate Charts' to create it.")

if __name__ == "__main__":
    main()
