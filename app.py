import streamlit as st
import pandas as pd
from datetime import timedelta, date
import plotly.express as px
import io
import streamlit.components.v1 as components
from dotenv import load_dotenv
import os
import plotly.io as pio

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Case Analysis Dashboard",
    page_icon="📅",
    # layout="wide",
)

os.environ["KALEIDO_CHROME_PATH"] = "/usr/bin/chromium"
# pio.kaleido.scope.chromium_executable = None
pio.renderers.default = "png" 
# Example figure
fig = px.bar(x=["A", "B", "C"], y=[1, 3, 2])

# Convert to image
img_bytes = fig.to_image(format="png")
load_dotenv()

def get_status_list(env_key):
    value = os.getenv(env_key, "")
    return [status.strip() for status in value.split(",") if status.strip()]

OPEN_STATUSES = get_status_list("OPEN_STATUSES")
CLOSED_STATUSES = get_status_list("CLOSED_STATUSES")
OPEN_STATUSESAVG = get_status_list("OPEN_STATUSES_AVG")
selected_owners = get_status_list("SELECTED_OWNERS")

def add_pdf_export():
    """
    Adds CSS for a perfectly structured, print-friendly PDF export and a button to trigger it.
    """
    st.sidebar.markdown("---")
    
    print_css = """
<style>
    @media print {
        /* --- General Layout --- */
        section[data-testid="stSidebar"] {
            display: none !important;
        }

        /* --- Page Break Rules --- */
        h1, h2, h3 {
            break-before: page !important;
            padding-top: 2rem !important;
        }

        /* --- Element Splitting Prevention --- */
        div[data-testid="stPlotlyChart"],
        div[data-testid="stMetric"],
        div[data-testid="stHorizontalBlock"] {
            break-inside: avoid !important;
        }
        
        /* --- Hide Unnecessary Elements --- */
        div[data-testid="stDownloadButton"] { display: none !important; }
        iframe { display: none !important; }
        
        /* --- Professional Table Styling for PDF --- */
        div[data-testid="stDataFrame"] {
            break-inside: avoid !important;
        }

        div[data-testid="stDataFrame"] > div > table {
            width: 100% !important;
            border-collapse: collapse !important;
            border: 1px solid #a8a8a8 !important;
        }
        
        div[data-testid="stDataFrame"] > div > table th {
            background-color: #f0f0f0 !important;
            color: black !important;
            border: 1px solid #a8a8a8 !important;
            padding: 8px 12px !important;
            text-align: left !important;
        }

        div[data-testid="stDataFrame"] > div > table td {
            border: 1px solid #dcdcdc !important;
            padding: 8px 12px !important;
            
            /* --- THIS IS THE FIX --- */
            /* Force the text color to be black so it's visible on a white background */
            color: black !important; 
        }
    }
</style>
    """
    st.markdown(print_css, unsafe_allow_html=True)
    
    # Read the content of the HTML file for the button
    try:
        with open("print_button.html", "r") as f:
            html_content = f.read()
    except FileNotFoundError:
        st.sidebar.error("Error: 'print_button.html' not found.")
        return

    # Embed the custom button from the HTML file
    with st.sidebar:
        components.html(html_content, height=50)


# --- HELPER FUNCTION ---
def create_download_buttons(fig, df, file_name, index=False):
    """Creates Streamlit download buttons for a chart (PNG) and its data (Excel)."""
    # Convert plot to image bytes
    img_bytes = fig.to_image(format="png")

    # Prepare Excel data bytes
    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        df.to_excel(writer, index=index, sheet_name='ChartData')
    excel_data = output_excel.getvalue()

    # Create two columns for the buttons
    btn_col1, btn_col2 = st.columns(2)

    with btn_col1:
        st.download_button(
            label="📥 Download Chart",
            data=img_bytes,
            file_name=f"{file_name}_chart.png",
            mime="image/png",
            key=f"download_chart_{file_name}"
        )

    with btn_col2:
        st.download_button(
            label="📥 Download Data",
            data=excel_data,
            file_name=f"{file_name}_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"download_data_{file_name}"
        )


# --- APP TITLE ---
st.title("📅 Case Analysis Dashboard")

# --- FILE UPLOADER ---
uploaded_file = st.file_uploader("Upload your Excel file to begin", type=['xlsx'])

if uploaded_file:
    # Read and prepare the data
    df = pd.read_excel(uploaded_file)
    try:
        df['Opened Date'] = pd.to_datetime(df['Opened Date'], dayfirst=True)
        if 'Case Last Modified Date' in df.columns:
            df['Case Last Modified Date'] = pd.to_datetime(df['Case Last Modified Date'], dayfirst=True)
    except KeyError:
        st.error("Error: The uploaded file must contain an 'Opened Date' column.")
        st.stop()
    except Exception as e:
        st.error(f"Error processing Excel file: {e}")
        st.stop()

    # --- SIDEBAR FOR TIME FRAME SELECTION ---
    st.sidebar.header("Select Time Frame")
    selection_mode = st.sidebar.radio(
        "How do you want to select the time frame?",
        ('By Date Range', 'By Week')
    )
    add_pdf_export()

    start_date = None
    end_date = None

    if selection_mode == 'By Date Range':
        min_available_date = df['Opened Date'].min().date()
        max_available_date = df['Opened Date'].max().date()
        start_date = st.sidebar.date_input("Start Date", min_available_date, min_value=min_available_date, max_value=max_available_date)
        end_date = st.sidebar.date_input("End Date", max_available_date, min_value=min_available_date, max_value=max_available_date)
    
    else: # By Week
        today_date_obj = date.today()
        min_available_date = df['Opened Date'].min().date()
        max_available_date = df['Opened Date'].max().date()
        default_week_day = today_date_obj
        if not (min_available_date <= today_date_obj <= max_available_date):
            default_week_day = min_available_date
        selected_day = st.sidebar.date_input("Select any day within the desired week", default_week_day, min_value=min_available_date, max_value=max_available_date)
        start_date = selected_day - timedelta(days=selected_day.weekday())
        end_date = start_date + timedelta(days=6)

    # --- MAIN CALCULATION AND DISPLAY ---
    if start_date and end_date:
        start_date_dt = pd.to_datetime(start_date)
        end_date_dt = pd.to_datetime(end_date)
        
        # --- Reports based on OPENED DATE ---
        st.header(f"Report for Cases Opened/Closed Between: {start_date_dt.strftime('%d %b, %Y')} and {end_date_dt.strftime('%d %b, %Y')}")

        # --- DYNAMIC SUMMARY BOXES ---
        opened_summary_data = []
        closed_summary_data = []

        # Iterate through each week in the selected date range
        for week_start in pd.date_range(start=start_date, end=end_date, freq='W-MON'):
            week_end = week_start + timedelta(days=6)
            
            opened_mask = (df['Opened Date'] >= week_start) & (df['Opened Date'] <= week_end ) & (df['Status'].isin(OPEN_STATUSES))
            opened_count = len(df[opened_mask])
            
            closed_count = 0
            if 'Case Last Modified Date' in df.columns:
                closed_mask = (df['Case Last Modified Date'] >= week_start) & (df['Case Last Modified Date'] <= week_end) & (df['Status'].isin(CLOSED_STATUSES))
                closed_count = len(df[closed_mask])

            week_num = week_start.isocalendar().week
            year = week_start.isocalendar().year
            date_range_str = f"{week_start.strftime('%m/%d/%Y')} – {week_end.strftime('%m/%d/%Y')}"
            week_str = f"Week {week_num} FY {year} ({date_range_str})"
            
            opened_summary_data.append({'Week': week_str, 'Cases Opened': opened_count})
            closed_summary_data.append({'Week': week_str, 'Cases Closed': closed_count})

        opened_summary_df = pd.DataFrame(opened_summary_data)
        closed_summary_df = pd.DataFrame(closed_summary_data)
        
        box_col1, box_col2 = st.columns(2)
        with box_col1:
            st.markdown("##### Weekly Overview: Cases Opened")
            st.dataframe(opened_summary_df)
            output_opened_summary = io.BytesIO()
            with pd.ExcelWriter(output_opened_summary, engine='openpyxl') as writer:
                opened_summary_df.to_excel(writer, index=False, sheet_name='Opened_Summary')
            st.download_button(
                label="📥 Download Opened Summary",
                data=output_opened_summary.getvalue(),
                file_name='weekly_opened_summary.xlsx',
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with box_col2:
            st.markdown("##### Weekly Overview: Cases Closed")
            st.dataframe(closed_summary_df)
            output_closed_summary = io.BytesIO()
            with pd.ExcelWriter(output_closed_summary, engine='openpyxl') as writer:
                closed_summary_df.to_excel(writer, index=False, sheet_name='Closed_Summary')
            st.download_button(
                label="📥 Download Closed Summary",
                data=output_closed_summary.getvalue(),
                file_name='weekly_closed_summary.xlsx',
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        st.markdown("---")

        cases_in_range = df[(df['Opened Date'] >= start_date_dt) & (df['Opened Date'] <= end_date_dt) & (df['Status'].isin(OPEN_STATUSES)) &
                            (df['Case Owner'].isin(selected_owners))].copy()
        
        open_cases_data = cases_in_range[cases_in_range['Status'].isin(OPEN_STATUSES)].copy()
        closed_cases_data = cases_in_range[cases_in_range['Status'].isin(CLOSED_STATUSES)].copy()
        
        closed_in_period_df = pd.DataFrame()
        if 'Case Last Modified Date' in df.columns:
             closed_in_period_df = df[
                 
                (df['Case Last Modified Date'] >= start_date_dt) & 
                (df['Case Last Modified Date'] <= end_date_dt) &
                (df['Status'].isin(CLOSED_STATUSES))
            ].copy()

        st.subheader("Metrics for Cases Opened in Period")
        metric_col1, metric_col2, metric_col3 = st.columns(3)
        with metric_col1:
            st.metric(label="Total Open Cases", value=len(open_cases_data))
        with metric_col2:
            st.metric(label="Of Those, Now Closed", value=len(closed_cases_data))
        with metric_col3:
            st.metric(label="Total Cases Closed in Period", value=len(closed_in_period_df))
        with st.expander("View Detailed Report for Cases Opened in Period"):
            st.dataframe(cases_in_range)
            
            # Make a copy and ensure datetime columns are correct
            df_to_export = cases_in_range.copy()
            
            # Convert 'Opened Date' and 'Case Last Modified Date' to proper datetime if they exist
            if 'Opened Date' in df_to_export.columns:
                df_to_export['Opened Date'] = pd.to_datetime(df_to_export['Opened Date'])
            if 'Case Last Modified Date' in df_to_export.columns:
                df_to_export['Case Last Modified Date'] = pd.to_datetime(df_to_export['Case Last Modified Date'])
            
            # Create Excel in memory
            output_open = io.BytesIO()
            with pd.ExcelWriter(output_open, engine='openpyxl', datetime_format='yyyy-mm-dd') as writer:
                df_to_export.to_excel(writer, index=False, sheet_name='Open_In_Period')
            excel_data_open = output_open.getvalue()

            st.download_button(
                label="📥 Download This Detailed Report",
                data=excel_data_open,
                file_name="open_in_period_detailed_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        
        st.markdown("---")

        st.subheader("Breakdown of Cases Opened in Period")

        st.markdown("##### All Cases Opened")
        report_col1, chart_col1 = st.columns(2)
        with report_col1:
            if 'Case Reason' in df.columns:
                st.markdown("###### By Reason")
                if not cases_in_range.empty:
                    all_summary_df = cases_in_range.groupby(['Product Line', 'Case Reason']).size().reset_index(name='Record Count')
                    total_all_count = all_summary_df['Record Count'].sum()
                    total_row_all = pd.DataFrame([{'Product Line': '**Grand Total**', 'Case Reason': '', 'Record Count': total_all_count}])
                    all_summary_df = pd.concat([all_summary_df, total_row_all], ignore_index=True)
                    st.dataframe(all_summary_df)
            else:
                st.warning("Missing 'Case Reason' column.")
        with chart_col1:
            st.markdown("###### By Product Line")
            if not cases_in_range.empty:
                all_product_summary = cases_in_range.groupby('Product Line').size().reset_index(name='Record Count')
                fig_all = px.pie(
                    all_product_summary, 
                    values='Record Count', 
                    names='Product Line', 
                    title='All Cases by Product Line',
                    color_discrete_sequence=px.colors.qualitative.Vivid
                )
                fig_all.update_traces(textinfo='percent+value')
                st.plotly_chart(fig_all, use_container_width=True)
                create_download_buttons(fig_all, all_product_summary, "all_cases_by_product_line")
        
        st.markdown("---")

        st.markdown("##### Of Those Opened, Which Are Now Closed")
        report_col2, chart_col2 = st.columns(2)
        with report_col2:
            if 'Case Reason' in df.columns:
                st.markdown("###### By Reason")
                if not closed_cases_data.empty:
                    closed_summary_df = closed_cases_data.groupby(['Product Line', 'Case Reason']).size().reset_index(name='Record Count')
                    total_closed_count = closed_summary_df['Record Count'].sum()
                    total_row_closed = pd.DataFrame([{'Product Line': '**Grand Total**', 'Case Reason': '', 'Record Count': total_closed_count}])
                    closed_summary_df = pd.concat([closed_summary_df, total_row_closed], ignore_index=True)
                    st.dataframe(closed_summary_df)
                else:
                    st.info("No cases opened in this period are closed.")
            else:
                st.warning("Missing 'Case Reason' column.")

        with chart_col2:
            st.markdown("###### By Product Line")
            if not closed_cases_data.empty:
                closed_product_summary = closed_cases_data.groupby('Product Line').size().reset_index(name='Record Count')
                fig_closed = px.pie(
                    closed_product_summary, 
                    values='Record Count', 
                    names='Product Line', 
                    title='Closed Cases by Product Line',
                    color_discrete_sequence=px.colors.qualitative.Vivid
                )
                fig_closed.update_traces(textinfo='percent+value')
                st.plotly_chart(fig_closed, use_container_width=True)
                create_download_buttons(fig_closed, closed_product_summary, "closed_cases_by_product_line")
        
        st.markdown("---")
        
        st.header("Breakdown of Cases Closed in Selected Period")
        if 'Case Last Modified Date' in df.columns and not closed_in_period_df.empty:
            report_col3, chart_col3 = st.columns(2)
            with report_col3:
                if 'Case Reason' in df.columns:
                    st.markdown("###### By Reason")
                    closed_period_summary_df = closed_in_period_df.groupby(['Product Line', 'Case Reason']).size().reset_index(name='Record Count')
                    total_closed_period_count = closed_period_summary_df['Record Count'].sum()
                    total_row_closed_period = pd.DataFrame([{'Product Line': '**Grand Total**', 'Case Reason': '', 'Record Count': total_closed_period_count}])
                    closed_period_summary_df = pd.concat([closed_period_summary_df, total_row_closed_period], ignore_index=True)
                    st.dataframe(closed_period_summary_df)
                else:
                    st.warning("Missing 'Case Reason' column.")

            with chart_col3:
                st.markdown("###### By Product Line")
                closed_in_period_summary = closed_in_period_df.groupby('Product Line').size().reset_index(name='Record Count')
                fig_closed_period = px.pie(
                    closed_in_period_summary, 
                    values='Record Count', 
                    names='Product Line', 
                    title='Cases Closed in Period by Product Line',
                    color_discrete_sequence=px.colors.qualitative.Vivid
                )
                fig_closed_period.update_traces(textinfo='percent+value')
                st.plotly_chart(fig_closed_period, use_container_width=True)
                create_download_buttons(fig_closed_period, closed_in_period_summary, "closed_in_period_by_product_line")

            with st.expander("View Detailed Report for Cases Closed in Period"):
                st.dataframe(closed_in_period_df)

                # Copy and ensure datetime columns are proper
                df_closed_export = closed_in_period_df.copy()
                if 'Opened Date' in df_closed_export.columns:
                    df_closed_export['Opened Date'] = pd.to_datetime(df_closed_export['Opened Date'])
                if 'Case Last Modified Date' in df_closed_export.columns:
                    df_closed_export['Case Last Modified Date'] = pd.to_datetime(df_closed_export['Case Last Modified Date'])
                
                # Excel export
                output_closed = io.BytesIO()
                with pd.ExcelWriter(output_closed, engine='openpyxl', datetime_format='yyyy-mm-dd') as writer:
                    df_closed_export.to_excel(writer, index=False, sheet_name='Closed_In_Period')
                excel_data_closed = output_closed.getvalue()

                st.download_button(
                    label="📥 Download Closed Cases Detailed Report",
                    data=excel_data_closed,
                    file_name="closed_in_period_detailed_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("No cases closed in this period or 'Case Last Modified Date' column is missing.")
        # st.markdown("---")
        st.markdown("---")

        # --- MODIFIED SECTION STARTS HERE ---
        st.subheader("Additional Analysis (YTD Open Cases for Key Product Lines)")
        st.info("This section provides a separate breakdown of currently open cases (from YTD) for each of the key product lines: Barcode, RFID, PRI, and Reach.")

        # Define YTD date range and the key product lines to analyze
        today = date.today()
        start_of_year = date(today.year, 1, 1)
        key_product_lines = ['Barcode', 'RFID', 'PRI', 'Reach']
        type = "RMA request"

        # Loop through each product line and create a separate analysis section
        for product in key_product_lines:
            st.markdown(f"#### Analysis for: **{product}**")

            # Filter for open cases, within YTD, for the specific product line in this loop iteration
            product_specific_data = df[
                (df['Opened Date'] >= pd.to_datetime(start_of_year)) &
                (df['Opened Date'] <= pd.to_datetime(today)) &
                (df['Status'].isin(OPEN_STATUSES)) &
                (df['Product Line'] == product) &
                (df['Type']!=type)
            ].copy()

            # --- THIS IS THE ADDED LINE ---
            # Display the total count for the current product using a metric card.
            st.metric(label="Total Open Cases (YTD)", value=len(product_specific_data))

            if not product_specific_data.empty:
                drill_col1, drill_col2, drill_col3 = st.columns(3)
                
                with drill_col1:
                    if 'Product Model' in df.columns:
                        model_counts = product_specific_data.groupby('Product Model').size().reset_index(name='Count')
                        if not model_counts.empty:
                            fig_model = px.pie(
                                model_counts, 
                                values='Count', 
                                names='Product Model', 
                                title=f'Breakdown by Model',
                                color_discrete_sequence=px.colors.qualitative.Vivid
                            )
                            fig_model.update_traces(textinfo='percent+value')
                            fig_model.update_layout(height=450)
                            st.plotly_chart(fig_model, use_container_width=True)
                            create_download_buttons(fig_model, model_counts, f"ytd_{product}_by_model")

                with drill_col2:
                    if 'Case Reason' in df.columns:
                        reason_counts = product_specific_data.groupby('Case Reason').size().reset_index(name='Count')
                        if not reason_counts.empty:
                            fig_reason = px.pie(
                                reason_counts, 
                                values='Count', 
                                names='Case Reason', 
                                title=f'Breakdown by Case Reason',
                                color_discrete_sequence=px.colors.qualitative.Vivid
                            )
                            fig_reason.update_traces(textinfo='percent+value')
                            fig_reason.update_layout(height=450)
                            st.plotly_chart(fig_reason, use_container_width=True)
                            create_download_buttons(fig_reason, reason_counts, f"ytd_{product}_by_reason")
                
                with drill_col3:
                    if 'Case Owner' in df.columns:
                        owner_counts = product_specific_data.groupby('Case Owner').size().reset_index(name='Count')
                        if not owner_counts.empty:
                            fig_owner = px.pie(
                                owner_counts, 
                                values='Count', 
                                names='Case Owner', 
                                title=f'Breakdown by Case Owner',
                                color_discrete_sequence=px.colors.qualitative.Vivid
                            )
                            fig_owner.update_traces(textinfo='percent+value')
                            fig_owner.update_layout(height=450)
                            st.plotly_chart(fig_owner, use_container_width=True)
                            create_download_buttons(fig_owner, owner_counts, f"ytd_{product}_by_owner")
            # The info message below will now only show if the count is zero
            elif len(product_specific_data) == 0:
                st.info(f"No open cases found for '{product}' from the start of the year to date.")
            
            st.markdown("---") # Add a separator after each product line's analysis
        st.header("Year-to-Date Open Case Backlog Analysis")
        st.info("This analysis shows the backlog of open cases from January 1st to today, assigned only to Users Defined in .env file.")

        today = date.today()
        start_of_year = date(today.year, 1, 1)

        # First, ensure the 'Case Owner' column exists to prevent errors
        if 'Case Owner' not in df.columns:
            st.warning("Cannot perform YTD Backlog Analysis: The 'Case Owner' column is missing.")
        else:
            # Define the specific owners to include in the analysis
            allowed_owners = ['Akhila Kotha', 'Manasa Lakshmi', 'Surendra Moilla']

            # Apply all filters: date range, open statuses, and the allowed case owners
            ytd_open_cases = df[
                (df['Opened Date'] >= pd.to_datetime(start_of_year)) &
                (df['Opened Date'] <= pd.to_datetime(today)) &
                (df['Status'].isin(OPEN_STATUSES)) &
                (df['Case Owner'].isin(allowed_owners))
            ].copy()

            if not ytd_open_cases.empty:
                # Display the total count based on the filters
                st.metric(label="Total Open Cases (YTD, Filtered Owners)", value=len(ytd_open_cases))

                # --- NEW: Create the Excel file in memory ---
                output_excel = io.BytesIO()
                with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                    ytd_open_cases.to_excel(writer, index=False, sheet_name='YTD_Backlog_Data')
                excel_data = output_excel.getvalue()

                # --- NEW: Add the download button for the Excel report ---
                st.download_button(
                    label="📥 Download YTD Backlog Report",
                    data=excel_data,
                    file_name="ytd_open_case_backlog.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.markdown("---")
                
                # The rest of the pie charts will now correctly reflect the filtered data
                ytd_row1_col1, ytd_row1_col2 = st.columns(2)
                ytd_row2_col1, ytd_row2_col2 = st.columns(2)

                with ytd_row1_col1:
                    ytd_product_counts = ytd_open_cases.groupby('Product Line').size().reset_index(name='Count')
                    fig_ytd_product = px.pie(ytd_product_counts, values='Count', names='Product Line', title='By Product Line')
                    st.plotly_chart(fig_ytd_product, use_container_width=True)
                    create_download_buttons(fig_ytd_product, ytd_product_counts, "ytd_backlog_by_product")
                
                with ytd_row1_col2:
                    if 'Product Model' in df.columns:
                        ytd_model_counts = ytd_open_cases.groupby('Product Model').size().reset_index(name='Count')
                        fig_ytd_model = px.pie(ytd_model_counts, values='Count', names='Product Model', title='By Product Model')
                        st.plotly_chart(fig_ytd_model, use_container_width=True)
                        create_download_buttons(fig_ytd_model, ytd_model_counts, "ytd_backlog_by_model")
                
                with ytd_row2_col1:
                    if 'Case Reason' in df.columns:
                        ytd_reason_counts = ytd_open_cases.groupby('Case Reason').size().reset_index(name='Count')
                        fig_ytd_reason = px.pie(ytd_reason_counts, values='Count', names='Case Reason', title='By Case Reason')
                        st.plotly_chart(fig_ytd_reason, use_container_width=True)
                        create_download_buttons(fig_ytd_reason, ytd_reason_counts, "ytd_backlog_by_reason")
                
                with ytd_row2_col2:
                    ytd_owner_counts = ytd_open_cases.groupby('Case Owner').size().reset_index(name='Count')
                    fig_ytd_owner = px.pie(ytd_owner_counts, values='Count', names='Case Owner', title='By Case Owner')
                    st.plotly_chart(fig_ytd_owner, use_container_width=True)
                    create_download_buttons(fig_ytd_owner, ytd_owner_counts, "ytd_backlog_by_owner")
            else:
                st.info("No open cases found for the specified owners from the start of the year to date.")

        st.markdown("---")
        # --- Year-to-Date Performance Trends ---
        st.header("Year-to-Date Performance Trends")
        st.info("These charts analyze trends from January 1st of the current year until today, independent of the date filter above.")

        today = date.today()
        start_of_year = date(today.year, 1, 1)

        # --- Specific owners to consider ---


        # --- Chart: Average Case Age (Only Open Cases) ---
        st.markdown("##### Average Case Age (YTD)")
        st.caption("This chart shows the average number of days open cases have remained open for the **Barcode, RFID, PRI, and Reach** product lines, calculated weekly from the start of the year.")

        key_product_lines = ['Barcode', 'RFID', 'PRI', 'Reach']

        # ✅ Filter only open cases for selected owners
        all_open_cases_ytd = df[
            (df['Status'].isin(OPEN_STATUSESAVG)) &
            (df['Product Line'].isin(key_product_lines)) &
            (df['Case Owner'].isin(selected_owners)) &
            (df['Type'] != type)
        ].copy()

        if not all_open_cases_ytd.empty:
            # Ensure Opened Date is datetime
            all_open_cases_ytd['Opened Date'] = pd.to_datetime(all_open_cases_ytd['Opened Date'], errors='coerce')
            all_open_cases_ytd = all_open_cases_ytd.dropna(subset=['Opened Date'])

            daily_avg_age_data = []

            # --- Calculate age dynamically ---
            for day in pd.date_range(start=start_of_year, end=today):
                open_on_day = all_open_cases_ytd[all_open_cases_ytd['Opened Date'] <= day]
                if not open_on_day.empty:
                    ages = (day - open_on_day['Opened Date']).dt.days
                    avg_age = ages.mean()
                    daily_avg_age_data.append({'Date': day, 'Average Age (Days)': avg_age})

            if daily_avg_age_data:
                trend_df = pd.DataFrame(daily_avg_age_data).set_index('Date')
                weekly_trend_df = trend_df.resample('W-Mon').mean().reset_index()
                weekly_trend_df['Week Number'] = weekly_trend_df.index

                # --- Plotly Chart ---
                fig_age_trend = px.line(
                    weekly_trend_df,
                    x='Week Number',
                    y='Average Age (Days)',
                    title=f"Weekly Trend of Avg. Open Case Age (YTD) - Owners: {', '.join(selected_owners)}",
                    markers=True,
                    line_shape='spline'
                )

                fig_age_trend.update_layout(
                    height=500,
                    yaxis_title="Average Case Age (Days)",
                    xaxis_title="Week Number (Since Start of Year)"
                )

                # Make x-axis labels readable
                fig_age_trend.update_xaxes(
                    tickmode='linear',
                    dtick=1,
                    tickangle=-45,
                    tickfont=dict(size=10)
                )

                st.plotly_chart(fig_age_trend, use_container_width=True)
                owners_str = "_".join([o.replace(" ", "_") for o in selected_owners])
                create_download_buttons(fig_age_trend, weekly_trend_df, f"ytd_average_case_age_{owners_str}")

        else:
            st.warning("No open cases found for the selected product lines and owners.")

        st.markdown("---")




        # --- Chart 2: Average Resolution Time (YTD) Considering Only Specific Statuses ---
        st.markdown("##### Average Resolution Time (YTD)")
        st.caption("This chart shows the average number of days taken to close or the current age of cases (only specific open and closed statuses), calculated weekly from the start of the year.")

        OPEN_STATUSES = ['New', 'In Process', 'Waiting for customer response', 'ON HOLD (Bug/Enhancement)', 'Reopened']
        CLOSED_STATUSES = ['Closed - Complete']

        if 'Case Last Modified Date' in df.columns:

            # --- Filter dataset for relevant statuses ---
            relevant_cases_ytd = df[
                (df['Status'].isin(OPEN_STATUSESAVG + CLOSED_STATUSES)) &
                (df['Opened Date'] >= pd.to_datetime(start_of_year)) &
                (df['Type'] != type)
            ].copy()

            if not relevant_cases_ytd.empty:
                weekly_avg_data = []

                # Iterate weekly through the year
                for week_start in pd.date_range(start=start_of_year, end=today, freq='W-MON'):

                    # --- OPEN CASES: calculate how long they've been open so far ---
                    open_cases = relevant_cases_ytd[
                        relevant_cases_ytd['Status'].isin(OPEN_STATUSESAVG) &
                        (relevant_cases_ytd['Opened Date'] <= week_start)
                    ]
                    avg_open_age = None
                    if not open_cases.empty:
                        open_cases['Age (Days)'] = (week_start - open_cases['Opened Date']).dt.days
                        avg_open_age = open_cases['Age (Days)'].mean()

                    # --- CLOSED CASES: calculate resolution time ---
                    closed_cases = relevant_cases_ytd[
                        (relevant_cases_ytd['Status'].isin(CLOSED_STATUSES)) &
                        (relevant_cases_ytd['Case Last Modified Date'] <= week_start)
                    ]
                    avg_closed_time = None
                    if not closed_cases.empty:
                        closed_cases['Resolution Time (Days)'] = (
                            closed_cases['Case Last Modified Date'] - closed_cases['Opened Date']
                        ).dt.days
                        closed_cases = closed_cases[closed_cases['Resolution Time (Days)'] >= 0]
                        avg_closed_time = closed_cases['Resolution Time (Days)'].mean()

                    # Combine both into a single average if any data exists
                    valid_averages = [v for v in [avg_open_age, avg_closed_time] if v is not None]
                    if valid_averages:
                        combined_avg = sum(valid_averages) / len(valid_averages)
                        weekly_avg_data.append({
                            'Week Start': week_start,
                            'Average Time (Days)': combined_avg
                        })

                if weekly_avg_data:
                    weekly_trend_df = pd.DataFrame(weekly_avg_data)
                    weekly_trend_df['Week Number'] = range(1, len(weekly_trend_df) + 1)

                    # --- Plotly line chart ---
                    fig_close_trend = px.line(
                        weekly_trend_df,
                        x='Week Number',
                        y='Average Time (Days)',
                        title='Weekly Trend of Avg. Case Time (Open + Closed Cases)',
                        markers=True,
                        line_shape='spline'
                    )

                    fig_close_trend.update_layout(
                        height=500,
                        yaxis_title="Average Time (Days)",
                        xaxis_title="Week Number (Since Start of Year)",
                        xaxis=dict(
                            tickmode='linear',
                            dtick=1,           # show every 4th week
                            tickangle=-30,
                            tickfont=dict(size=11),
                            automargin=True
                        ),
                        margin=dict(l=50, r=30, t=70, b=100),
                    )

                    st.plotly_chart(fig_close_trend, use_container_width=True)
                    create_download_buttons(fig_close_trend, weekly_trend_df, "ytd_average_resolution_time")
                else:
                    st.warning("No open or closed cases found with the specified statuses for YTD analysis.")
            else:
                st.warning("No cases found with the specified open/closed statuses.")
        else:
            st.warning("Missing 'Case Last Modified Date' column.")

        st.markdown("---")


        # --- Charts 3-6: Average Case Age by Product Line ---
    st.subheader("Average Case Age by Product Line (YTD) Without RMA")

    key_product_lines_for_loop = ['Barcode', 'RFID', 'PRI', 'Reach']

    for product in key_product_lines_for_loop:
        st.markdown(f"##### Trend for: **{product}**")
        st.caption("This shows the average number of days open cases for this product have been active, calculated weekly (Year-to-Date).")

        # Filter for open cases (excluding RMA type)
        product_specific_open_cases = df[
            (df['Status'].isin(OPEN_STATUSESAVG)) &
            (df['Product Line'] == product) &
            (df['Case Owner'].isin(selected_owners)) &
            (df['Type'] != type)
        ].copy()

        if not product_specific_open_cases.empty:
            # Create a DataFrame for each Monday in the year (weekly points)
            weekly_avg_age_data = []
            for week_start in pd.date_range(start=start_of_year, end=today, freq='W-MON'):
                open_on_week = product_specific_open_cases[
                    product_specific_open_cases['Opened Date'] <= week_start
                ]
                if not open_on_week.empty:
                    open_on_week['Age (Days)'] = (week_start - open_on_week['Opened Date']).dt.days
                    avg_age = open_on_week['Age (Days)'].mean()
                    weekly_avg_age_data.append({
                        'Week Start': week_start,
                        'Average Age (Days)': avg_age
                    })

            if weekly_avg_age_data:
                weekly_trend_df_product = pd.DataFrame(weekly_avg_age_data)
                weekly_trend_df_product['Week Number'] = range(1, len(weekly_trend_df_product) + 1)

                # --- Plotly line chart ---
                fig_product_trend = px.line(
                    weekly_trend_df_product,
                    x='Week Number',
                    y='Average Age (Days)',
                    markers=True,
                    line_shape='spline',
                    title=f'Average Case Age Trend (YTD) - {product}'
                )

                fig_product_trend.update_layout(
                    height=450,
                    yaxis_title="Average Case Age (Days)",
                    xaxis_title="Week Number (Since Start of Year)",
                    xaxis=dict(
                        tickmode='linear',
                        dtick=1,             # Show every week number
                        tickangle=-45,       # Rotate labels slightly
                        tickfont=dict(size=10),
                        automargin=True
                    ),
                    margin=dict(l=50, r=30, t=70, b=120),
                )

                st.plotly_chart(fig_product_trend, use_container_width=True)
                create_download_buttons(fig_product_trend, weekly_trend_df_product, f"ytd_avg_case_age_{product}")

        else:
            st.info(f"No open cases found for '{product}' to analyze YTD age trend.")


else:
    st.info("Please upload an Excel file to get started.")














































