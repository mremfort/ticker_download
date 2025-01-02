import streamlit as st
import yfinance as yf
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime

# Function to fetch stock or index data and
st.set_page_config(layout="wide", page_title="Data Viewer", page_icon=':bar_chart:')

def fetch_stock_data(ticker, start_date, end_date):
    stock = yf.Ticker(ticker)

    # Try to fetch stock history data
    try:
        # Use actions=True to ensure adjusted close is included in the data
        stock_data = stock.history(period="1d", start=start_date, end=end_date, actions=True)
        if stock_data.empty:
            raise ValueError(f"No data returned for {ticker}. Please check the ticker symbol.")
    except Exception as e:
        st.error(f"Error fetching data for {ticker}: {str(e)}")
        return None, None, None, None

    # Fetch company or index info
    info = stock.info
    company_name = info.get('longName', 'N/A')
    company_desc = info.get('longBusinessSummary', 'No description available')

    # Get the first available date for start_date
    first_available_date = stock_data.index.min().tz_localize(None)

    # Get current bid price
    bid_price = info.get('bid', 'N/A')

    return stock_data, company_name, company_desc, first_available_date, bid_price
# Function to prepare data for Excel export
def prepare_for_export(stock_data, ticker, start_date, end_date):
    # Create a new Excel file with the ticker and date range as the sheet name
    filename = f"{ticker}_{start_date}_{end_date}.xlsx"

    # Clean the filename to remove special characters that may be problematic
    filename = filename.replace(":", "_").replace("^", "_")  # Replace any invalid characters

    # Optionally, specify a directory where you know you have permission to write files:
    # filename = f"C:/Users/{your_username}/Documents/{filename}"

    # Convert stock data to DataFrame
    stock_data['Date'] = stock_data.index.tz_localize(None)  # Remove timezone information
    stock_data.reset_index(drop=True, inplace=True)

    # Save data to Excel
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        stock_data.to_excel(writer, index=False, sheet_name=ticker)

    return filename

# Function to calculate performance over various time periods
def calculate_performance(stock_data):
    # Calculate the percentage change for each period: 1D, 1M, 1Y, 3Y, 5Y, 10Y, YTD
    performance = {}

    # Ensure that 'Adj Close' is available, otherwise fallback to 'Close'
    if 'Adj Close' not in stock_data.columns:
        stock_data['Adj Close'] = stock_data['Close']  # Fallback to using Close if Adj Close is missing

    # 1D performance
    performance["1D"] = (stock_data['Adj Close'].pct_change(periods=1).iloc[-1]) * 100

    # 1M performance
    performance["1M"] = (stock_data['Adj Close'].pct_change(periods=30).iloc[-1]) * 100

    # 1Y performance
    performance["1Y"] = (stock_data['Adj Close'].pct_change(periods=252).iloc[-1]) * 100  # Approx 252 trading days in a year

    # 3Y performance
    performance["3Y"] = (stock_data['Adj Close'].pct_change(periods=252*3).iloc[-1]) * 100

    # 5Y performance
    performance["5Y"] = (stock_data['Adj Close'].pct_change(periods=252*5).iloc[-1]) * 100

    # 10Y performance
    performance["10Y"] = (stock_data['Adj Close'].pct_change(periods=252*10).iloc[-1]) * 100

    # YTD (Year-to-Date) performance
    start_of_year = stock_data[stock_data.index.month == 1].iloc[0]  # Get first data point of the year
    performance["YTD"] = ((stock_data['Adj Close'].iloc[-1] - start_of_year['Adj Close']) / start_of_year['Adj Close']) * 100

    return performance

# Streamlit UI
st.title("Stock and Index Data Viewer")

with st.sidebar:
    st.title("Settings")
    # Input for ticker symbol
    ticker = st.text_input("Enter Stock or Index Ticker Symbol", "")

# Initialize start_date and end_date as None
start_date = None
end_date = None

# Fetch data when ticker is entered
if ticker:

    # Fetch data from Yahoo Finance
    stock_data, company_name, company_desc, first_available_date, bid_price = fetch_stock_data(ticker, "2000-01-01", datetime.today())

    if stock_data is not None:
        # Set start_date to the first available month
        with st.sidebar:
            start_date = st.date_input("Start Date", first_available_date, min_value=None, max_value=datetime.today())
            end_date = st.date_input("End Date", datetime.today(), min_value=start_date, max_value=datetime.today())

        # Display company or index info
        st.subheader(f"Name: {company_name} - ${bid_price}")
        st.write(company_desc)

        # Fetch the stock data again based on the selected date range
        stock_data_filtered, _, _, _, _ = fetch_stock_data(ticker, start_date, end_date)

        if stock_data_filtered is not None:
            # Calculate performance for different periods
            performance = calculate_performance(stock_data_filtered)
            col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
            # Display performance at the top
            st.subheader(f"{company_name} Performance")
            with col1:
                st.write(f"1D: {performance['1D']:.2f}%")
            with col2:
                st.write(f"1M: {performance['1M']:.2f}%")
            with col3:
                st.write(f"1Y: {performance['1Y']:.2f}%")
            with col4:
                st.write(f"3Y: {performance['3Y']:.2f}%")
            with col5:
                st.write(f"5Y: {performance['5Y']:.2f}%")
            with col6:
                st.write(f"10Y: {performance['10Y']:.2f}%")
            with col7:
                st.write(f"YTD: {performance['YTD']:.2f}%")  # YTD Performance

            # Ensure the index is set to datetime
            stock_data_filtered.index = pd.to_datetime(stock_data_filtered.index)

            # Resample to Monthly data and get the last day of the month
            st.subheader("Monthly Performance")
            stock_data_monthly = stock_data_filtered.resample('ME').last()  # Resample to monthly data (last data point of each month)

            # Calculate Percentage Change Month over Month
            stock_data_monthly['% Change'] = stock_data_monthly['Adj Close'].pct_change() * 100

            # Format Date to Month Day, Year
            stock_data_monthly['Formatted Date'] = stock_data_monthly.index.strftime('%b %d, %Y')

            # Create Plotly Line Graph
            fig = go.Figure(data=[go.Scatter(x=stock_data_monthly.index,
                                            y=stock_data_monthly['Adj Close'],
                                            mode='lines',  # Line graph mode
                                            name='Adj Close',  # Label for the line
                                            line=dict(color='blue'))])  # Customize line color

            fig.update_layout(title=f"{ticker} Monthly Performance", xaxis_title="Date", yaxis_title="Price")

            # Update x-axis format for better visibility
            fig.update_xaxes(tickformat='%b %d, %Y')

            st.plotly_chart(fig)

            # Display Data Table with only Adjusted Close and % Change
            st.subheader("Stock Data Table")
            st.dataframe(stock_data_monthly[['Adj Close', '% Change']])

            # Initialize excel_filename as None
            with st.sidebar:
                # Export Data to Excel
                st.subheader("Export Data to Excel")

                excel_filename = None

                if st.button("Export to Excel"):
                    # Generate Excel file when button is clicked
                    excel_filename = prepare_for_export(
                        stock_data_monthly[['Formatted Date', 'Adj Close', '% Change']], ticker, start_date, end_date)
                    st.success(f"File Created successfully!")

                # If excel_filename is generated, provide download link
                if excel_filename:
                    with open(excel_filename, 'rb') as f:
                        st.download_button(
                            label="Download Excel File",
                            data=f,
                            file_name=excel_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
