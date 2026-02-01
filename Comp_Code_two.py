import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.filedialog import asksaveasfilename
import requests
import threading
import csv

# Your AlphaVantage API key
API_KEY = "TP5VHEAWYRO2BM3Y"

# Updated dictionary for all S&P 500 sectors with sample tickers
sp500_companies = {
    "Communication Services": ["GOOGL", "META", "DIS"],
    "Consumer Discretionary": ["AMZN", "HD", "NKE"],
    "Consumer Staples": ["PG", "KO", "WMT"],
    "Energy": ["XOM", "CVX", "COP"],
    "Financials": ["JPM", "BAC", "GS"],
    "Healthcare": ["JNJ", "PFE", "MRK"],
    "Industrials": ["BA", "CAT", "UPS"],
    "Information Technology": ["AAPL", "MSFT", "NVDA"],
    "Materials": ["LIN", "SHW", "ECL"],
    "Real Estate": ["AMT", "PLD", "CCI"],
    "Utilities": ["NEE", "DUK", "SO"]
}

def format_market_cap(value):
    """
    Convert a market cap string to an integer and format it as a dollar amount with commas.
    If conversion fails (e.g., value is 'N/A'), return the original value.
    """
    try:
        value_int = int(float(value))
        return f"${value_int:,}"
    except:
        return value

def fetch_company_overview(symbol):
    """Fetch company overview data from AlphaVantage for a given ticker symbol."""
    url = "https://www.alphavantage.co/query"
    params = {
        "function": "OVERVIEW",
        "symbol": symbol,
        "apikey": API_KEY
    }
    try:
        response = requests.get(url, params=params, timeout=10)
        data = response.json()
        # Check if the required fields are in the response
        if "MarketCapitalization" in data and "ReturnOnEquityTTM" in data and "ReturnOnAssetsTTM" in data:
            return {
                "Symbol": symbol.upper(),
                "MarketCap": data.get("MarketCapitalization", "N/A"),
                "ROE": data.get("ReturnOnEquityTTM", "N/A"),
                "ROA": data.get("ReturnOnAssetsTTM", "N/A")
            }
        else:
            return None
    except Exception as e:
        print(f"Error fetching data for {symbol}: {e}")
        return None

def search_companies():
    """
    Use either manually entered tickers (from the Text widget) or a preset sector's tickers.
    Fetch data for each ticker using the AlphaVantage API.
    """
    # Get ticker symbols from the Text widget (allows multiple lines) and remove extra whitespace
    tickers_input = tickers_text.get("1.0", tk.END).strip()
    selected_sector = sector_var.get().strip()

    tickers = []
    if tickers_input:
        # Allow comma-separated or one ticker per line
        # Replace newlines with commas, then split
        tickers = [ticker.strip().upper() for ticker in tickers_input.replace("\n", ",").split(",") if ticker.strip()]
    elif selected_sector:
        tickers = sp500_companies.get(selected_sector, [])
    else:
        messagebox.showwarning("Input Error", "Please enter ticker symbols or select a sector.")
        return

    if not tickers:
        messagebox.showwarning("Input Error", "No valid ticker symbols found.")
        return

    # Disable the search button to prevent re-clicks while processing
    search_btn.config(state=tk.DISABLED)
    # Clear previous results
    for row in tree.get_children():
        tree.delete(row)
    status_label.config(text="Fetching data... (please be patient due to API rate limits)")

    def worker():
        results = []
        for ticker in tickers:
            data = fetch_company_overview(ticker)
            if data:
                results.append(data)
            else:
                print(f"No data for {ticker} or API limit reached.")
        # Update the table on the main thread
        root.after(0, lambda: update_treeview(results))
        root.after(0, lambda: search_btn.config(state=tk.NORMAL))
        root.after(0, lambda: status_label.config(text="Done."))

    threading.Thread(target=worker).start()

def update_treeview(results):
    """Update the Treeview widget with fetched company data, formatting Market Cap."""
    for item in results:
        market_cap_formatted = format_market_cap(item["MarketCap"]) if item["MarketCap"] != "N/A" else "N/A"
        tree.insert("", tk.END, values=(item["Symbol"], market_cap_formatted, item["ROE"], item["ROA"]))

def export_data():
    """Export the displayed table data to a CSV file via a file dialog."""
    data = []
    for row_id in tree.get_children():
        row = tree.item(row_id)["values"]
        data.append(row)
    if not data:
        messagebox.showwarning("No Data", "There is no data to export.")
        return

    filename = asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv")],
        title="Save as..."
    )
    if not filename:
        return  # User canceled the save dialog

    try:
        with open(filename, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["Company", "Market Cap", "ROE", "ROA"])
            writer.writerows(data)
        messagebox.showinfo("Export Successful", f"Data exported to {filename}")
    except Exception as e:
        messagebox.showerror("Export Failed", f"An error occurred while exporting: {e}")

# Set up the main GUI window
root = tk.Tk()
root.title("Comparable Companies Finder")

frame = ttk.Frame(root, padding="10")
frame.pack(fill=tk.BOTH, expand=True)

# Section for selecting a preset sector
ttk.Label(frame, text="Select S&P 500 Sector:").pack(pady=5)
sector_var = tk.StringVar()
sector_combobox = ttk.Combobox(frame, textvariable=sector_var, state="readonly")
sector_combobox['values'] = list(sp500_companies.keys())
sector_combobox.pack(pady=5)

# Separator label
ttk.Label(frame, text="OR").pack(pady=5)

# Section for entering ticker symbols manually using a multi-line Text widget
ttk.Label(frame, text="Enter Ticker Symbols (comma or newline separated):").pack(pady=5)
tickers_text = tk.Text(frame, height=3, width=50)
tickers_text.pack(pady=5)

# Button to trigger the search
search_btn = ttk.Button(frame, text="Get Comparables", command=search_companies)
search_btn.pack(pady=10)

# Status label for user feedback
status_label = ttk.Label(frame, text="")
status_label.pack(pady=5)

# Table (Treeview) for displaying the results
columns = ("Company", "Market Cap", "ROE", "ROA")
tree = ttk.Treeview(frame, columns=columns, show="headings")
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, anchor="center")
tree.pack(pady=10, fill=tk.BOTH, expand=True)

# Button to export the table data to CSV
export_btn = ttk.Button(frame, text="Export to CSV", command=export_data)
export_btn.pack(pady=10)

root.mainloop()
