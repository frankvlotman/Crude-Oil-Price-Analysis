import pandas as pd
import matplotlib.pyplot as plt
from adjustText import adjust_text
from tkinter import Tk, Label, Button, messagebox, ttk, Entry
from tkcalendar import DateEntry
from datetime import datetime

# Load the Excel file, skipping the first three rows
file_path = 'C:\\Users\\Frank\\Desktop\\Brent Crude Prices.xlsx'
df = pd.read_excel(file_path, skiprows=3)

# Print the column names and the first few rows to inspect the data
print("Column Names:", df.columns)
print(df.head())

# Strip any leading/trailing spaces from column names
df.columns = df.columns.str.strip()

# Ensure 'Date' column exists
if 'Date' not in df.columns:
    raise KeyError("The 'Date' column is not found in the dataset. Please check the column names.")

# Convert Date column to datetime
df['Date'] = pd.to_datetime(df['Date'])

def analyze_data(start_date, end_date):
    # Convert start_date and end_date to datetime64[ns]
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)

    # Filter data by date range
    mask = (df['Date'] >= start_date) & (df['Date'] <= end_date)
    filtered_df = df[mask].dropna()
    
    if filtered_df.empty:
        messagebox.showerror("Error", "No data available for the selected date range")
        return None

    return filtered_df

def calculate_total_value(quantity, price_per_barrel):
    # Calculate the total value
    total_value = quantity * price_per_barrel
    return total_value

def plot_prices(filtered_df, quantity):
    # Ensure the data is sorted by date in ascending order
    filtered_df = filtered_df.sort_values(by='Date')

    # Get the most recent closing price
    most_recent_price = filtered_df['Close'].iloc[-1]
    print(f"Most Recent Closing Price: {most_recent_price}")  # Debug print
    
    # Calculate the 'Total Value' for each date
    filtered_df['Total Value'] = filtered_df['Close'].apply(
        lambda price: calculate_total_value(quantity, price)
    )
    
    # Calculate the difference between the most recent date's 'Total Value' and each previous date's 'Total Value'
    latest_total_value = calculate_total_value(quantity, most_recent_price)
    filtered_df['Difference'] = latest_total_value - filtered_df['Total Value']

    # Plot the closing prices over time
    plt.figure(figsize=(14, 7))
    plt.plot(filtered_df['Date'], filtered_df['Close'], marker='o', linestyle='-', color='blue')
    
    # Add labels for each data point
    for i, row in filtered_df.iterrows():
        plt.text(row['Date'], row['Close'], f"{row['Close']:.2f}", fontsize=9, ha='center', color='green', va='bottom')

    # Adjust the labels to avoid overlap
    texts = [plt.text(row['Date'], row['Close'] - 0.2, f"Diff: {row['Difference']:.2f}", fontsize=9, ha='center', color='purple') ##################################################################3
             for i, row in filtered_df.iterrows()]

    adjust_text(texts, arrowprops=dict(arrowstyle='->', color='red'))

    # Add the total value for the most recent date
    latest_date = filtered_df['Date'].max()
    plt.text(latest_date, filtered_df.loc[filtered_df['Date'] == latest_date, 'Close'].values[0], 
             f"Total Value: {latest_total_value:.2f}", fontsize=12, ha='center', color='blue', va='top')
    
    plt.title('Closing Prices Over Time - By Value Difference on Total Value From Most Recent Date')
    plt.xlabel('Date')
    plt.ylabel('Closing Price')
    plt.grid(True)
    plt.xticks(rotation=45)

    plt.tight_layout()
    plt.show()

def update_table(filtered_df):
    # Clear the table
    for row in table.get_children():
        table.delete(row)
    
    # Insert the new data
    for index, row in filtered_df.iterrows():
        tag = 'evenrow' if index % 2 == 0 else 'oddrow'
        table.insert("", "end", values=(row['Date'].date(), row['Open'], row['Close']), tags=(tag,))

def on_submit():
    start_date = start_date_entry.get_date()
    end_date = end_date_entry.get_date()
    
    if start_date > end_date:
        messagebox.showerror("Error", "Start date must be before end date")
        return

    filtered_df = analyze_data(start_date, end_date)
    
    if filtered_df is not None:
        quantity = int(quantity_entry.get())
        
        plot_prices(filtered_df, quantity)
        update_table(filtered_df)

# Create GUI
root = Tk()
root.title("Crude Oil Price Analysis - Total Value Gain or Loss")

# Adjusted the layout of the date labels and inputs
Label(root, text="Start Date:").grid(row=0, column=0, padx=(10, 5), pady=5, sticky='e')
start_date_entry = DateEntry(root, date_pattern='yyyy-MM-dd')
start_date_entry.grid(row=0, column=1, padx=(5, 10), pady=5, sticky='w')

Label(root, text="End Date:").grid(row=1, column=0, padx=(10, 5), pady=5, sticky='e')
end_date_entry = DateEntry(root, date_pattern='yyyy-MM-dd')
end_date_entry.grid(row=1, column=1, padx=(5, 10), pady=5, sticky='w')

Label(root, text="Quantity:").grid(row=2, column=0, padx=(10, 5), pady=5, sticky='e')
quantity_entry = Entry(root)
quantity_entry.grid(row=2, column=1, padx=(5, 10), pady=5, sticky='w')

submit_button = Button(root, text="Submit", command=on_submit)
submit_button.grid(row=3, column=0, columnspan=2, pady=10)

# Create table view
columns = ("Date", "Open", "Close")
table = ttk.Treeview(root, columns=columns, show="headings")
for col in columns:
    table.heading(col, text=col, anchor='center')
    table.column(col, anchor='center')
table.grid(row=4, column=0, columnspan=2, sticky='nsew')

# Add scrollbar to the table
scrollbar = ttk.Scrollbar(root, orient="vertical", command=table.yview)
table.configure(yscrollcommand=scrollbar.set)
scrollbar.grid(row=4, column=2, sticky='ns')

# Apply styling
style = ttk.Style()
style.theme_use("default")
style.configure("Treeview.Heading", anchor="center")
style.configure("Treeview", rowheight=25)
style.map("Treeview", background=[('selected', '#347083')])

# Add striped row tags
table.tag_configure('oddrow', background="lightgrey")
table.tag_configure('evenrow', background="white")

root.mainloop()
