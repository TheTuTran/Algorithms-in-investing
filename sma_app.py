from datetime import datetime
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from tkcalendar import DateEntry
import pandas as pd
import numpy as np
import yfinance as yf
import os

def fetch_stock_data(ticker, start_date, end_date):
    data = yf.download(ticker, start=start_date, end=end_date)
    return data['Close']

def calculate_sma(data, window):
    return data.rolling(window=window).mean()

def generate_signals(data, short_window, long_window):
    signals = pd.DataFrame(index=data.index)
    signals['price'] = data
    signals['short_sma'] = calculate_sma(data, short_window)
    signals['long_sma'] = calculate_sma(data, long_window)
    signals['signal'] = 0.0

    # Generate buy/sell signals
    signals['signal'][short_window:] = np.where(signals['short_sma'][short_window:] > signals['long_sma'][short_window:], 1.0, 0.0)
    signals['positions'] = signals['signal'].diff()

    # Initialize the profit tracking columns
    signals['signal_profit'] = np.nan
    signals['cumulative_profit'] = 0.0

    # Track the buy price and sell price for calculating profit
    buy_price = None
    for i in range(len(signals)):
        if signals['positions'][i] == 1:  # A buy signal
            buy_price = signals['price'][i]
        elif signals['positions'][i] == -1 and buy_price is not None:  # A sell signal
            sell_price = signals['price'][i]
            profit = sell_price - buy_price
            signals.at[signals.index[i], 'signal_profit'] = profit
            buy_price = None  # Reset buy price for the next trade

    # Calculate cumulative profit
    signals['cumulative_profit'] = signals['signal_profit'].cumsum().fillna(method='ffill')

    return signals

def test_sma_combinations(data, short_window_range, long_window_range):
    best_profit = -np.inf
    best_combination = (0, 0)
    results = []

    for short_window in range(short_window_range[0], short_window_range[1] + 1):
        for long_window in range(long_window_range[0], long_window_range[1] + 1):
            if short_window >= long_window:
                continue  # Skip invalid combinations
            signals = generate_signals(data, short_window, long_window)
            cumulative_profit = signals['cumulative_profit'].iloc[-1]
            profitable_signals = signals['signal_profit'].dropna()
            profit_percentage = (profitable_signals > 0).mean() * 100  # Percentage of profitable trades
            number_of_trades = len(profitable_signals)
            results.append((short_window, long_window, cumulative_profit, profit_percentage, number_of_trades))
            if cumulative_profit > best_profit:
                best_profit = cumulative_profit
                best_combination = (short_window, long_window)

    return results, best_combination, best_profit

class StockApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Simple Moving Average Analysis Tool")
        self.root.geometry("1200x600")
        self.root.configure(bg='#22262b')  # Set the background color for the root window
        
        self.create_widgets()

    def create_widgets(self):
        # Define the style for ttk widgets
        style = ttk.Style()
        style.theme_use('default')  # Start with the default theme
        
        # Configure the style for the Treeview
        style.configure("Treeview",
                        background="#22262b",
                        foreground="#CCCCCC",
                        fieldbackground="#22262b")
        style.map('Treeview', background=[('selected', '#4E9F3D')])

        # Configure the style for other ttk widgets (e.g., Button, Label)
        style.configure("TButton",
                        background="#22262b",
                        foreground="#FFFFFF",
                        borderwidth=1)
        style.map("TButton",
                background=[('active', '#22262b')],
                foreground=[('active', '#FFFFFF')])


        # Input Frame
        input_frame = ttk.Frame(self.root, padding="10")
        input_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        input_frame.columnconfigure(1, weight=1)

        # Results Frame
        results_frame = ttk.Frame(self.root, padding="10")
        results_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        # Input Fields
        dir_path = os.path.dirname(os.path.realpath(__file__))
        csv_file_path = os.path.join(dir_path, 'nasdaq_tickers.csv')
        ticker_df = pd.read_csv(csv_file_path)
        ticker_list = ticker_df['Symbol'].astype(str).tolist()
        ttk.Label(input_frame, text="Ticker:").grid(column=0, row=0, sticky=tk.W)
        self.ticker = AutocompleteEntry(ticker_list, input_frame, width=50)
        self.ticker.grid(column=1, row=0, sticky=(tk.W, tk.E))

        ttk.Label(input_frame, text="Start Date:").grid(column=0, row=1, sticky=tk.W)
        self.start_date = DateEntry(input_frame, width=12, background='darkblue', foreground='white', borderwidth=2)
        self.start_date.grid(column=1, row=1, sticky=(tk.W, tk.E))

        ttk.Label(input_frame, text="End Date:").grid(column=0, row=2, sticky=tk.W)
        self.end_date = DateEntry(input_frame, width=12, background='darkblue', foreground='white', borderwidth=2)
        self.end_date.grid(column=1, row=2, sticky=(tk.W, tk.E))

        ttk.Label(input_frame, text="Short Term Window Range (eg. 1-10):").grid(column=0, row=3, sticky=tk.W)
        self.short_term_window = ttk.Entry(input_frame)
        self.short_term_window.grid(column=1, row=3, sticky=(tk.W, tk.E))

        ttk.Label(input_frame, text="Long Term Window Range (eg. 20-50):").grid(column=0, row=4, sticky=tk.W)
        self.long_term_window = ttk.Entry(input_frame)
        self.long_term_window.grid(column=1, row=4, sticky=(tk.W, tk.E))

        ttk.Button(input_frame, text="Analyze", command=self.analyze).grid(column=1, row=5, sticky=tk.W, pady=4)
        self.download_button = ttk.Button(results_frame, text="Download Signals", state='disabled', command=self.download_signals)
        self.download_button.pack(pady=10)

        # Results Table
        self.results_tree = ttk.Treeview(results_frame, columns=("Short Window", "Long Window", "Cumulative Profit", "Profit Percentage", "Number of Trades"), show="headings")
        results_scroll = ttk.Scrollbar(results_frame, orient="vertical", command=self.results_tree.yview)
        results_scroll.pack(side="right", fill="y")
        self.results_tree.configure(yscrollcommand=results_scroll.set)
        self.results_tree.pack(side="left", fill="both", expand=True)

        # Attach the sort function to each column
        for col in self.results_tree['columns']:
            self.results_tree.heading(col, text=col, command=lambda _col=col, _tv=self.results_tree: self.treeview_sort_column(_tv, _col, False))
        self.results_tree.bind("<Double-1>", self.on_result_select)  # Bind double click
        self.results_tree.pack(expand=True, fill='both')

    def treeview_sort_column(self, tv, col, reverse):
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        try:
            l.sort(key=lambda t: float(t[0]), reverse=reverse)  # Convert to float for sorting
        except ValueError:
            l.sort(reverse=reverse)  # Fallback to default sort if conversion fails

        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)

        tv.heading(col, command=lambda _col=col: self.treeview_sort_column(tv, _col, not reverse))

    def analyze(self):
        ticker = self.ticker.get()
        start_date = self.start_date.get()
        end_date = self.end_date.get()
        start_date_formatted = datetime.strptime(start_date, '%m/%d/%y').strftime('%Y-%m-%d')
        end_date_formatted = datetime.strptime(end_date, '%m/%d/%y').strftime('%Y-%m-%d')
        short_term_range = self.parse_range(self.short_term_window.get())
        long_term_range = self.parse_range(self.long_term_window.get())

        if not all([ticker, start_date, end_date, short_term_range, long_term_range]):
            messagebox.showerror("Error", "All fields are required and must be valid.")
            return

        try:
            data = fetch_stock_data(ticker, start_date_formatted, end_date_formatted)
            results, best_combination, best_profit = test_sma_combinations(data, short_term_range, long_term_range)

            # Clear existing items in the tree
            for i in self.results_tree.get_children():
                self.results_tree.delete(i)

            # Insert new results into the tree
            for result in results:
                self.results_tree.insert("", "end", values=result)

            messagebox.showinfo("Analysis Complete", "Analysis complete. Double-click a row to view signals.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def parse_range(self, range_str):
        try:
            start, end = map(int, range_str.split('-'))
            return (start, end)
        except ValueError:
            messagebox.showerror("Error", "Invalid range format. Please use the format start-end (e.g., 1-50).")
            return None
    
    def on_result_select(self, event):
        item = self.results_tree.identify('item', event.x, event.y)
        if not item:  # If no item was clicked, do nothing
            return
        
        for selected_item in self.results_tree.selection():
            item = self.results_tree.item(selected_item)
            short_window, long_window, _, _, _ = item['values']
            data = fetch_stock_data(self.ticker.get(), self.start_date.get_date(), self.end_date.get_date())
            self.current_signals = generate_signals(data, short_window, long_window)

            # Create a new window to display signals
            signals_window = tk.Toplevel(self.root)
            signals_window_title = f"Signals of {short_window} SMA and {long_window} SMA"
            signals_window.title(signals_window_title)
            signals_tree = ttk.Treeview(signals_window, columns=("Date", "Price", "Short SMA", "Long SMA", "Hold", "Position", "Signal Profit", "Cumulative Profit"), show="headings")
            signals_tree.pack(expand=True, fill='both')

            signals_scroll = ttk.Scrollbar(signals_window, orient="vertical", command=signals_tree.yview)
            signals_scroll.pack(side="right", fill="y")
            signals_tree.configure(yscrollcommand=signals_scroll.set)
            signals_tree.pack(side="left", fill="both", expand=True)

            for col in signals_tree['columns']:
                signals_tree.heading(col, text=col)
                signals_tree.column(col, anchor="center")
            
            for date, row in self.current_signals.iterrows():
                signals_tree.insert("", "end", values=(date, row['price'], row['short_sma'], row['long_sma'], row['signal'], row['positions'], row['signal_profit'], row['cumulative_profit']))

            # Enable the download button
            self.download_button['state'] = 'normal'
    
    def download_signals(self):
        if self.current_signals is not None:
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if file_path:  # Ensure the user didn't cancel the dialog
                try:
                    self.current_signals.to_excel(file_path, engine='openpyxl')
                    messagebox.showinfo("Success", "Signals saved successfully.")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to save file. Error: {str(e)}")
        else:
            messagebox.showwarning("Warning", "No signals to download. Please select a result first.")

class AutocompleteEntry(ttk.Entry):
    def __init__(self, autocompleteList, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.autocompleteList = autocompleteList
        self.var = self["textvariable"]
        if self.var == '':
            self.var = self["textvariable"] = tk.StringVar()

        self.var.trace('w', self.changed)
        self.bind("<Return>", self.selection)
        self.bind("<Up>", self.move_up)
        self.bind("<Down>", self.move_down)

        self.listboxUp = False

    def changed(self, name, index, mode):
        if self.var.get() == '':
            if self.listboxUp:
                self.listbox.destroy()
                self.listboxUp = False
        else:
            words = self.comparison()
            if words:
                if not self.listboxUp:
                    self.listbox = tk.Listbox(width=self["width"], height=4)
                    self.listbox.bind("<Button-1>", self.selection)
                    self.listbox.bind("<Right>", self.selection)
                    self.listbox.place(x=self.winfo_x(), y=self.winfo_y() + self.winfo_height())
                    self.listboxUp = True
                
                self.listbox.delete(0, tk.END)
                for word in words:
                    self.listbox.insert(tk.END, word)
            else:
                if self.listboxUp:
                    self.listbox.destroy()
                    self.listboxUp = False

    def selection(self, event):
        if self.listboxUp:
            self.var.set(self.listbox.get(tk.ACTIVE))
            self.listbox.destroy()
            self.listboxUp = False
            self.icursor(tk.END)

    def move_up(self, event):
        if self.listboxUp:
            if self.listbox.curselection() == (0,):
                return
            self.listbox.select_clear(self.listbox.curselection())
            self.listbox.select_set(self.listbox.curselection()[0] - 1)
            self.listbox.activate(self.listbox.curselection()[0] - 1)

    def move_down(self, event):
        if self.listboxUp:
            if self.listbox.curselection() == (self.listbox.size() - 1,):
                return
            self.listbox.select_clear(self.listbox.curselection())
            self.listbox.select_set(self.listbox.curselection()[0] + 1)
            self.listbox.activate(self.listbox.curselection()[0] + 1)

    def comparison(self):
        user_input = self.var.get().lower()
        return [str(w) for w in self.autocompleteList if str(w).lower().startswith(user_input)]

def main():
    root = tk.Tk()
    app = StockApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()