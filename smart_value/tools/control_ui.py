import tkinter as tk
from smart_value.tools import model_update, stock_monitor
from smart_value.tools.model_new import new_stock_model


def full_update():
    # Step 1: update the xls with the latest template, and market data
    model_update.update_models(False)
    # Step 2: update the monitor
    stock_monitor.update_monitor(True)


def invest_proc():
    # Create the main application window
    root = tk.Tk()
    root.title("Invest Proc_v3")

    # Create a button
    full_update_button = tk.Button(root, text="Full Update", command=full_update)
    full_update_button.pack(side=tk.LEFT)
    full_monitor_button = tk.Button(root, text="Full Monitor Update",
                                    command=lambda: stock_monitor.update_monitor(False))
    full_monitor_button.pack(side=tk.RIGHT)
    simple_monitor_button = tk.Button(root, text="Simple Monitor Update",
                                      command=lambda: stock_monitor.update_monitor(True))
    simple_monitor_button.pack(side=tk.RIGHT)

    # declaring string variable for storing symbol and comp_group
    symbol_var = tk.StringVar()
    comp_var = tk.StringVar()

    # creating a label and an entry for symbol
    symbol_label = tk.Label(root, text='symbol', font=('calibre', 10, 'bold'))
    symbol_entry = tk.Entry(root, textvariable=symbol_var, font=('calibre', 10, 'normal'))

    # creating a label and an entry for comp_group
    comp_label = tk.Label(root, text='comp_group', font=('calibre', 10, 'bold'))
    comp_entry = tk.Entry(root, textvariable=comp_var, font=('calibre', 10, 'normal'))

    # creating a button
    sub_btn = tk.Button(root, text='Submit', command=lambda: new_stock_model(ticker, comp_group))

    # placing the label and entry in the required position
    symbol_label.grid(row=0, column=0)
    symbol_entry.grid(row=0, column=1)
    comp_label.grid(row=1, column=0)
    comp_entry.grid(row=1, column=1)
    sub_btn.grid(row=2, column=1)

    # Run the application
    root.mainloop()
