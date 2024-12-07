import tkinter as tk
from smart_value.tools import model_update, stock_monitor


def full_update():
    # Step 1: update the xls with the latest template, and market data
    model_inputs.update_models(False)
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

    # Run the application
    root.mainloop()


if __name__ == '__main__':
    invest_proc()
