from smart_value.tools import stock_monitor, macro_monitor
from smart_value.tools.find_docs import macro_monitor_file_path

if __name__ == '__main__':
    # macro_monitor.update_marco(macro_monitor_file_path, "Free")
    stock_monitor.update_monitor(False)
