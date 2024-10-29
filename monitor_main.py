from smart_value import tools
from smart_value.tools.find_docs import macro_monitor_file_path


def update_monitor(is_outdated=False):
    """Update the stock monitor Excel"""

    if is_outdated:
        tools.macro_monitor.update_marco(macro_monitor_file_path, "Free")
    tools.stock_monitor.update_monitor()


if __name__ == '__main__':
    update_monitor()
