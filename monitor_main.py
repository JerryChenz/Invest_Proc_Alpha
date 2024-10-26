from smart_value import tools


def update_monitor():
    """Update the stock monitor Excel"""

    tools.stock_monitor.update_monitor()


if __name__ == '__main__':
    update_monitor()
