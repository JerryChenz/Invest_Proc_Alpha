import os
import xlwings
from smart_value.tools.find_docs import get_template_paths, new_latest_model
from smart_value.tools.model_dash import model_pos
from smart_value.data import yq_data as yq


def new_stock_model(ticker):
    """Creates a new model if it doesn't already exist, then update.

    :param ticker: the string ticker of the stock
    """

    template_path_list = get_template_paths()

    # New model path
    model_name = ticker + "_" + os.path.basename(template_path_list[0])
    model_path = new_latest_model(model_name)
    # update the model
    update_new_model(ticker, model_name, model_path)


def update_new_model(ticker, model_name, model_path):
    """Update the model.

    :param ticker: the string ticker of the stock
    :param model_name: the model file name
    :param model_path: the model file path
    """
    # update the new model
    print(f'Updating {model_name}...')
    company = yq.YqStock(ticker)

    with xlwings.App(visible=False) as app:
        model_xl = app.books.open(model_path)
        new_inputs(model_xl.sheets('Inputs'), company)
        model_xl.save(model_path)
        model_xl.close()


def new_inputs(model_xl, company):
    """Get the new inputs

    :param model_xl: the Excel file of the model
    :param company: the YqStock object of the stock
    """

    input_sheet = model_xl.sheets('Inputs')
    dash_sheet = model_xl.sheets('Dashboard')
    pass
