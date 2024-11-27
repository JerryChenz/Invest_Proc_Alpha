import xlwings
import os
import pathlib
import re
from smart_value.tools.find_docs import models_folder_path, template_folder_path


def update_models(ticker, source):
    """Update the dashboard of the model.

    :param ticker: the string ticker of the stock
    :param source: String data source selector
    :raises FileNotFoundError: raises an exception when there is an error related to the model files or path
    """

    # finds the list of all models
    path_list = [val_file_path for val_file_path in models_folder_path.iterdir()
                 if models_folder_path.is_dir() and val_file_path.is_file()]
    for p in path_list:
        # renames the existing model to mark as old
        original_path = pathlib.Path(p)
        print(original_path)
        #new_path = original_path.rename('old_' + p)
        # if ticker in p.stem:
        #     with xlwings.App(visible=False) as app:
        #         xl_book = app.books.open(p)
        #         input_sheet = xl_book.sheets('Inputs')
        #
        #         xl_book.close()
        # pathlib.Path.unlink(new_path)


def new_model():
    """Renames the existing model to mark as old, then creates a new model and update.

    :param ticker: the string ticker of the stock
    :param source: String data source selector
    :raises FileNotFoundError: raises an exception when there is an error related to the model files or path
    """
    stock_regex = re.compile(".*Stock_Valuation")
    negative_regex = re.compile(".*~.*")

    new_bool = False

    try:
        # Check if the template exists
        if pathlib.Path(template_folder_path).exists():
            path_list = [val_file_path for val_file_path in template_folder_path.iterdir()
                         if template_folder_path.is_dir() and val_file_path.is_file()]
            template_path_list = list(item for item in path_list if stock_regex.match(str(item)) and
                                      not negative_regex.match(str(item)))
            if len(template_path_list) > 1 or len(template_path_list) == 0:
                raise FileNotFoundError("The template file error", "temp_file")
        else:
            raise FileNotFoundError("The stock_template folder doesn't exist", "temp_folder")
    except FileNotFoundError as err:
        if err.args[1] == "temp_folder":
            print("The stock_template folder doesn't exist")
        if err.args[1] == "temp_file":
            print("The template file error")
