import shutil
import xlwings
import pathlib
from pathlib import Path
import re
from smart_value.tools.find_docs import models_folder_path, template_folder_path

input_dict = {
    "info": "C4:C22",
    "income_statement": "C25:M33",
    "balance_sheet": "D34:M43",
    "dividend": "C44:M44",
    "Others": "C48:E87",
    "Valuation": "E91:F98"
}


def update_models():
    """Update the dashboard of the model."""

    stock_regex = re.compile(".*Stock_Valuation(?!_old)")
    negative_regex = re.compile(".*~.*")

    # finds the list of all models
    path_list = [val_file_path for val_file_path in models_folder_path.iterdir()
                 if models_folder_path.is_dir() and val_file_path.is_file()]
    opportunities_path_list = list(item for item in path_list if stock_regex.match(str(item)) and
                                   not negative_regex.match(str(item)))
    for p in opportunities_path_list:
        # Step 1: renames the existing model to mark as old
        print(f"Updating {p.name}")
        marked_path = p.rename(Path(p.parent, p.stem + "_old" + p.suffix))
        # Step 2: creates a new model with the latest template
        updated_path = new_updated_model(p.name, models_folder_path)
        # Step 3: copy the inputs from the marked path and paste to the updated model with inputs
        copy_inputs(marked_path, updated_path)
        # Step 4: delete the marked model
        # pathlib.Path.unlink(marked_path)


def new_updated_model(file_name, file_path):
    """creates a new model with the latest template, returns the new path

    :param file_name: the file name of the model
    :param file_path: the directory path for the model
    :raises FileNotFoundError: raises an exception when there is an error related to the model files or path
    """
    stock_regex = re.compile(".*Stock_Valuation")
    negative_regex = re.compile(".*~.*")

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
    else:
        # New model path
        model_path = file_path / file_name
        if not pathlib.Path(model_path).exists():
            # Creates a new model file if not already exists in cwd
            print(f'Creating {file_name}...')
            shutil.copy(template_path_list[0], model_path)
        return model_path


def copy_inputs(old_path, new_path):
    """reads the inputs of the model at marked_path, returns the inputs in a dictionary

    :param old_path: the file path of the marked model
    :param new_path: the file path of the new model
    """

    with xlwings.App(visible=False) as app:
        xl_old_model = app.books.open(old_path)
        xl_new_model = app.books.open(new_path)
        old_input_sheet = xl_old_model.sheets('Inputs')
        new_input_sheet = xl_new_model.sheets('Inputs')
        for inputs in input_dict:
            old_input_sheet.range(input_dict[inputs]).copy()
            new_input_sheet.range(input_dict[inputs]).paste(paste="formulas")
        xl_old_model.close()
        xl_new_model.save(new_path)
        xl_new_model.close()
