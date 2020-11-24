"""
Parsing related functions, including reading tables and evaluating user expressions.
"""
import logging
logger = logging.getLogger('email-automation')

import re
import numpy
import pandas as pd
import xlrd
import traceback

from .typography import unite_similar_runs, create_floating_run
from .validate import validate_tables


def read_tables(vars_file):
    """
    Reads the Files, Queries, Sources, Values, and Formats sheets into DataFrames.

    Parameters
    ----------
    vars_file : str
        The file path to the vars spreadsheet

    Returns
    -------
    files : pd.DataFrame
    # TODO finish this docstring im too lazy and there's too much to do
    """
    # TODO unicodedata.normalize

    try:
        files = pd.read_excel(vars_file, sheet_name='Files', index_col=0)
        queries = pd.read_excel(vars_file, sheet_name='Queries', index_col=0)
        sources = pd.read_excel(vars_file, sheet_name='Sources', index_col=0)
        values = pd.read_excel(vars_file, sheet_name='Values', index_col=0)
        formats = pd.read_excel(vars_file, sheet_name='Formats', index_col=0)

    except (FileNotFoundError, OSError) as e:
        logger.critical(f'Could not find the vars spreadsheet at "{vars_file}".')
    except xlrd.biffh.XLRDError as e:
        sheet = e.args[0][17:-2]
        logger.critical(f'The variables workbook does not contain the "{sheet}" sheet.')

    # Read formats sheet again, this time to get typography information
    workbook = xlrd.open_workbook(vars_file, formatting_info=True)
    sheet = workbook.sheet_by_name('Formats')
    for col in range(1, len(formats.columns) + 1, 2):
        for row in range(0, len(formats)):
            cell = sheet.cell(row + 1, col)
            if not cell.value:
                continue

            runs = []
            run_end = len(cell.value)

            # xlrd keeps the run information in a dictionary for every sheet (sheet.rich_text_runlist_map), meaning
            # each cell has a list of offsets. These offsets are how long the run is (in characters), i.e.
            # if a cell is 10 characters long and the first  half is bold, then the cell's entry in the dictionary
            # would be [5]. Then, xlrd stores the default run in another dictionary that covers the whole
            # workbook (workbook.xf_list).
            # For every cell in Formats, we get all runs in reverse order by offset, and then the default run,
            # and create a Docx run to insert into the template later.
            for offset, font_index in reversed(sheet.rich_text_runlist_map.get((row + 1, col)) or []):
                text = cell.value[offset:run_end]
                font = workbook.font_list[font_index]
                color = workbook.colour_map.get(font.colour_index) or (0, 0, 0)

                # We divide font.height by 20 because xlrd measures font size in Twips, which are 1/20th of a Point.
                # For example, an 18 pt. font is 360 twips.
                runs.append(create_floating_run(text=text, font_name=font.name, font_size=font.height / 20,
                                                bold=bool(font.bold), italic=bool(font.italic),
                                                underline=bool(font.underlined), color=color))
                run_end = offset

            # Get the remaining (default) run in the cell
            text = cell.value[:run_end]
            font = workbook.font_list[workbook.xf_list[cell.xf_index].font_index]
            color = workbook.colour_map.get(font.colour_index) or (0, 0, 0)  # If no color, leave it black ie. (0, 0, 0)

            runs.append(create_floating_run(text=text, font_name=font.name, font_size=font.height / 20,
                                            bold=bool(font.bold), italic=bool(font.italic),
                                            underline=bool(font.underlined), color=color))
            formats.iat[row, col - 1] = runs[::-1]


    # In case the user has any empty rows in any of the sheets, this will remove them.
    for df in [files, queries, sources, values, formats]:
        df.dropna(how='all', inplace=True)
        df.fillna('', inplace=True)
        # TODO .fillna('') seems like a bad idea with side effects I cannot see now but will regret later

    validate_tables(files, queries, sources, values, formats)
    return files, queries, sources, values, formats


def format_value(var, value, modifiers):
    """
    Gets the final formatted string for a variable, including rounding, converting to a percentage or currency,
    and so on.
    Parameters
    ----------
    var : str
        The name of the variable.
    value :
        The final value of the variable
    modifiers : str
        A string containing several format tokens, such as $ or %, which indicate how to format the variable.
    Returns
    -------
    formatted_value : str
        The final formatted value for the variable
    """
    if not modifiers:
        try:
            return str(value)
        except ValueError:
            logger.critical(f'Could not convert variable "{var}" into a string')
    elif type(value) not in {int, float, numpy.float16, numpy.float32, numpy.float64, numpy.floating}:
        logger.warning(f'Non-numeric variable "{var}" has formatting modifiers, which is not implemented yet. '
                       f'Ignoring modifiers')
        return str(value)

    original_value = value

    # Adjust value depending on modifiers
    if 'MK' in modifiers:
        if abs(value) >= 1e6:
            value = value / 1e6
        else:
            value = value / 1e3
    elif '%' in modifiers:
        value = value * 100

    # Determine precision
    if '.' not in modifiers:  # Default precisions when none is given
        if 'MK' in modifiers:
            precision = 1 if original_value > 1e6 else 0
        elif '%' in modifiers:
            precision = 1
        elif '$' in modifiers:
            precision = 2
    else:
        precision = modifiers[modifiers.find('.') + 1]
        if not precision.isnumeric():
            logger.warning(f'Invalid precision value in modifiers for var "{var}": "{precision}", '
                           f'it will be rounded to one decimal place by default')
            precision = 1
        precision = int(precision)

    # Construct formatted string and then fill it in. You can read the specification for them here:
    # https://docs.python.org/3.4/library/string.html#format-specification-mini-language
    format_string = '{0:'

    if ',' in modifiers:
        format_string += ','

    if '+-' in modifiers:
        format_string += '+'

    format_string += '.' + str(precision) + 'f}'
    formatted_value = format_string.format(round(value, precision))


    # Add final touches for certain modifiers
    if '$' in modifiers:
        if original_value < 0:
            formatted_value = '-$' + formatted_value[1:]
        elif '+-' in modifiers:
            formatted_value = '+$' + formatted_value[1:]
        else:
            formatted_value = '$' + formatted_value
    if '%' in modifiers:
        formatted_value += '%'
    if 'MK' in modifiers:
        if original_value > 1e6:
            formatted_value += 'M'
        else:
            formatted_value += 'K'

    return formatted_value


def order_vars(values):
    """
    Order the defined variables in a stack based on their dependent variables.
    The stack might contain duplicates, but if consumed in the correct (back to front) order,
    it will raise no dependency issues.

    For example, row 1 contains variable A which requires variable B and C. When we try to calculate A's value,
    and we don't have B or C's value, we will run into an error. This function rearranges the order in which
    we evaluate the values for every variable in order to avoid this problem.

    Parameters
    ----------
    values : pandas.DataFrame
        DataFrame containing the definitions for each variable. This is identical to the second
        table in the template file.

    Returns
    -------
    stack : list
        A stack containing the variables ordered (back to front) for evaluation.
    """
    stack = []
    visited = set()

    def visit(var):
        """
        Check through a variable's dependencies and add them (recursively) to the stack.
        This is essentially a Depth First Search to find a leaf node in a graph whose nodes
        are variables and whose edges are dependencies.

        Parameters
        ----------
        var : str
            The name of the variable that is being visited.
        """

        # If stack is growing out of control then we assume there is a circular import.
        # TODO Graph algorithms for cycle detection instead of this jank
        if len(stack) > len(values.index) * 50:
            raise RecursionError(f'Circular dependencies detected: {"->".join(stack[-15:])}...')

        visited.add(var)
        stack.append(var)
        text = ''.join([values.loc[var][col] for col in values.columns[2:]])
        for dependency in set(re.findall(r"`(.+?)`", text)):
            if dependency != var and dependency in values.index:  # If not in values, we assume it is defined in sources
                visit(dependency)

    for var in values.index:
        if var not in visited:
            visit(var)

    return stack


def parse_if_expression(variable, col, sheet, expression, variables):
    """
    Parses a a user-given If expression by substituting variables and evaluating.

    Parameters
    ----------
    variable : str
        The name of the variable who's Value/Format is being determined
    col : str
        The name of the column the If expression is in
    sheet : str
        The name of the sheet the If expression is in
    expression : str
        The expression to format
    variables : dict
        The current mapping of variables to values

    Returns
    -------
    bool
        True if the if expression evaluates to True, False otherwise
    """
    if type(expression) == list:  # TODO find why this happens sometimes
        print(f'List-like expression found for {variable}: \n', expression, '\n This should not happen?')
        expression = ''.join([run.text for run in expression])
    elif type(expression) != str:
        try:
            expression = str(expression)
        except ValueError:
            logger.critical(f'Unable to parse the "{col}" column on the {sheet} for variable "{variable}"')

    if not expression.strip():
        return False

    try:
        for var in re.findall(r"`(.+?)`", expression):
            expression = expression.replace(f'`{var}`', f"{variables[f'{var}']}")
    except KeyError as e:
        logger.critical(f'"{col}" column on the {sheet} sheet for variable "{variable}" references an '
                        f'undefined variable: "{e.args[0]}"')

    expression = expression.replace('”', '"')
    expression = expression.replace('“', '"')

    try:
        if eval(expression):
            return True
        else:
            return False
    except TypeError as e:
        logger.critical(f'"{col}" column on the {sheet} sheet for variable "{variable}" contains an invalid operation '
                        f'for two or more variables in the If expression (for example, trying to add a number and '
                        f'a string).')

    except KeyError as e:  # This might be overkill but just in case?
        logger.critical(f'"{col}" column on the {sheet} sheet for variable "{variable}" references an '
                        f'undefined variable: "{e.args[0]}"')
    except SyntaxError:
        logger.critical(f'Error in syntax for "{variable}" in column "{col}" on the {sheet} sheet')

    except Exception as e:
        tb = traceback.format_exc()
        logger.critical(f'An unexpected error occurred when evaluating the "{col}" column on the {sheet} sheet '
                        f'for the variable "{variable}": \n\n{tb}')


def parse_value_expression(variable, col, expression, variables, cast_type=None):
    """
    Parse a user-given Value expression by substituting variables, then evaluating the expression, and then
    casting the result to the specified type.

    Parameters
    ----------
    variable : str
        The name of the variable who's Value is being evaluated
    col : str
        The name of the column the Value expression is in
    expression : str
        The expression to format
    variables : dict
        The current mapping of variables to values
    cast_type : type, optional
        The final type the value of the variable will be cast to if specified, otherwise it will be interpreted

    Returns
    -------
    value
        The resulting value of the expression
    """
    if type(expression) != str:
        try:
            expression = str(expression)
        except ValueError:
            logger.critical(f'Unable to parse the "{col}" column on the Values sheet for variable "{variable}"')

    for var in re.findall(r"`(.+?)`", expression):
        expression = expression.replace(f'`{var}`', f"variables['{var}']")

    expression = expression.replace('”', '"')  # TODO unicode normalize?
    expression = expression.replace('“', '"')

    if not expression:
        logger.critical(f'Variable "{variable}" has an empty expression in column "{col}" on the Values sheet')

    else:
        try:
            value = eval(expression)
        except KeyError as e:
            logger.critical(f'"{col}" column on the Values sheet for variable "{variable}" references an undefined '
                            f'variable: "{e.args[0]}"')
        except SyntaxError:
            logger.critical(f'Error in syntax for "{variable}" in column "{col}" on the Values sheet')
        except TypeError:
            logger.critical(f'"{col}" column on the Values sheet for variable "{variable}" contains an invalid '
                            f'operation for two or more variables in the Value expression (for example, '
                            f'trying to add a number and a string).')
        except Exception as e:
            tb = traceback.format_exc()
            logger.critical(f'An unexpected error occurred when evaluating the "{col}" column on the Values sheet '
                            f'for the variable "{variable}": \n\n{tb}')

        return value


def parse_format_expression(variable, col, runs, variables):
    """
    Parses a user given format expression by substituting values into the constituent runs found in the
    Formats sheet.

    Parameters
    ----------
    variable : str
        The name of the variable the Format is being evaluated
    col : str
        The name of the column the Format expression is in
    runs : list
        The list of runs that make up the typography of the variable.
    variables : dict
        The dictionary containing all of the final values for every variable.

    Returns
    -------
    runs : list
        The list of runs with the variables values substituted.
    """
    runs = unite_similar_runs(runs)

    for run in runs:
        for var, modifiers in re.findall(r"`(.+?)`(?:\[(.+?)\])?", run.text):
            if var not in variables:
                logger.critical(f'"{col}" column on the Formats sheet for variable "{variable}" references an '
                                f'undefined variable: "{var}"')
            else:
                formatted_value = format_value(var, variables[var], modifiers)

            if modifiers:
                run.text = run.text.replace(f'`{var}`[{modifiers}]', formatted_value)
            else:
                run.text = run.text.replace(f'`{var}`', formatted_value)

    return runs

