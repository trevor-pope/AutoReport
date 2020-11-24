import logging
logger = logging.getLogger('email-automation')


def validate_tables(files, queries, sources, values, formats):
    validate_files(files)
    validate_queries(queries)
    validate_sources(sources)
    validate_values(values)
    validate_formats(formats)


def validate_files(df):
    if list(df.reset_index().columns) != ['File', 'Pattern']:
        logger.critical(f'Files sheet in the variables workbook has incorrect columns: \n'
                        f'\t[{", ".join(list(df.reset_index().columns))}]\n'
                        f'They should be: \n'
                        f'\t[File, Pattern]')

    if list(df.index[df.index.duplicated()]):
        logger.critical(f'Files sheet in the variables workbook contains the following duplicate entries: \n'
                        f'\t[{", ".join(list(df.index[df.index.duplicated()]))}]\n'
                        f'Please remove them.')


def validate_queries(df):
    if list(df.reset_index().columns) != ['Query', 'Server', 'Database', 'SQL']:
        logger.critical(f'Queries sheet in the variables workbook has incorrect columns: \n'
                        f'\t[{", ".join(list(df.reset_index().columns))}]\n'
                        f'They should be: \n'
                        f'\t[Query, Server, Database, SQL]')

    if list(df.index[df.index.duplicated()]):
        logger.critical(f'Queries sheet in the variables workbook contains the following duplicate entries: \n'
                        f'\t[{", ".join(list(df.index[df.index.duplicated()]))}]\n'
                        f'Please remove them.')


def validate_sources(df):
    if list(df.reset_index().columns) != ['VarName', 'File', 'Worksheet', 'Cell', 'Query', 'Row', 'Col']:
        logger.critical(f'Sources sheet in the variables workbook has incorrect columns: \n'
                        f'\t[{", ".join(list(df.reset_index().columns))}]\n'
                        f'They should be: \n'
                        f'\t[VarName, File, Worksheet, Cell, Query, Row, Col]')

    if list(df.index[df.index.duplicated()]):
        logger.critical(f'Sources sheet in the variables workbook contains the following duplicate entries: \n'
                        f'\t[{", ".join(list(df.index[df.index.duplicated()]))}]\n'
                        f'Please remove them.')

def validate_values(df):
    cols = df.reset_index().columns
    if len(cols) > len(set(cols)):
        logger.critical(f'Values sheet in the variables workbook contains duplicate column names.')

    has_value_else = False
    for i, col in enumerate(cols):
        if i == 0:
            if col != 'VarName':
                logger.critical(f'Values sheet in the variables workbook has invalid first column: "{col}". It should be "VarName"')
        elif i == 1:
            if col != 'Value1':
                logger.critical(f'Values sheet in the variables workbook has invalid third column: "{col}". It should be "Value1"')
        elif i == 2:
            if col != 'If1':
                logger.critical(f'Values sheet in the variables workbook has invalid third column: "{col}". It should be "If1"')
        elif col.startswith('Value') and col[5:].isnumeric():
            if not cols[i+1] == 'If' + col[5:]:
                logger.critical(f'Values sheet in the variables workbook has a Value column that does not '
                                f'precede a corresponding If column: {col}')
        elif col.startswith('If') and col[2:].isnumeric():
            continue
        elif col == 'ValueElse':
            has_value_else = True
            if i != len(cols) - 1:
                logger.critical(f'Values sheet in the variables workbook has a ValueElse column, '
                                f'but it is not the last column in the sheet.')
        else:
            logger.critical(f'Values sheet in the variables workbook has an invalid column: "{col}"')

    if not has_value_else:
        logger.critical(f'Values sheet in the variables workbook does not contain a ValueElse column.')

    if list(df.index[df.index.duplicated()]):
        logger.critical(f'Values sheet in the variables workbook contains the following duplicate entries: \n'
                        f'\t[{", ".join(list(df.index[df.index.duplicated()]))}]\n'
                        f'Please remove them.')


def validate_formats(df):
    cols = df.reset_index().columns
    if len(cols) > len(set(cols)):
        logger.critical(f'Formats sheet in the variables workbook contains duplicate column names.')

    has_format_else = False
    for i, col in enumerate(cols):
        if i == 0:
            if col != 'VarName':
                logger.critical(f'Formats sheet in the variables workbook has invalid first column: "{col}". It should be "VarName"')
        elif i == 1:
            if col != 'Format1':
                logger.critical(f'Formats sheet in the variables workbook has invalid third column: "{col}". It should be "Format1"')
        elif i == 2:
            if col != 'If1':
                logger.critical(
                    f'Formats sheet in the variables workbook has invalid third column: "{col}". It should be "If1"')
        elif col.startswith('Format') and col[6:].isnumeric():
            if not cols[i + 1] == 'If' + col[6:]:
                logger.critical(f'Formats sheet in the variables workbook has a Format column that does not '
                                f'precede a corresponding If column: {col}')
        elif col.startswith('If') and col[2:].isnumeric():
            continue
        elif col == 'FormatElse':
            has_format_else = True
            if i != len(cols) - 1:
                logger.critical(f'Formats sheet in the variables workbook has a FormatElse column,'
                                f' but it is not the last column in the sheet.')
        else:
            logger.critical(f'Formats sheet in the variables workbook has an invalid column: "{col}"')
    if not has_format_else:
        logger.critical(f'Formats sheet in the variables workbook does not contain a FormatElse column.')

    if list(df.index[df.index.duplicated()]):
        logger.critical(f'Formats sheet in the variables workbook contains the following duplicate entries: \n'
                        f'\t[{", ".join(list(df.index[df.index.duplicated()]))}]\n'
                        f'Please remove them.')
