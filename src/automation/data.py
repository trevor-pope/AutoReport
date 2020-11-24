"""
Data reading related tasks, including reading Excel sheets and SQL Server databases
"""

import logging
logger = logging.getLogger('email-automation')

import urllib
import pandas as pd
import xlrd
import pendulum
import sqlalchemy
import pyodbc
import traceback


def get_filepattern_fillins():
    """
    This will build all of the fill-ins that can be used when creating file patterns.
    I.e. these are the possible variables that can be used inside braces when specifying file names.

    If you'd like to add more fill-ins, please add entries to the fillins dictionary below.
    For example, if you wanted to add the date of the current week's Monday, you could do:

    fillins['monday'] = str(pendulum.today().last(pendulum.MONDAY))
    
    Returns
    -------
    fillins: dict
        A dict containing the possible variables that can be used in a file pattern
    """
    fillins = {}

    for prefix in ['', 'tomorrow', 'yesterday', 'weekending']:
        if prefix == '':
            base = pendulum.today()
        if prefix == 'tomorrow':
            base = pendulum.tomorrow()
        elif prefix == 'yesterday':
            base = pendulum.yesterday()
        elif prefix == 'weekending':
            base = pendulum.today().next(pendulum.SATURDAY)

        day = str(base.day)
        if len(day) == 1:
            day = '0' + day
        if prefix == '':
            fillins['day'] = day
        else:
            fillins[prefix] = day

        if prefix == '':
            prefix = 'today'
        fillins[prefix + 'nameupper'] = base.strftime('%A')
        fillins[prefix + 'namelower'] = base.strftime('%A').lower()
        fillins[prefix + 'nametruncupper'] = base.strftime('%a')
        fillins[prefix + 'nametrunclower'] = base.strftime('%a').lower()
        if prefix == 'today':
            prefix = ''

        month = str(base.month)
        if len(month) == 1:
            month = '0' + month
        fillins[prefix + 'month'] = str(month)

        year = str(base.year)
        fillins[prefix + 'year'] = year
        fillins[prefix + 'truncyear'] = year[-2:]

    return fillins


def connect(server, db):
    """
    Connects to a given database on the server. A user must have an ODBC Driver installed to connect to SQL Server
    databases. This logs in based on Active Directory (i.e. Windows Authentication), so the user currently logged in
    must have permissions to access the database.
    Parameters
    ----------
    server : str
        The MSSQL server that is being connected to
    db : str
        The database that is on the given server being connected to

    Returns
    -------
    con : sqlalchemy.engine.Engine
        A SQLAlchemy engine that contains a connection to the given database and server.
    """
    logger.info(f'Connecting to {db} on {server} using Windows Authentication.')
    if 'ODBC Driver 17 for SQL Server' in pyodbc.drivers():
        params = urllib.parse.quote_plus(r'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + r';DATABASE=' + db + r';Trusted_Connection=yes')
        fast_executemany = True
    elif 'ODBC Driver 13 for SQL Server' in pyodbc.drivers():
        params = urllib.parse.quote_plus(r'DRIVER={ODBC Driver 13 for SQL Server};SERVER=' + server + r';DATABASE=' + db + r';Trusted_Connection=yes')
        fast_executemany = False
    else:
        logger.critical(f'You do not have a valid SQL Server ODBC Driver installed. \n'
                        f'Please install the x64 version of ODBC Driver 17 for SQL Server at the following link: \n' 
                        f'https://docs.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server \n'
                        f'(You can copy by clicking on this message and pressing Ctrl+C and paste with Ctrl+V)')
    try:
        con = sqlalchemy.create_engine('mssql+pyodbc:///?odbc_connect={}'.format(params), fast_executemany=fast_executemany)
        return con
    except sqlalchemy.exc.InterfaceError:
        logger.critical(f'Could not connect to database "{db}" on server "{server}". '
                        f'Perhaps the user currently logged in lacks permission or the database/server does not exist.')


def read_data(files, queries, sources):
    """
    Grabs data necessary to assign initial variable values.

    Parameters
    ----------
    files : pd.DataFrame
        Files DataFrame that contains file paths to be read.
    queries : pd.DataFrame
        Queries DataFrame that contains SQL database and queries to be ran.
    sources : pd.DataFrame
        Sources DataFrame that list the initial value for variables (either from Files or Queries).

    Returns
    -------
    variables : dict
        Dictionary mapping variable names to their initial values.
    """

    variables = {}
    read_files = {}
    completed_queries = {}
    connections = {}
    filepattern_fillins = get_filepattern_fillins()

    for var in sources.index:
        if sources.loc[var]['File'] and sources.loc[var]['Query']:
            logger.warning(f'Both File and Query are listed as a source for variable "{var}". '
                           f'Defaulting to reading from the specified File')

        if sources.loc[var]['File']:
            try:
                file_pattern = files.loc[sources.loc[var]['File']]['Pattern']
            except KeyError as e:
                logger.critical(f'Sources sheet references File that does not exist: "{e.args[0]}"')
            try:
                file_pattern = file_pattern.format(**filepattern_fillins)
            except ValueError as e:
                logger.critical(f'Invalid fill-in value for "{sources.loc[var]["File"]}": "{{{e.args[0]}}}".')

            if file_pattern.endswith('.xlsx'):
                worksheet = sources.loc[var]['Worksheet']
                if not (file_pattern, worksheet) in read_files:
                    logger.info(f'Reading file "{file_pattern}" on worksheet "{worksheet}"')
                    try:
                        file = pd.read_excel(file_pattern, sheet_name=sources.loc[var]['Worksheet'], header=None)
                        read_files[file_pattern, worksheet] = file
                    except (FileNotFoundError, OSError) as e:
                        logger.critical(f'Could not find file "{file_pattern}" for variable "{var}"')
                    except (ValueError, xlrd.biffh.XLRDError):
                        logger.critical(f'Worksheet "{worksheet}" does not exist in File "{file_pattern}"')

                file = read_files[file_pattern, worksheet]
                try:
                    variables[var] = file.iat[convert_cell_to_index(sources.loc[var]['Cell'])]

                except IndexError:
                    logger.critical(f'Cell "{sources.loc[var]["Cell"]}" in file "{file_pattern}" contains no data')


            elif file_pattern.endswith('.csv'):  # TODO TEST, THIS PROBABLY DOESNT WORK AT ALL
                if not (file_pattern, None) in read_files:
                    logger.info(f'Reading file "{file_pattern}"')
                    try:
                        file = pd.read_csv(file_pattern, index_col=False)
                    except (FileNotFoundError, OSError):
                        logger.critical(f'Could not find file "{file_pattern}" for variable "{var}"')

                    read_files[file_pattern, sources.loc[var]['Worksheet']] = file

                file = read_files[file_pattern, sources.loc[var['Worksheet']]]
                row, col = convert_cell_to_index(sources.loc[var]['Cell'])
                try:
                    variables[var] = file.iat[row, col]  # assuming they don't count column headers as row numbers
                except IndexError:
                    logger.critical(f'Cell "{sources.loc[var]["Cell"]}" in file "{file_pattern}" contains no data')

            else:
                logger.critical(f'File "{file_pattern}" is of invalid type and cannot be read.')

        else:
            query = sources.loc[var]['Query']
            if not query:
                logger.critical(f'Variable "{var}" does not have a file or query listed in the Sources sheet')

            if query not in completed_queries:
                try:
                    server, db, sql = queries.loc[query]['Server'], queries.loc[query]['Database'], queries.loc[query]['SQL']
                except KeyError:
                    logger.critical(f'Variable "{var}" on the Sources sheet references a Query "{query}" which does not exist.')

                if (server, db) not in connections:
                    con = connect(server, db)
                    connections[server, db] = con
                else:
                    con = connections[server, db]

                try:
                    completed_queries[query] = pd.read_sql(sql=sql, con=con)
                except (sqlalchemy.exc.ProgrammingError, sqlalchemy.exc.DataError) as e:
                    logger.critical(f'An error occurred when running the SQL command for Query "{query}": \n'
                                    f'{e.args}')
                except Exception as e:
                    logger.critical(f'An unexpected error occurred when running the SQL command for Query "{query}": \n'
                                    f'{e.args}\n '
                                    f'{traceback.format_exc()}')

            row, col = sources.loc[var]['Row'], sources.loc[var]['Col']
            if not row and not col:
                row, col = 1, 1
            elif row < 1 or col < 1:
                logger.critical(f'Invalid entry for row and column for query "{query}": {row}, {col}')

            try:
                variables[var] = completed_queries[query].iat[int(row) - 1, int(col) - 1]
            except ValueError:
                try:
                    variables[var] = completed_queries[query][col].iloc[int(row) - 1]
                except ValueError:
                    logger.critical(f'Row and column {row}, {col} do not exist for query "{query}"')

    return variables


def convert_cell_to_index(cell):
    """
    Converts an Excel cell location to a zero-based row and col index.

    Parameters
    ----------
    cell : str
        The Excel cell location. For example, B17 or AF86

    Returns
    -------
    row : int
        The converted row num
    col : int
        The converted col num
    """

    split = None
    for i, c in enumerate(cell):
        if c.isnumeric():
            split = i
            break

    col = 0
    for i, c in enumerate(cell[:split]):
        col += (ord(c) - 64) + (26 * i) - 1

    row = int(cell[split:]) - 1  # -1 for 0 based index

    return row, col


