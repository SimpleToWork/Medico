import os
import sys
currentdir = os.path.dirname(os.path.realpath(__file__))
parentdir = os.path.dirname(currentdir)
sys.path.append(parentdir)
import getpass
sys.path.append(f'C:\\Users\\{getpass.getuser()}\\Desktop\\New Projects\\Medico\\Medico\\venv\\Lib\\site-packages')
import json
import crayons
import numpy as np
import pandas as pd
from sqlalchemy import create_engine, inspect
import time
import sqlalchemy
import platform
import requests

import datetime

class objdict(dict):
    def __getattr__(self, name):
        if name in self:
            return self[name]
        else:
            raise AttributeError("No such attribute: " + name)

    def __setattr__(self, name, value):
        self[name] = value

    def __delattr__(self, name):
        if name in self:
            del self[name]
        else:
            raise AttributeError("No such attribute: " + name)

class ProgramCredentials:
    def __init__(self, environment):
        filename = __file__
        filename = filename.replace('/', "\\")
        folder_name = '\\'.join(filename.split('\\')[:-2])
        if environment == 'development':
            file_name = f'{folder_name}\\credentials_development.json'
        elif environment == 'production':
            file_name = f'{folder_name}\\credentials_production.json'

        f = json.load(open(file_name))

        self.project_folder = folder_name
        self.drive_credentials_file = f['drive_credentials_file'].replace("%USERNAME%", getpass.getuser())
        self.drive_token_file = f['drive_token_file'].replace("%USERNAME%", getpass.getuser())
        self.drive_scopes = f['drive_scopes']

        self.gsheet_credentials_file = f['gsheet_credentials_file'].replace("%USERNAME%", getpass.getuser())
        self.gsheet_token_file = f['gsheet_token_file'].replace("%USERNAME%", getpass.getuser())
        self.gsheet_scopes = f['gsheet_scopes']

        self.gmail_credentials_file = f['gmail_credentials_file'].replace("%USERNAME%", getpass.getuser())
        self.gmail_token_file = f['gmail_token_file'].replace("%USERNAME%", getpass.getuser())
        self.gmail_scopes = f['gmail_scopes']
        self.email_sender = f['email_sender']

        self.mle_folder = f['mle_folder']

        self.google_sheet_published = f['google_sheet_published']
        self.google_sheet_form_responses = f['google_sheet_form_responses']
        self.google_sheet_response_detail = f['google_sheet_response_detail']
        self.google_sheet_merge_process = f['google_sheet_merge_process']

        self.gmail_upload_folder_id = f['gmail_upload_folder_id']

        self.auto_publish_sheet_name = f['auto_publish_sheet_name']
        self.published_folder = f['published_folder']
        self.sub_published_folder = f['sub_published_folder']

        self.stw_gsheet_credentials_file = f['stw_gsheet_credentials_file'].replace("%USERNAME%", getpass.getuser())
        self.stw_gsheet_token_file = f['stw_gsheet_token_file'].replace("%USERNAME%", getpass.getuser())
        self.stw_gsheet_scopes = f['stw_gsheet_scopes']
        self.stw_gsheet_dashboard_id = f['stw_gsheet_dashboard_id']
        self.stw_gsheet_dashboard_sheet_name = f['stw_gsheet_dashboard_sheet_name']
        self.stw_gsheet_gid = f['stw_gsheet_gid']
        self.merge_gsheet_gid = f['merge_gsheet_gid']

        self.username = f['username']
        self.password = f['password']
        self.hostname = f['hostname']
        self.port = f['port']
        self.project_name = f['project_name']

        self.webhook_url = f['webhook_url']

    def set_attributes(self, params):

        params = objdict(params)
        for key, val in params.items():
            params[key] = objdict(val)

        return params


def print_color(*text, color='', _type='', output_file=None):
    ''' color_choices = ['r','g','b', 'y']
        _type = ['error','warning','success','sql','string','df','list']
    '''
    color = color.lower()
    _type = _type.lower()

    if color == "g" or _type == "success":
        crayon_color = crayons.green
    elif color == "r" or _type == "error":
        crayon_color = crayons.red
    elif color == "y" or _type in ("warning", "sql"):
        crayon_color = crayons.yellow
    elif color == "b" or _type in ("string", "list"):
        crayon_color = crayons.blue
    elif color == "p" or _type == "df":
        crayon_color = crayons.magenta
    elif color == "w":
        crayon_color = crayons.normal
    else:
        crayon_color = crayons.normal


    print(*map(crayon_color, text))
    if output_file is not None:
        # print(output_file)
        # print(os.path.exists(output_file))
        if os.path.exists(output_file) is False:
            # print("Right Here")
            file1 = open(output_file, 'w')
            file1.writelines(f'#################\n')
            file1.close()
            # file1 = open(output_file, 'w')
            # file1.close()
        # print(os.path.exists(output_file))
        file1 = open(output_file, 'a')
        file1.writelines(f'{str(text)}\n')
        file1.close()
        # print("Here")


def convert_dataframe_types(df=None):
    # print_color(df, color='y')
    columnLenghts = np.vectorize(len)

    # df = pd.DataFrame({'col': [1, 2, 10, np.nan, 'a'],
    #                    'col2': ['a', 10, 30, 40, 50],
    #                    'col3': [1, 2, 3, 4.36, np.nan]})

    col_is_numeric = df.replace(np.nan, 0).replace("nan", 0).replace("Nan",0).apply(lambda s: pd.to_numeric(s, errors='coerce')).notnull().all().tolist()
    col_list = df.columns.tolist()

    # print_color(col_is_numeric, color='g')
    # print_color(col_list, color='g')
    df_original_types = df.dtypes.tolist()
    for i, val in enumerate(col_is_numeric):
        if val == True:
            # print(df_original_types[i], col_list[i])
            if "datetime" not in str(df_original_types[i]):
                if "float" in str(df_original_types[i]):
                    # print( df[col_list[i]])
                    # print(df[col_list[i]].replace(np.nan, 0).replace("nan",0).astype(str).str.split(".", n=2, expand = True))
                    decimal_level = df[col_list[i]].replace(np.nan, 0).replace("nan",0).astype(str).str.split(".", n=2, expand = True)[1].unique().tolist()
                else:
                    decimal_level = ['0']
                if len(decimal_level) == 1 and decimal_level[0] == '0':
                    df[col_list[i]] = df[col_list[i]].replace(np.nan, 0)
                    df[col_list[i]] = pd.to_numeric(df[col_list[i]], errors='ignore', downcast='integer')
                else:
                    df[col_list[i]] = pd.to_numeric(df[col_list[i]], errors='ignore')

    return df


class create_folder():
    def __init__(self, foldername=""):
        if not os.path.exists(foldername):
            os.mkdir(foldername)


class Get_SQL_Types():
    def __init__(self,DataFrame):
        columnLenghts = np.vectorize(len)
        # print(np.nan)
        ## CONVERT DATAFRAME TYPES TO PROPER NUMERIC OR INTERGER BASED COLUMN TYPES
        # for col in DataFrame.columns:
        #     print_color(f'Attempting to Convert {col}')
        #     DataFrame[col] = DataFrame[col] .replace(np.nan, 0).replace("nan", 0).replace("Nan", 0).apply(lambda s: pd.to_numeric(s, errors='coerce'))

        # col_is_numeric = DataFrame.notnull().all().tolist()
        col_is_numeric =  DataFrame.replace(np.nan, 0).replace("nan", 0).replace("Nan", 0).apply(lambda s: pd.to_numeric(s, errors='coerce')).notnull().all().tolist()
        # print_color(f'col_is_numeric: {col_is_numeric}' , color='p')
        col_list = DataFrame.columns.tolist()
        df_original_types = DataFrame.dtypes.tolist()
        # print(df_original_types)
        for i, val in enumerate(col_is_numeric):
            if val == True:
                # print(df_original_types[i])
                if "datetime" not in str(df_original_types[i]) :
                    # print(col_list[i])
                    # print(DataFrame[col_list[i]].replace(np.nan, 0).astype(str).replace("nan", 0))
                    # if "." in DataFrame[col_list[i]].replace(np.nan, 0).astype(str).replace("nan", 0):
                    decimal_level = DataFrame[col_list[i]].replace(np.nan, 0).astype(str).replace("nan", 0).str.split(".", n=2, expand=True)
                    # else:
                    #     decimal_level = DataFrame[col_list[i]].replace(np.nan, 0).astype(str).replace("nan", 0)
                    # print(decimal_level)
                    if len(decimal_level.columns) > 1:
                        decimal_level = decimal_level[1].unique().tolist()
                        if len(decimal_level) == 1 and decimal_level[0] == '0':
                            DataFrame[col_list[i]] = DataFrame[col_list[i]].replace(np.nan, 0)
                            DataFrame[col_list[i]] = pd.to_numeric(DataFrame[col_list[i]], errors='ignore', downcast='integer')
                        else:
                            DataFrame[col_list[i]] = pd.to_numeric(DataFrame[col_list[i]], errors='ignore')
                    else:
                        DataFrame[col_list[i]] = DataFrame[col_list[i]].replace(np.nan, 0)
                        DataFrame[col_list[i]] = pd.to_numeric(DataFrame[col_list[i]], errors='ignore',downcast='integer')

        # print(DataFrame.dtypes)

        ## GET THE APPROPRIATE MYSQL COLUMN TYPES FOR THE DATAFRAME OBJECT
        data_types = dict()
        for col in DataFrame.columns:
            # print(col)
            if DataFrame.shape[0]>0:
                Col_Length = columnLenghts(DataFrame[col].values.astype(str)).max(axis=0)
            else:
                Col_Length = 5
            Col_Type = DataFrame[col].dtypes
            # print("column", col, Col_Length, Col_Type)
            if Col_Type == "object":
                if Col_Length > 255:
                    column_type = {col:sqlalchemy.types.TEXT()}
                    data_types.update(column_type)
                elif Col_Length >= 100:
                    column_type = {col:sqlalchemy.types.VARCHAR(255)}
                    data_types.update(column_type)
                elif  Col_Length >= 50:
                    column_type = {col:sqlalchemy.types.VARCHAR(100)}
                    data_types.update(column_type)
                elif  Col_Length >= 25:
                    column_type = {col:sqlalchemy.types.VARCHAR(50)}
                    data_types.update(column_type)
                elif Col_Length >= 15:
                    column_type = {col:sqlalchemy.types.VARCHAR(25)}
                    data_types.update(column_type)
                elif Col_Length >= 10:
                    column_type = {col:sqlalchemy.types.VARCHAR(15)}
                    data_types.update(column_type)
                elif Col_Length >= 5:
                    column_type = {col:sqlalchemy.types.VARCHAR(10)}
                    data_types.update(column_type)
                elif Col_Length >= 1:
                    column_type = {col:sqlalchemy.types.VARCHAR(5)}
                    data_types.update(column_type)
                elif  Col_Length == 0:
                    column_type = {col: sqlalchemy.types.VARCHAR(5)}
                    data_types.update(column_type)
            if Col_Type == "float" or Col_Type == "float64":
                new_data = DataFrame[col].to_frame()
                new_data = new_data.fillna(0)
                new_data[col] = new_data[col].astype(str)
                new = new_data[col].str.split(".", n = 1, expand = True)
                new.columns = ["First","Second"]
                Integer_Depth = columnLenghts(new['First'].values.astype(str)).max(axis=0)
                Decimal_Depth = columnLenghts(new['Second'].values.astype(str)).max(axis=0)
                # print_color(col, 'Integer_Depth', Integer_Depth, color='p')
                # print_color(col, 'Decimal_Depth',Decimal_Depth,color='p')
                if Decimal_Depth <=2:

                    if Col_Length <=10:
                        column_type = {col: sqlalchemy.types.Numeric(12,2)}
                        # column_type = {col: sqlalchemy.types.FLOAT(precision=12, asdecimal=True,decimal_return_scale=3)}
                    elif Col_Length <=20:
                        column_type = {col: sqlalchemy.types.Numeric(22, 2)}
                        # column_type = {col: sqlalchemy.types.FLOAT(20, 2)}
                    data_types.update(column_type)
                else:
                    if Col_Length <=10:
                        # column_type = {col: sqlalchemy.types.FLOAT(12,4)}
                        column_type = {col: sqlalchemy.types.Numeric(12, 4)}
                    elif Col_Length <=20:
                        # column_type = {col: sqlalchemy.types.FLOAT(20, 4)}
                        column_type = {col: sqlalchemy.types.Numeric(22, 4)}
                    data_types.update(column_type)
            if Col_Type == "int32" or Col_Type == "int64" or Col_Type == "int8" :
                if Col_Length >= 10:
                        column_type = {col: sqlalchemy.types.BIGINT()}
                else:
                    column_type = {col: sqlalchemy.types.INTEGER()}
                data_types.update(column_type)
            if "datetime64[ns]" in str(Col_Type) or "datetime64" in str(Col_Type):
                date_level = len(DataFrame[col].astype(str).str.split(" ", n=1, expand=True).columns)
                if date_level ==1:
                    column_type = {col: sqlalchemy.types.DATE()}
                    data_types.update(column_type)
                else:
                    column_type = {col: sqlalchemy.types.DATETIME()}
                    data_types.update(column_type)
            if Col_Type == "bool":
                column_type = {col: sqlalchemy.types.BOOLEAN()}
                data_types.update(column_type)

            # print("Column", col, Col_Type, column_type)

        self.data_types = data_types


class View_SQL_Column_Lengths():
    def __init__(self, engine='', Project_Name='', Table_Name=''):
        script = f'Select Ordinal_Position as "#", column_name AS `COLUMN`, upper(COLUMN_TYPE) as TYPE From information_schema.columns b1 where b1.table_schema = "{Project_Name}" And b1.table_name = "{Table_Name}" order by ORDINAL_POSITION;'
        df = pd.read_sql(script, con=engine)
        script1 =''
        count=0

        text_column_count = len(df[df['TYPE'] == 'TEXT']) -1

        for i in range(df.shape[0]):
            # print(i)
            column_name = df['COLUMN'].iloc[i]
            column_type = df['TYPE'].iloc[i]
            # print(column_name, column_type)

            # if column_type == "TEXT":

            if count == df.shape[0]-1:
                script1 = script1 + '\n(SELECT "' + column_name + '" as "Column_Name",Max(Length(`' + column_name + '`)) as Char_Length \n From ' \
                          + Table_Name + '\n limit 1);'
            else:
                script1 = script1 + '\n(SELECT "' + column_name + '" as "Column_Name",Max(Length(`' + column_name + '`)) as Char_Length \n From ' \
                          + Table_Name + '\n limit 1) \n UNION ALL'


                count += 1

        df1 = pd.DataFrame()

        if script1 !='':
            df1 = pd.read_sql(script1, con=engine)
            # print(df1)

        self.DataFrame = df1


class Change_Sql_Column_Types():
    def __init__(self, engine='', Project_name='', Table_Name='', DataTypes='', DataFrame=''):

        # df3 = View_SQL_Column_Lengths(engine=engine, Project_Name=Project_name, Table_Name=Table_Name).DataFrame
        script = f'Select Ordinal_Position as "#", column_name AS `COLUMN`, upper(COLUMN_TYPE) as TYPE From information_schema.columns b1 where b1.table_schema = "{Project_name}" And b1.table_name = "{Table_Name}" order by ORDINAL_POSITION;'

        df2 = pd.read_sql(script, con=engine)
        df = DataFrame
        modify_script = ""
        DataType = {k.upper(): v for k,v in DataTypes.items()}
        # print(script)
        for i in range(df2.shape[0]):
            column = str(df2['COLUMN'].iloc[i])
            comparable_column = str(df2['COLUMN'].iloc[i]).upper()
            Column_Type = str(df2['TYPE'].iloc[i]).replace("'", '').replace('b', '')                    #THIS IS THE MYSQL COLUMN TYPE
            # sql_column_length = df3.loc[df3['Column_Name'] == column]['Char_Length'].iloc[0]
            # print_color(f'{column}, {Column_Type}, {DataType.keys()}', color='p')

            if comparable_column in DataType:
                dataframe_column_type = str(DataType[comparable_column]).replace(" ", "")                          # THIS IS THE DATAFRAME COLUMN TYPE
                # dataframe_column_type = dataframe_column_type

                # print_color(column, Column_Type, dataframe_column_type, color='p')
                if Column_Type == "FLOAT(12,4)" or Column_Type == "FLOAT(12,2)" or Column_Type == "FLOAT(20,4)" or Column_Type == "FLOAT(20,2)" or Column_Type == "VARCHAR(5)":
                    df[column] = df[column].replace(np.nan, 0)

                if Column_Type != dataframe_column_type:
                    if (Column_Type == "INT(11)" or Column_Type == "INT") and dataframe_column_type == "INTEGER":
                        pass
                    elif Column_Type == "BIGINT(20)" and dataframe_column_type == "BIGINT":
                        pass
                    elif (Column_Type == "BIGINT(20)" or Column_Type == "BIGINT") and dataframe_column_type == "INTEGER":
                        pass
                    elif Column_Type == "DATETIME" and "VARCHAR" in dataframe_column_type:
                        pass
                    elif "VARCHAR" in Column_Type and "DATETIME" in dataframe_column_type:
                        pass
                    elif Column_Type == "TIMESTAMP" and "DATETIME"  in dataframe_column_type:
                        pass
                    elif Column_Type == "DATE" and "VARCHAR" in dataframe_column_type:
                        pass
                    elif (Column_Type == "FLOAT(12,4)" or Column_Type == "FLOAT(20,4)" or Column_Type == "FLOAT(12,2)" or Column_Type == "FLOAT(20,2)") and dataframe_column_type == "INTEGER":
                        pass
                    elif Column_Type == "TINYINT(1)" and dataframe_column_type == "BOOLEAN":
                        pass
                    elif "FLOAT(20,4)" in Column_Type and "FLOAT(12,4)" in dataframe_column_type:
                        pass
                    elif "FLOAT(20,2)" in Column_Type and "FLOAT(12,2)" in dataframe_column_type:
                        pass
                    elif "DECIMAL(20,4)" in Column_Type and "DECIMAL(12,4)"in dataframe_column_type or "NUMERIC(12,4)" in dataframe_column_type:
                        pass
                    elif "DECIMAL(20,2)" in Column_Type and "DECIMAL(12,2)" in dataframe_column_type or "NUMERIC(12,2)" in dataframe_column_type:
                        pass
                    elif "NUMERIC(20,4)" in Column_Type and "NUMERIC(12,4)" in dataframe_column_type or "DECIMAL(12,4)"in dataframe_column_type:
                        pass
                    elif "NUMERIC(20,2)" in Column_Type and "NUMERIC(12,2)" in dataframe_column_type or "DECIMAL(12,2)" in dataframe_column_type:
                        pass

                    elif "VARCHAR" in Column_Type and ("NUMERIC" in dataframe_column_type or "DECIMAL" in dataframe_column_type or "FLOAT" in dataframe_column_type
                                                       or "INTEGER" in dataframe_column_type or "BIGINT" in dataframe_column_type):
                        pass
                    elif Column_Type == "DATE" and "DATETIME" in dataframe_column_type:
                        pass
                    elif "VARCHAR" in Column_Type and "BOOLEAN" in dataframe_column_type:
                        pass
                    elif "TEXT" in Column_Type and "VARCHAR" in dataframe_column_type:
                        pass
                    elif "TEXT" in Column_Type and "FLOAT" in dataframe_column_type:
                        pass
                    elif "VARCHAR" in Column_Type and "VARCHAR" in dataframe_column_type:
                        database_column_length = int(Column_Type.split("(")[1].split(")")[0])
                        dataframe_column_length = int(dataframe_column_type.split("(")[1].split(")")[0])

                        if dataframe_column_length > database_column_length:
                            # print_color(column, Column_Type, dataframe_column_type, color='y')
                            # print_color(database_column_length, dataframe_column_length, color='y')
                            # print(column, Column_Type, dataframe_column_length, database_column_length, dataframe_column_type,dataframe_column_length)
                            if modify_script == "":
                                modify_script += "MODIFY COLUMN `" + column + "` " + "VARCHAR(" + str(dataframe_column_length) + ")"
                            else:
                                modify_script += ", \nMODIFY COLUMN `" + column + "` " + "VARCHAR(" + str(dataframe_column_length) + ")"
                    elif "INT" in Column_Type and "DATETIME" in dataframe_column_type:
                        if modify_script == "":
                            modify_script += "MODIFY COLUMN `" + column + "` " + "VARCHAR(25)"
                        else:
                            modify_script += ", \nMODIFY COLUMN `" + column + "` " + "VARCHAR(25)"

                    else:
                        check_values = df[column].unique()
                        # print(check_values)
                        if len(check_values) == 1 and (check_values[0] == 0 or str(check_values[0]) == 'nan'):
                            pass
                        else:
                            # print_color(column, Column_Type, dataframe_column_type, color='y')
                            if modify_script == "":
                                modify_script += "MODIFY COLUMN `" + column + "` " + dataframe_column_type
                            else:
                                modify_script += ", \nMODIFY COLUMN `" + column + "` " + dataframe_column_type
        scripts=[]

        alter_script = "ALTER TABLE `" + Table_Name + '`\n'
        if modify_script != "":
            scripts.append(f'SET FOREIGN_KEY_CHECKS=0;')
            scripts.append(alter_script + modify_script)
            scripts.append(f'SET FOREIGN_KEY_CHECKS=1;')
            run_sql_scripts(engine=engine, scripts=scripts)

        self.DataFrame = DataFrame


def engine_setup(project_name=None, hostname = None, username=None, password=None, port=None, pool_pre_ping=True, echo=False):
    if project_name is None:
        engine = create_engine(f'mysql+mysqlconnector://{username}:{password}@{hostname}:{port}',pool_pre_ping=pool_pre_ping, echo=echo)
    else:
        engine = create_engine(f'mysql+mysqlconnector://{username}:{password}@{hostname}:{port}/{project_name}?charset=utf8',pool_pre_ping=pool_pre_ping, echo=echo)
    return engine


def log_sql_scripts(log_scripts=False, log_engine=None, log_database=None,  script_name=None, profile_name=None,
                    project_name=None, company_name=None,
                    query=None, start_time=None, end_time=None, duration=None):

    if log_scripts is True:
        t = [(profile_name, project_name, company_name, script_name, query, start_time, end_time, duration)]
        df = pd.DataFrame(t)

        df.columns = ['profile_name', 'project_name', 'company_name', 'script_name', 'sql_query', 'start_time', 'end_time', 'duration']
        print(df)

        table = 'sql_runtime_scripts'
        sql_types = Get_SQL_Types(df).data_types
        Change_Sql_Column_Types(engine=log_engine, Project_name=log_database, Table_Name=table,
                                                 DataTypes=sql_types, DataFrame=df)
        df.to_sql(name=table, con=log_engine, if_exists='append', index=False, schema=log_database, chunksize=1000,dtype=sql_types)


    # pass


def run_sql_scripts(engine=None, scripts=None, tryexcept=False,
                    log_scripts=False, log_engine=None, log_database=None, script_name=None, profile_name=None,
                    project_name=None, company_name=None):
    real_start_time = time.time()
    time_list = []
    if tryexcept is True:
        for script in scripts:
            run_method = True
            run_attempt = 0
            time_now = time.time()
            start_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            while run_method is True and run_attempt < 5:
                print_color(script, color='y')
                print_color("Running", color='p')
                try:
                    engine.execute(script)
                    time_list.append(time.time() - time_now)
                    log_sql_scripts(log_scripts=log_scripts, log_engine=log_engine, log_database=log_database, script_name=script_name,
                                    profile_name=profile_name, project_name=project_name, company_name=company_name,
                                    query=script, start_time=start_time,
                                    end_time=datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                    duration=time.time() - time_now)

                    print_color(f'Script Complete -- Took {round(time.time() - time_now, 4)} seconds to Run --', color='p')
                    run_method = False
                except Exception as e:
                    print_color(str(e), color='r')
                    if "(mysql.connector.errors.DatabaseError) 1206 (HY000): The total number of locks exceeds the lock table size" in str(e)\
                            or "(mysql.connector.errors.DatabaseError) 1205 (HY000): Lock wait timeout exceeded;" in str(e)\
                            or "(mysql.connector.errors.InternalError) 1213 (40001): Deadlock found when trying to get lock;" in str(e):

                        print_color("Going To Handle Here", color='r')
                        for t in range(0, 60, 10):
                            print_color(f"Mysql Table is Locked. Waiting {60 - t} Seconds to run.", color='y')
                            time.sleep(10)
                        run_attempt +=1

                    elif "(mysql.connector.errors.InterfaceError) 2013: Lost connection to MySQL server during query" in str(e):
                        print_color("Going To Handle Here", color='r')
                        for t in range(0, 15, 5):
                            print_color(f"Mysql Lost connection;  Waiting {15 - t} Seconds to run.", color='y')
                            time.sleep(5)
                        run_attempt += 1

                    else:
                        raise ValueError('Script Broke in Runtime')
                        run_method = False
            else:
                if run_attempt >=5:
                    print_color("Number of Tries Exceeded Attempt Threshold \n Forcing Break.", color='r')
                if run_method is False:
                    print_color("Method Complete", color='g')

    else:
        for script in scripts:
            time_now = time.time()
            start_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            print_color(script, color='y')

            engine.execute(script)

            time_list.append(time.time() - time_now)
            log_sql_scripts(log_scripts=log_scripts, log_engine=log_engine, log_database=log_database,
                            script_name=script_name,
                            profile_name=profile_name, project_name=project_name, company_name=company_name,
                            query=script, start_time=start_time,
                            end_time=datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            duration=time.time() - time_now)
            print_color(f'Script Complete -- Took {round(time.time() - time_now,4)} seconds to Run --', color='p')
    print_color(f'Scripts Complete --All Scripts Took {time.time() - real_start_time} seconds to Run --', color='b')


def record_program_performance(x, program_name, method):
    ip = requests.get('https://api.ipify.org').content.decode('utf8')

    database_name = "stw_task_manager"

    computer_name = platform.node()
    user = getpass.getuser()
    time_now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print_color(f'Data imported', color='g')

    url = x.webhook_url
    run_task_webhook = f'{url}/method_performance'
    print_color(f'Attempting to Hit: {run_task_webhook}')

    headers = {'Content-Type': 'application/json'}

    data = {
        'DateTime':time_now,
        'Computer':computer_name,
        'User':user,
        'Program Name':program_name,
        'Function':method,
        'Success':True
    }
    print_color(data, color='r')

    data = {"ip": ip, "data": data}

    response = requests.post(url=run_task_webhook, headers=headers, json=json.dumps(data))

    print("Request URL:", response.request.url)
    print("Request method:", response.request.method)
    print("Request headers:", response.request.headers)
    print("Request body:", response.request.body)

    print_color(f"Request status: {response.status_code}", color='g')
    print_color(f"Request content: {response.content}", color='y')