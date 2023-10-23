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



        self.google_sheet_published = f['google_sheet_published']
        self.google_sheet_form_responses = f['google_sheet_form_responses']
        self.google_sheet_response_detail = f['google_sheet_response_detail']

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

