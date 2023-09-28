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

# class ProgramCredentials:
#     def __init__(self, environment):
#         filename = __file__
#         filename = filename.replace('/', "\\")
#         folder_name = '\\'.join(filename.split('\\')[:-2])
#         if environment == 'development':
#             file_name = f'{folder_name}\\credentials_development.json'
#         elif environment == 'production_kyf':
#             file_name = f'{folder_name}\\credentials_production_kyf.json'
#         elif environment == 'production_fc':
#             file_name = f'{folder_name}\\credentials_production_fc.json'
#         elif environment == 'production_gn':
#             file_name = f'{folder_name}\\credentials_production_gn.json'
#
#         f = json.load(open(file_name))
#
#         self.project_folder = folder_name
#         self.qb_hostname = f['qb_hostname']
#         self.qb_auth = f['qb_auth']
#         self.qb_app_id = f['qb_app_id']
#         self.qb_app_token = f['qb_app_token']
#         self.qb_user_token = f['qb_user_token']
#         self.deal_table_id = f['deal_table_id']
#         self.payment_table_id = f['payment_table_id']
#         self.customer_table_id = f['customer_table_id']
#         self.submissions_table_id = f['submissions_table_id']
#         self.payback_batch_table_id = f['payback_batch_table_id']
#         self.subsmission_log_table_id = f['subsmission_log_table_id']
#         self.username = f['username']
#         self.password = f['password']
#         self.achworks_username = f['achworks_username']
#         self.achworks_password = f['achworks_password']
#         self.achworks_loc_id = f['achworks_loc_id']
#         self.achworks_sss = f['achworks_sss']
#         self.achworks_wsdl = f['achworks_wsdl']
#
#         self.admin_name = f['admin_name']
#         self.app_name  = f['app_name']
#         self.outbound_email = f['outbound_email']
#         self.inbound_email = f['inbound_email']
#
#         self.change_payments = self.set_attributes(f['change_payments'])
#         self.add_broken_deal = self.set_attributes(f['add_broken_deal'])
#         self.add_deal_payments =  self.set_attributes(f['add_deal_payments'])
#
#         self.add_all_deal_payments = self.set_attributes(f['add_all_deal_payments'])
#         self.get_payment_data =self.set_attributes(f['get_payment_data'])
#
#         self.send_ach_payments = self.set_attributes(f['send_ach_payments'])
#         self.get_ach_status_data = self.set_attributes(f['get_ach_status_data'])
#
#
#     def set_attributes(self, params):
#
#         params = objdict(params)
#         for key, val in params.items():
#             params[key] = objdict(val)
#
#         return params
#



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

