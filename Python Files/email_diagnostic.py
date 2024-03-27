import datetime
import pandas as pd
from gmail_api import GoogleGmailAPI
from google_sheets_api import GoogleSheetsAPI
from global_modules import print_color, create_folder, error_handler, engine_setup


def get_email_process_history(GsheetAPI):
    date_time = datetime.datetime.combine(datetime.datetime.now(), datetime.time.min) - datetime.timedelta(days=1)
    print_color(date_time, color='g')
    df = GsheetAPI.get_data_from_sheet(sheetname="Auto Publish Data", range_name="A:Q")
    df['import_date'] = pd.to_datetime(df['import_date'])
    df = df[df['import_date']>=date_time]
    df = df.sort_values(by=['document_emailed', 'import_date', 'document_name'])

    print_color(df, color='r')
    print_color(df.columns, color='r')

    email_body = f'<br><span style="color:Black;font-weight:Bold; font-size:24px;">Auto Publish Process:</span>'
    email_body += f'<br><span style="color:Black; ">See Below Satus of Records Processed in the last 24 hours.</span>'
    email_body += f'<br><br><table>'
    email_body += f'''
        <tr>
            <th style="border: solid; border-width:1px; width: 100px; padding:0;">Import Date</td>
            <th style="border: solid; border-width:1px; width: 500px; padding:0;">Document Name</td>
            <th style="border: solid; border-width:1px; width: 100px; padding:0;">Date of Report</td>
            <th style="border: solid; border-width:1px; width: 100px; padding:0;">All Fields Assigned</td>
            <th style="border: solid; border-width:1px; width: 100px; padding:0;">Approved to Send Out</td>
            <th style="border: solid; border-width:1px; width: 100px; padding:0;">Document Emailed</td>
            <th style="border: solid; border-width:1px; width: 100px; padding:0;">Document Moved</td>
        </tr>'''

    for k in range(df.shape[0]):
        import_date = df['import_date'].iloc[k].strftime('%Y-%m-%d')
        document_name = df['document_name'].iloc[k]
        date_of_report = df['date_of_report'].iloc[k]
        all_fields_assigned = df['all_fields_assigned'].iloc[k]
        approved_to_send_out = df['approved_to_send_out_?'].iloc[k]
        document_emailed = df['document_emailed'].iloc[k]
        document_moved = df['document_moved'].iloc[k]

        print_color('document_name', document_name, color='y')

        all_fields_assigned_status = 'Green' if all_fields_assigned == 'TRUE' else 'Red'
        approved_to_send_out_status = 'Green' if approved_to_send_out == 'TRUE' else 'Red'
        document_emailed_status = 'Green' if document_emailed == 'TRUE' else 'Red'
        document_moved_status = 'Green' if document_moved == 'TRUE' else 'Red'

        email_body += f'''
            <tr>
                <td style="border: solid 1px black; width: 100px; padding:0; text-indent: 5px; ">{import_date}</td>
                <td style="border: solid 1px black; width: 500px; padding:0; text-indent: 5px;">{document_name}</td>
                <td style="border: solid 1px black; width: 100px; padding:0; text-indent: 5px;">{date_of_report}</td>
                <td style="border: solid 1px black; width: 100px; padding:0; text-indent: 5px; color:{all_fields_assigned_status}">{all_fields_assigned}</td>
                <td style="border: solid 1px black; width: 100px; padding:0; text-indent: 5px; color:{approved_to_send_out_status}">{approved_to_send_out}</td>
                <td style="border: solid 1px black; width: 100px; padding:0; text-indent: 5px; color:{document_emailed_status}">{document_emailed}</td>
                <td style="border: solid 1px black; width: 100px; padding:0; text-indent: 5px; color:{document_moved_status}">{document_moved}</td>
            </tr>'''
    email_body += f'</table>'

    return email_body

def get_upload_process_history(engine):
    df = pd.read_sql(f'''
            select * from 
            (SELECT *, row_number() over (partition by Patient_Folder__Name order by datetime desc) as ranking 
            FROM program_performance 
            where module_name = "Upload Process"
            and datetime >= current_timestamp() - interval 1 day) A
            where ranking = 1
            order by module_complete desc, datetime asc, Patient_Folder__Name asc;''', con=engine)

    print_color(df, color='y')

    email_body = f'<br><span style="color:Black;font-weight:Bold; font-size:24px;">Upload Process:</span>'
    email_body += f'<br><span style="color:Black; ">See Below Satus of Upload Records Processed in the last 24 hours.</span>'
    email_body += f'<br><br><table>'
    email_body += f'''
              <tr>
                  <th style="border: solid; border-width:1px; width: 200px; padding:0;">Datetime</td>
                  <th style="border: solid; border-width:1px; width: 500px; padding:0;">Patient Folder</td>
                  <th style="border: solid; border-width:1px; width: 200px; padding:0;">Upload Complete ?</td>

              </tr>'''

    for k in range(df.shape[0]):
        datetime = df['datetime'].iloc[k]
        patient_folder = df['Patient_Folder__Name'].iloc[k]
        module_complete = df['module_complete'].iloc[k]
        module_complete = 'TRUE' if module_complete == 1 else 'FALSE'

        module_complete_status = 'Green' if module_complete == 'TRUE' else 'Red'

        email_body += f'''
                     <tr>
                         <td style="border: solid 1px black; width: 200px; padding:0; text-indent: 5px; ">{datetime}</td>
                         <td style="border: solid 1px black; width: 500px; padding:0; text-indent: 5px;">{patient_folder}</td>
                         <td style="border: solid 1px black; width: 200px; padding:0; text-indent: 5px; color:{module_complete_status}">{module_complete}</td>

                     </tr>'''
    email_body += f'</table>'

    return email_body


def get_merge_process_history(engine):
    df = pd.read_sql(f'''
        select * from 
        (SELECT *, row_number() over (partition by Patient_Folder__Name order by datetime desc) as ranking 
        FROM program_performance 
        where module_name = "Merge Process"
        and datetime >= current_timestamp() - interval 1 day) A
        where ranking = 1
        order by module_complete desc, datetime asc, Patient_Folder__Name asc;''', con=engine)

    print_color(df, color='y')

    email_body = f'<br><span style="color:Black;font-weight:Bold; font-size:24px;">Merge Process:</span>'
    email_body += f'<br><span style="color:Black; ">See Below Satus of Merge Records Processed in the last 24 hours.</span>'
    email_body += f'<br><br><table>'
    email_body += f'''
          <tr>
              <th style="border: solid; border-width:1px; width: 200px; padding:0;">Datetime</td>
              <th style="border: solid; border-width:1px; width: 500px; padding:0;">Patient Folder</td>
              <th style="border: solid; border-width:1px; width: 200px; padding:0;">Merge Complete ?</td>
            
          </tr>'''

    for k in range(df.shape[0]):
        datetime = df['datetime'].iloc[k]
        patient_folder = df['Patient_Folder__Name'].iloc[k]
        module_complete = df['module_complete'].iloc[k]
        module_complete = 'TRUE' if module_complete == 1 else 'FALSE'

        module_complete_status = 'Green' if module_complete == 'TRUE' else 'Red'

        email_body += f'''
                 <tr>
                     <td style="border: solid 1px black; width: 200px; padding:0; text-indent: 5px; ">{datetime}</td>
                     <td style="border: solid 1px black; width: 500px; padding:0; text-indent: 5px;">{patient_folder}</td>
                     <td style="border: solid 1px black; width: 200px; padding:0; text-indent: 5px; color:{module_complete_status}">{module_complete}</td>
                   
                 </tr>'''
    email_body += f'</table>'


    return email_body


def run_email_diagnostic(x, environment):
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:00:00" )
    GmailAPI = GoogleGmailAPI(credentials_file=x.gmail_credentials_file, token_file=x.gmail_token_file,
                              scopes=x.gmail_scopes)
    GsheetAPI = GoogleSheetsAPI(credentials_file=x.gsheet_credentials_file, token_file=x.gsheet_token_file,
                                scopes=x.gsheet_scopes,
                                sheet_id=x.google_sheet_published)

    engine = engine_setup(project_name=x.project_name, hostname=x.hostname, username=x.username, password=x.password,
                          port=x.port)
    email_body = f'<br><span style="color:Black;">Hello,</span>'
    email_body += f'<br><span style="color:Black;">See Below Diagnostic on Our Programs Performance;</span><br>'

    process_history_body = get_email_process_history(GsheetAPI)
    email_body += process_history_body

    email_body += '<br><br>'

    upload_history_body = get_upload_process_history(engine)
    email_body += upload_history_body

    email_body += '<br><br>'

    merge_history_body = get_merge_process_history(engine)
    email_body += merge_history_body

    GmailAPI.send_email(email_to=", ".join(x.notification_email), email_sender=x.email_sender,
                        email_subject=f'Medico Program Diagnostic {now}', email_cc=None, email_bcc=None,
                        email_body=email_body)






