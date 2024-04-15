import datetime
import pandas as pd
from gmail_api import GoogleGmailAPI
from google_sheets_api import GoogleSheetsAPI
from global_modules import print_color, create_folder, error_handler, engine_setup


def get_email_process_history(GsheetAPI):
    date_time = datetime.datetime.combine(datetime.datetime.now(), datetime.time.min) - datetime.timedelta(days=1)
    hour_now = datetime.datetime.now().hour
    print_color(f'hour_now: {hour_now}',color='y')
    print_color(date_time, color='g')
    df = GsheetAPI.get_data_from_sheet(sheetname="Auto Publish Data", range_name="A:Q")
    df['import_date'] = pd.to_datetime(df['import_date'])
    df = df[df['import_date']>=date_time]

    df2 = df[(df['all_fields_assigned'] == 'TRUE') | (df['approved_to_send_out_?'] == 'TRUE') | (
            df['document_emailed'] == 'TRUE') | (df['document_moved'] == 'TRUE')]

    df3 = df[(df['all_fields_assigned'] == 'FALSE') | (df['approved_to_send_out_?'] == 'FALSE') | (
                df['document_emailed'] == 'FALSE') | (df['document_moved'] == 'FALSE')]

    if hour_now > 12:
        df = df[(df['all_fields_assigned'] == 'FALSE') | (df['approved_to_send_out_?'] == 'FALSE') | (df['document_emailed'] == 'FALSE')| (df['document_moved'] == 'FALSE')]


    if df.shape[0] >0:
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


    else:
        email_body = f'<br><span style="color:Black;font-weight:Bold; font-size:24px;">Auto Publish Process:</span>'
        email_body += f'<br><span style="color:Black; ">All Records in the last 24 hours were Processed Correctly.</span>'

    email_body += f'<br><br><span style="color:Black; ">Records Processed in Total {df2.shape[0]}</span>'
    email_body += f'<br><span style="color:Black; ">Records Failed in Total {df3.shape[0]}</span>'


    return email_body


def get_upload_process_history(engine):
    df = pd.read_sql(f'''
            select * from 
            (SELECT *, row_number() over (partition by Patient_Folder__Name order by `datetime` desc) as ranking 
            FROM program_performance 
            where module_name = "Upload Process"
--             and module_complete = 0
            and `datetime` >= current_timestamp() - interval 1 day) A
            where ranking = 1
            order by module_complete desc, datetime asc, Patient_Folder__Name asc;''', con=engine)
    # import datetime
    hour_now = datetime.datetime.now().hour
    df2 = df[df['module_complete'] == 1]
    df3 = df[df['module_complete'] == 0]

    if hour_now > 12:
        df = df[df['module_complete'] ==0]
    print_color(df, color='y')

    if df.shape[0]>0:

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
            date_time = df['datetime'].iloc[k]
            patient_folder = df['Patient_Folder__Name'].iloc[k]
            module_complete = df['module_complete'].iloc[k]
            module_complete = 'TRUE' if module_complete == 1 else 'FALSE'

            module_complete_status = 'Green' if module_complete == 'TRUE' else 'Red'

            email_body += f'''
                         <tr>
                             <td style="border: solid 1px black; width: 200px; padding:0; text-indent: 5px; ">{date_time}</td>
                             <td style="border: solid 1px black; width: 500px; padding:0; text-indent: 5px;">{patient_folder}</td>
                             <td style="border: solid 1px black; width: 200px; padding:0; text-indent: 5px; color:{module_complete_status}">{module_complete}</td>
    
                         </tr>'''
        email_body += f'</table>'

    else:
        email_body = f'<br><span style="color:Black;font-weight:Bold; font-size:24px;">Upload Process:</span>'
        email_body += f'<br><span style="color:Black; ">All Records in the last 24 hours were Processed Correctly.</span>'

    email_body += f'<br><br><span style="color:Black; ">Records Processed in Total {df2.shape[0]}</span>'
    email_body += f'<br><span style="color:Black; ">Records Failed in Total {df3.shape[0]}</span>'

    return email_body


def get_merge_process_history(engine):
    df = pd.read_sql(f'''
        select A.*, case when B.Patient_Folder__Name is not null then True else False end as From_Upload from
        (select * from 
        (SELECT *, row_number() over (partition by Patient_Folder__Name order by datetime desc) as ranking 
        FROM program_performance 
        where module_name = "Merge Process"
--         and module_complete = 0
        and datetime >= current_timestamp() - interval 1 day) A
        where ranking = 1
      
        order by module_complete desc, datetime asc, Patient_Folder__Name asc) A
        left join
        (select * from (select *, row_number() over (partition by Patient_Folder__Name order by datetime desc) as ranking from program_performance where module_name = "Upload Process") A where ranking =1) B 
        on trim(a.Patient_Folder__Name) = trim(b.B.Patient_Folder__Name);''', con=engine)

    # import datetime
    hour_now = datetime.datetime.now().hour
    df2 = df[df['module_complete'] == 1]
    df3 = df[df['module_complete'] == 0]

    if hour_now > 12:

        df = df[df['module_complete'] == 0]

    print_color(df, color='y')

    if df.shape[0]>0:
        email_body = f'<br><span style="color:Black;font-weight:Bold; font-size:24px;">Merge Process:</span>'
        email_body += f'<br><span style="color:Black; ">See Below Satus of Merge Records Processed in the last 24 hours.</span>'
        email_body += f'<br><br><table>'
        email_body += f'''
              <tr>
                  <th style="border: solid; border-width:1px; width: 200px; padding:0;">Datetime</td>
                  <th style="border: solid; border-width:1px; width: 500px; padding:0;">Patient Folder</td>
                  <th style="border: solid; border-width:1px; width: 200px; padding:0;">Merge Complete ?</td>
                  <th style="border: solid; border-width:1px; width: 200px; padding:0;">From Upload Process ?</td>
                
              </tr>'''

        for k in range(df.shape[0]):
            date_time = df['datetime'].iloc[k]
            patient_folder = df['Patient_Folder__Name'].iloc[k]
            module_complete = df['module_complete'].iloc[k]
            module_complete = 'TRUE' if module_complete == 1 else 'FALSE'
            module_complete_status = 'Green' if module_complete == 'TRUE' else 'Red'
            from_upload = df['From_Upload'].iloc[k]
            from_upload = 'TRUE' if from_upload == 1 else 'FALSE'

            email_body += f'''
                     <tr>
                         <td style="border: solid 1px black; width: 200px; padding:0; text-indent: 5px; ">{date_time}</td>
                         <td style="border: solid 1px black; width: 500px; padding:0; text-indent: 5px;">{patient_folder}</td>
                         <td style="border: solid 1px black; width: 200px; padding:0; text-indent: 5px; color:{module_complete_status}">{module_complete}</td>
                         <td style="border: solid 1px black; width: 500px; padding:0; text-indent: 5px;">{from_upload}</td>
                     </tr>'''
        email_body += f'</table>'
    else:
        email_body = f'<br><span style="color:Black;font-weight:Bold; font-size:24px;">Merge Process:</span>'
        email_body += f'<br><span style="color:Black; ">All Records in the last 24 hours were Processed Correctly.</span>'

    email_body += f'<br><br><span style="color:Black; ">Records Processed in Total {df2.shape[0]}</span>'
    email_body += f'<br><span style="color:Black; ">Records Failed in Total {df3.shape[0]}</span>'

    return email_body


def get_task_performance_histroy(engine):
    df = pd.read_sql(f'''select `Function`, sum(Missed_run) as `Count of Times Program Failed to Complete` from
        (select *, case when time_difference >2 then 1 else 0 end as Missed_run from
        (select *, ifnull(TIME_TO_SEC(timediff(datetime, Prior_run))/ 60/ 60,0) as time_difference from
        (SELECT *, lag(datetime,1) over(partition by `function` order by datetime) as Prior_run
        FROM task_performance where DateTime >= current_timestamp() - interval 1 day
        and `function` != "Email Diagnostic"
        order by `function`, datetime) A) B) C
        group by `function`''', con=engine)

    email_body = f'<br><br><span style="color:Black;font-weight:Bold; font-size:24px;">Program RunTime Performance:</span>'
    email_body += f'<br><span style="color:Black; ">See Below History of Processes That Failed to Complete in the last 24 hours.</span>'

    if df.shape[0] > 0:

        email_body += f'<br><br><table>'
        email_body += f'''
                 <tr>
                     <th style="border: solid; border-width:1px; width: 200px; padding:0;">Module</td>
                     <th style="border: solid; border-width:1px; width: 500px; padding:0;">Fail Count</td>
                 </tr>'''

        for k in range(df.shape[0]):
            module = df['Function'].iloc[k]
            fail_count = df['Count of Times Program Failed to Complete'].iloc[k]

            email_body += f'''
                        <tr>
                            <td style="border: solid 1px black; width: 200px; padding:0; text-indent: 5px; ">{module}</td>
                            <td style="border: solid 1px black; width: 500px; padding:0; text-indent: 5px;">{fail_count}</td>
                          
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

    task_performance = get_task_performance_histroy(engine)
    email_body += task_performance

    GmailAPI.send_email(email_to=", ".join(x.diagnostic_email), email_sender=x.email_sender,
                        email_subject=f'Medico Program Diagnostic {now}', email_cc=None, email_bcc=None,
                        email_body=email_body)






