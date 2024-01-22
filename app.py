from flask import Flask, request, render_template, redirect, session
from werkzeug.utils import secure_filename
import os
import csv
import mysql.connector
import pandas as pd
import time
import chardet
from email.mime.text import MIMEText
import smtplib
import re
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import *
import base64
import shutil
import mimetypes
import json
import sqlite3

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'csv', 'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def verify_emails(data):
    pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b'
    for row in data:
        email= row['email']

        if re.match (pattern, email):
            row['checked'] = 1
        else:
            row['checked'] = 0
    return data

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def read_file_into_dict(filename):
    with open(filename, 'rb') as file:
        encoding = chardet.detect(file.read())['encoding']
    if filename.endswith('.csv'):
        with open(filename, 'r', encoding = encoding) as file:
            reader = csv.reader(file)
            next(reader, None)
            data = [{'email': row[0], 'name': row[1]} for row in reader]
    elif filename.endswith('.xlsx'):
        df = pd.read_excel(filename, header=None)
        data = [{'email': row[0], 'name': row[1]} for row in df.values]
    else:
        return None
    data_with_check = add_column_and_print(data)
    return data_with_check

def execute_sql(code):
    # Create a connection object to the database file
    conn = sqlite3.connect("mailer.db")

    # Create a cursor object to execute SQL commands
    cur = conn.cursor()

    # Create a table named "students" with three columns: id, name, and grade
    sql = code
    cur.execute(sql)

    # Commit the changes to the database
    conn.commit()

    # Close the connection
    conn.close()

def create_adresses_table():
    sql = "CREATE TABLE emails (email_adress VARCHAR(255) NOT NULL, name VARCHAR(255), PRIMARY KEY(email_adress));"
    execute_sql(sql)

def create_temp_adresses_table():
    sql2 = "DROP TABLE IF EXISTS temp_emails;"
    execute_sql(sql2)
    sql = "CREATE TABLE temp_emails (email_adress VARCHAR(255), name VARCHAR(255));"
    execute_sql(sql)

def insert_data_into_temp_table(data, table_name):
    # Create a connection object to the database file
    conn = sqlite3.connect("mailer.db")

    # Create a cursor object to execute SQL commands
    cur = conn.cursor()

    sql = f"INSERT INTO {table_name} (email_adress, name) VALUES (?, ?)"

    values = [(row['email'], row['name']) for row in data]
    cur.executemany(sql, values)

   # Commit the changes to the database
    conn.commit()
    # Close the connection
    conn.close()

def create_emails_sent_to_table():
    #email_id is the id in the emails_sent_table
    sql = "CREATE TABLE email_sent_to (email_id INT, emailed_to VARCHAR(255));"
    execute_sql(sql)

def create_emails_sent_table():
    sql = "CREATE TABLE emails_sent (id INTEGER PRIMARY KEY, subject VARCHAR(255), body text);"

    execute_sql(sql)

def read_table_into_array(table_name, columns):
    # Create a connection object to the database file
    conn = sqlite3.connect("mailer.db")
    cursor = conn.cursor()

    sql = f'SELECT {columns} FROM {table_name}'
    cursor.execute(sql)

    data = cursor.fetchall()
    conn.close()
    data_dic = [{'email': row[0], 'name': row[1]} for row in data]
    return data_dic

def read_mail_id(table_name, columns):
    # Create a connection object to the database file
    conn = sqlite3.connect("mailer.db")
    cursor = conn.cursor()

    sql = f'SELECT {columns} FROM {table_name}'
    cursor.execute(sql)

    data = cursor.fetchall()
    conn.close()
    #data_dic = [{'email_id': row[0]} for row in data]
    return data[-1]

def check_if_mail_was_sent(mail_body):
    # Create a connection object to the database file
    conn = sqlite3.connect("mailer.db")
    cursor = conn.cursor()

    sql = f"Select id From emails_sent Where body = '{mail_body}' "
    cursor.execute(sql)
    data = cursor.fetchall()
    cursor.close()
    conn.close()
    return data
    
def remove_adresess_that_the_mail_was_sent_to(mail_id):
    #sql = f"Delete From temp_emails Where NOT EXISTS (temp_emails.email_adress IN (Select emailed_to From email_sent_to Where (email_sent_to.email_id = {mail_id})))"
    sql = f"Delete From temp_emails Where temp_emails.email_adress IN (Select emailed_to From email_sent_to Where (email_sent_to.email_id != {mail_id}) )"
    execute_sql(sql)
#mails
def template():
    # Create a connection object to the database file
    conn = sqlite3.connect("mailer.db")
    cur = conn.cursor()
    cur.execute("SELECT * FROM email_templates")
    data = cur.fetchall()
    cur.close()
    conn.close()
    campaigns = []
    subjects = {}
    email_body = {}
    for row in data:
        campaigns.append(row[0])
        email_body[row[0]] = row[2]
        subjects[row[0]] = row[1]
    return campaigns, email_body, subjects

def insert_mail(campaign, subject, email_content):
    # Create a connection object to the database file
    conn = sqlite3.connect("mailer.db")
    cur = conn.cursor()
    sql = "INSERT INTO email_templates (campaign, subject, email_content) VALUES (?, ?, ?)"
    val = (campaign, email_content, subject)
    cur.execute(sql, val)
    conn.commit()

    cur.close()
    conn.close()

# Define the function that takes a folder path as input
def delete_all_files(folder_path):
    # Check if the folder path exists and is a directory
    if os.path.exists(folder_path) and os.path.isdir(folder_path):
        # Loop through all the files and subdirectories in the folder
        for item in os.listdir(folder_path):
            # Join the folder path with the item name
            item_path = os.path.join(folder_path, item)
            # If the item is a file, delete it using os.remove
            if os.path.isfile(item_path):
                os.remove(item_path)
            # If the item is a subdirectory, delete it and its contents using shutil.rmtree
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)
        # Return a success message
        return "All files and subdirectories in {} have been deleted.".format(folder_path)
    # If the folder path does not exist or is not a directory, return an error message
    else:
        return "Invalid folder path. Please provide a valid folder path."

#Almost ready
def send_email(sender, subject, body, api_key, recipients, attachment_paths):
    content = str(body)

    message = Mail(
        from_email=sender,
        to_emails=recipients,
        subject=subject,
        html_content= content
    )
    
    if attachment_paths != "No Files":
        for attachment_path in attachment_paths:
            # Get the file name and extension from the attachment path
            file_name, file_ext = os.path.splitext(attachment_path)
            file_name = file_name.split("\\")
            file_name = file_name[-1]

            mime_type, _ = mimetypes.guess_type(attachment_path)

            # Open the attachment file in binary mode and read its content
            with open(attachment_path, "rb") as f:
                file_data = f.read()
            
            # Encode the file data in base64 format
            file_encoded = base64.b64encode(file_data).decode()
            
            # Create an Attachment object with the file content, type, name, disposition, and content ID
            attachment = Attachment(
                file_content=FileContent(file_encoded),
                file_type=FileType(mime_type), 
                file_name=FileName(file_name + file_ext),
                disposition=Disposition("attachment"),
                content_id=ContentId(file_name)
            )
            
            # Add the attachment to the message
            message.add_attachment(attachment)      
    delete_all_files(os.path.join(app.root_path, 'temp'))                                            
    
    #message.add_header('')
    message.add_content('To stop receiving all emails from us, please click here: <%asm_global_unsubscribe_url%>', 'text/plain')

    try:
        sg = SendGridAPIClient(api_key)
        response = sg.send(message)
        print(response.status_code)
        print(response.body)
        print(response.headers)
    except Exception as e:
        print(e.message)
    return "Success"
    
def add_column_and_print(data):

    new_data = [{**d, 'checked': 1} for d in data]
    return new_data

def create_templates_taple():
    sql = "CREATE TABLE email_templates ( campaign varchar(255) NOT NULL, subject varchar(255) DEFAULT NULL, email_content text DEFAULT NULL)"
    execute_sql(sql)

# Define a function that takes a byte string as input
def aggregate_metrics(data):
  # Decode the byte string to a normal string
  data = data.decode()
  # Parse the string into a Python object
  data = json.loads(data)
  # Initialize an empty dictionary to store the aggregated metrics
  agg = {}
  # Loop through each dictionary in the input list
  for d in data:
    # Loop through each key-value pair in the dictionary
    for k, v in d.items():
      # If the key is 'date', skip it
      if k == 'date':
        continue
      # If the key is 'stats', loop through the list of dictionaries in the value
      elif k == 'stats':
        for s in v:
          # Loop through each key-value pair in the stats dictionary
          for sk, sv in s.items():
            # If the key is 'metrics', loop through the dictionary in the value
            if sk == 'metrics':
              for mk, mv in sv.items():
                # If the metric key is already in the aggregated dictionary, add the value to it
                if mk in agg:
                  agg[mk] += mv
                # Otherwise, initialize the metric key with the value
                else:
                  agg[mk] = mv
  # Return the aggregated dictionary
  return agg

@app.route('/')
def index():
    return render_template('index.html')

####Analytics page 1
@app.route('/analytics_start_date')
def analytics_start_date():
    return render_template('analytics_start_date.html')

########Analytics page 2 
@app.route('/show_analytics')
def show_analytics():
    api_keyo = request.args.get('key_string')
    date = request.args.get('string')
    sg = SendGridAPIClient(api_keyo)

    params = {'start_date': date}

    response = sg.client.stats.get(
        query_params=params
    )
    stats = aggregate_metrics(response.body)
    stats["start_date"] = date
    
    return render_template('show_analytics.html', data = stats)

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return 'No file part'
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'
    if file and allowed_file(file.filename):
        delete_all_files(os.path.join(app.root_path, 'temp'))
        filename = secure_filename(file.filename)
        
        file.save(filename)
        data = read_file_into_dict(filename)

        sql0 = "Delete from temp_emails"
        execute_sql(sql0)

        insert_data_into_temp_table(data, "temp_emails")
        return 'correct'
    else:
        return 'Invalid file type'

@app.route("/upload_attach", methods=["POST"])
def upload_attach():
    # Get the file from the request
    file = request.files["file"]

    # Save the file to the uploads folder with the new file name
    os.path.join(app.root_path, "temp/", file.filename)
    file.save(os.path.join(app.root_path, "temp/", file.filename))
    # Return a success message with the new file name
    return "File uploaded successfully as " + file.filename

@app.route('/listOfMails', methods=['GET', 'POST'])
def verify_mails():
    data = read_table_into_array("temp_emails","*")
    
    if request.method == 'POST':
        data = verify_emails(data)
        count = deleteunvalidated(data)
    return render_template('verify_page.html', data=data)

def deleteunvalidated(list_of_dicts):
  # Create a connection object to the database file
  connection = sqlite3.connect("mailer.db")
  cursor = connection.cursor()
  for row in list_of_dicts:
    if row["checked"] != 1:
      email = row["email"]
      sql_delete_query = "DELETE FROM temp_emails WHERE email_adress = ? ;"
      cursor.execute(sql_delete_query, (email,))
      connection.commit()
  connection.close()

@app.route('/update', methods=['GET'])
def update():
    return jsonify(data)

@app.route("/email-template")
def email_template():
    campaigns, email_body, subjects = template() 
    return render_template("meail-templates.html", campaigns=campaigns, subjects= subjects) 

@app.route("/fill", methods=["POST"])
def fill():
    campaign = request.form.get("campaign")
    campaigns, email_body, subjects = template()
    return [email_body[campaign], subjects[campaign]]
    
@app.route("/save", methods=["POST"])
def save():
    # Get the data from the request form
    subject = request.form.get("subject")
    text = request.form.get("text")
    name = request.form.get("name")

    insert_mail(name,text,subject)
    return '', 204

def get_send_list(id):

    # Create a connection object to the database file
    connection = sqlite3.connect("mailer.db")
    cursor = connection.cursor()
    sql_get_entries_query = "SELECT * FROM temp_emails"
    cursor.execute(sql_get_entries_query)
    entries = cursor.fetchall()
    final_emails = []
    try:
        create_emails_sent_table()
    except:
        pass
    for entry in entries:

        email = entry[0]
        name = entry[1]

        sql_save_address_query = f"Insert INTO emails (email_adress, name) VALUES (?, ?)"
        try:
            execute_sql(sql_save_address_query, (email, name))
        except:
            pass

        final_emails.append(entry[id])
    connection.close()
    return final_emails

def uploaded_files_names(source_folder, destination_folder, new_file_name):
    # Get the list of files in the source folder

    files = os.listdir(source_folder)
    files_names = []
    if (len(files)!=0):
        for file in files:
            files_names.append(os.path.basename(file))  

    return files_names

@app.route('/send', methods=['POST'])
def send():

    email = request.form['email']
    password = request.form['password']
    subject = request.form.get("hiddenSubject")
    text = request.form.get("hiddenText")

    mail_sent_check = check_if_mail_was_sent(text)

    if len(mail_sent_check) == 0:

        sql = f"INSERT INTO emails_sent (subject, body) VALUES ('{subject}', '{text}');"
        execute_sql(sql)
        mail_sent_id = read_mail_id("emails_sent", "id")[0]
    
    else:
        mail_sent_id = mail_sent_check[0][0]
        remove_adresess_that_the_mail_was_sent_to(mail_sent_id)
    
    sql = f'INSERT INTO email_sent_to (email_id, emailed_to) SELECT "{mail_sent_id}", email_adress FROM temp_emails;'
    execute_sql(sql)
    
    recipients = get_send_list(0)
    recipients_names = get_send_list(1)

    files_names = uploaded_files_names((app.root_path + '/' + 'temp'), app.root_path + '/' + 'uploads', str(mail_sent_id))


    if len(files_names) != 0:
        attachment_path = []
        for file_name in files_names:
            attachment_path.append(os.path.join(app.root_path, 'temp', file_name))
    else:
        attachment_path = "No Files"
        
    if recipients != []:
        for i in range(len(recipients)):
            personalized_text = text.replace("{{name}}", recipients_names[i])
            send_email(email, subject, personalized_text, password, recipients[i], attachment_path)
        return '', 204
    else:
        return "refresh the list!"
    
try:
    create_templates_taple()
except:
    pass
try:
    create_emails_sent_table()
except:
    pass
try:
    create_emails_sent_to_table()
except:
    pass
try:
    create_adresses_table()
except:
    pass
try:
    create_temp_adresses_table()
except:
    sql0 = "Delete from temp_email"
    execute_sql(sql0)
    pass

if __name__ == '__main__':
    app.run(debug=True)