import pandas as pd
from tkinter import Tk, Toplevel, mainloop, Button, \
    Label, filedialog, Text, Listbox, END, Entry, messagebox
from tkinter.filedialog import askopenfilename
from simple_salesforce import Salesforce
from salesforce_bulk import CsvDictsAdapter, SalesforceBulk


def login():
    """Takes the user's credentials via tkinter form and shows action options"""
    global sf
    try:
        global USERNAME, PASSWORD, TOKEN
        USERNAME, PASSWORD, TOKEN = user.get(), pwd.get(), token.get()
        sf = Salesforce(
            username=USERNAME,
            password=PASSWORD,
            security_token=TOKEN
        )
        messagebox.showinfo("Success!", "Login successful")
        Label(root, text="What do you wish to do?").grid(row=10, column=1)
        button_query = Button(
            root, height=1, width=15, text="Query",
            command=open_query
        )
        button_query.grid(row=11, column=1)
        Label(root, text="").grid(row=13, column=1)
        button_insert = Button(
            root, height=1, width=15, text="Insert",
            command=lambda: open_secondary("Insert")
        )
        button_insert.grid(row=12, column=1)
        button_update = Button(
            root, height=1, width=15, text="Update",
            command=lambda: open_secondary("Update")
        )
        button_update.grid(row=13, column=1)
        button_delete = Button(
            root, height=1, width=15, text="Delete",
            command=lambda: open_secondary("Delete")
        )
        button_delete.grid(row=14, column=1)
        Label(root, text=" ").grid(row=15, column=1)

    except Exception as e:
        messagebox.showerror("Login Error", e)


def browse_button():
    """Prompts the user to select a folder and saves its path"""
    global folder_path
    folder_path = filedialog.askdirectory()
    messagebox.showinfo("Success!", folder_path + " Selected")


def browse_file():
    """Prompts the user to select a file and saves its path"""
    global filename
    filename = filedialog.askopenfilename()
    messagebox.showinfo("Success!", filename + " Selected")


def select(lstbox, multiple):
    """Returns the user's selection a Listbox

    Keyword arguments:
    lstbox -- Tkinter Listbox Widget
    multiple -- indicates whether to return a string or a list"""
    reslist = list()
    selection = lstbox.curselection()
    for i in selection:
        entrada = lstbox.get(i)
        reslist.append(entrada)
    if (multiple == 1):
        return reslist
    else:
        return reslist[0]


def query_where(lista):
    """Takes a list and returns a string of its items separated by ','"""
    string = ""
    for record in lista:
        string += str(record) + ","
    string = string[:-1]
    return string


def extract_fields():
    """Compiles all the fields within a Salesforce.com object on a Tkinter Listbox"""
    global object_name
    object_name = select(entity, 0)
    options = sf.query_all(
        ("SELECT ID, QualifiedAPIName from FieldDefinition "
         "where EntityDefinitionId = '" + select(entity, 0)
         + "' order by QualifiedApiName")
    )
    optionList = []
    fields.delete(0, END)
    for record in options['records']:
        optionList.append(record['QualifiedApiName'])
    for each_item in range(len(optionList)):
        fields.insert(END, optionList[each_item])


def write_query():
    """Compiles a SOQL query in a Text widget"""
    string = "SELECT "
    string += query_where(select(fields, 1)) + " from " + object_name
    soql.delete('1.0', END)
    soql.insert('1.0', string)


def define_df():
    """Prompts the user to select a file and save it on a DataFrame"""
    global df
    name = askopenfilename(
        filetypes=[('CSV', '*.csv',), ('Excel', ('*.xls', '*.xlsx'))]
    )
    if name:
        if name.endswith('.csv'):
            df = pd.read_csv(name)
        else:
            df = pd.read_excel(name)
    messagebox.showinfo("Success!", name + " Selected")


def open_query():
    """Opens the Query TopLevel window"""
    query_window = Toplevel(root)
    query_window.title("Query")
    global entity
    entity = Listbox(query_window, selectmode="single", width=70)
    entity.grid(row=0, column=1)
    options = sf.query_all(
        ("SELECT ID, QualifiedAPIName from EntityDefinition"
         " order by QualifiedAPIName")
    )
    optionList = []
    entity.delete(0, END)
    for record in options['records']:
        optionList.append(record['QualifiedApiName'])
    for each_item in range(len(optionList)):
        entity.insert(END, optionList[each_item])

    show_fields = Button(query_window, text="Show Fields", command=extract_fields)
    show_fields.grid(row=1, column=1)
    global fields
    fields = Listbox(query_window, selectmode="multiple", width=70)
    fields.grid(row=2, column=1)

    create_query = Button(query_window, text="Create Query", command=write_query)
    create_query.grid(row=4, column=1)

    global soql
    soql = Text(query_window, width=70, height=5)
    soql.grid(row=5, column=1)
    soql.insert(END, " ")
    global folder
    folder = Label(query_window, text="Select Folder")
    folder.grid(row=6, column=1)

    button_browse = Button(query_window, text="Browse", command=browse_button)
    button_browse.grid(row=7, column=1)
    global file
    file = Entry(query_window, width=70)
    file.grid(row=8, column=1)
    Label(query_window, text="   ").grid(row=9, column=1)

    button_extract = Button(query_window, height=1, width=15, text="Extract", command=extract)
    button_extract.grid(row=10, column=1)


def open_secondary(operation):
    """Opens a TopLevel window used for 'Write' operations

    Keyword arguments
    operation -- determines the program's action (Delete, Insert, or Update)

    For all of the possible actions, the input Excel or CSV file must contain
    columns with the exact API Names found in the Salesforce organization"""
    secondary_window = Toplevel(root)
    secondary_window.title(operation)
    global entity
    entity = Listbox(secondary_window, selectmode="single", width=70)
    entity.grid(row=0, column=1)
    options = sf.query_all("SELECT ID, QualifiedAPIName from EntityDefinition order by QualifiedAPIName")
    optionList = []
    entity.delete(0, END)
    for record in options['records']:
        optionList.append(record['QualifiedApiName'])
    for each_item in range(len(optionList)):
        entity.insert(END, optionList[each_item])

    Label(secondary_window,
          text="Select a Excel or CSV file with the exact fields you want to " + operation.lower() + " with the correct API Names").grid(
        row=6, column=1)

    button_browse = Button(secondary_window, text="Browse Input File", command=define_df)
    button_browse.grid(row=7, column=1)

    global folder
    folder = Label(secondary_window, text="Select Folder to store results file")
    folder.grid(row=8, column=1)

    button_browse = Button(secondary_window, text="Browse Results Folder", command=browse_button)
    button_browse.grid(row=9, column=1)

    Label(secondary_window, text="   ").grid(row=10, column=1)

    button_action = Button(secondary_window, text=str(operation), command=lambda: action(operation))
    button_action.grid(row=11, column=1)


def action(operation):
    """Performs the Insertion, Deletion, or Update in the Salesforce org"""
    global object_name
    object_name = select(entity, 0)
    impacted_records = []
    for index in range(len(df)):
        record = {}
        for col in df.columns:
            record[col] = df[col][index]
        impacted_records.append(record)

    try:
        MsgBox = messagebox.askquestion("Operation", (
            'You are about to {action} {length} records within the {obj} object within your Salesforce org. Are you sure you want to proceed?').format(
            action=operation.lower(), length=str(len(impacted_records)), obj=object_name), icon='warning')
        if (MsgBox == 'yes'):
            bulk = SalesforceBulk(username=USERNAME, password=PASSWORD, security_token=TOKEN)
            if (operation == "Delete"):
                job = bulk.create_delete_job(object_name, contentType='CSV')
            elif (operation == "Insert"):
                job = bulk.create_insert_job(object_name, contentType='CSV')
            else:
                job = bulk.create_update_job(object_name, contentType='CSV')
            csv_iter = CsvDictsAdapter(iter(impacted_records))
            batch = bulk.post_batch(job, csv_iter)
            bulk.wait_for_batch(job, batch)
            bulk.close_job(job)

            result_df = pd.DataFrame(impacted_records)
            results = bulk.get_batch_results(bulk.get_batch_list(job)[0]['id'])
            result_df['ID'] = ""
            result_df['SUCCESS'] = ""
            result_df['ERROR'] = ""
            for index in range(len(result_df)):
                result_df['ID'][index] = str(results[index]).split("'")[1]
                result_df['SUCCESS'][index] = str(results[index]).split("'")[3]
                result_df['ERROR'][index] = str(results[index]).split("'")[7]
            input_file = folder_path + "/" + "results" + bulk.get_batch_list(job)[0]['id'] + ".xlsx"
            result_df.to_excel(input_file)

            messagebox.showinfo("Info", (
                'Job Details:\n\nNumber of Records Processed: {recordsProcessed}\n'
                'Number of Records Failed: {recordsFailed}').format(
                recordsProcessed=bulk.job_status(job)['numberRecordsProcessed'],
                recordsFailed=bulk.job_status(job)['numberRecordsFailed']))
    except Exception as e:
        messagebox.showerror("Error", e)


def extract():
    """Executes the query built by the user and saves the results"""
    try:
        query = sf.query_all(soql.get("1.0", END))['records']
        fields = list(query[0].keys())
        fields.remove('attributes')
        Records = pd.DataFrame(columns=fields)
    except Exception as e:
        messagebox.showerror("Verify your Query", e)
    for record in query:
        dic = {}
        for field in fields:
            dic[field] = [record[field]]
        temp = pd.DataFrame.from_dict(dic)
        Records = Records.append(temp)
    inputfileName = folder_path + "/" + file.get() + ".xlsx"
    Records = Records.set_index(Records.columns[0])
    Records.to_excel(inputfileName)
    messagebox.showinfo("Success!", file.get() + ".xlsx" + " created!")


root = Tk()
root.title("Salesforce API Extractor")

Label(root, text="Username").grid(row=2, column=1)
user = Entry(root, width=70)
user.grid(row=3, column=1)

Label(root, text="Password").grid(row=4, column=1)
pwd = Entry(root, width=70, show="*")
pwd.grid(row=5, column=1)

Label(root, text="Security Token").grid(row=6, column=1)
token = Entry(root, width=70, show="*")
token.grid(row=7, column=1)

buttonFields = Button(root, height=1, width=15, text="Login", command=login)
buttonFields.grid(row=8, column=1)
Label(root, text="").grid(row=9, column=1)

mainloop()
