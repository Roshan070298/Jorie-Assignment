from flask import Flask, request
import pandas as pd

app=Flask(__name__)

File = ""

def load_file():
    pd.DataFrame.to_excel(File)
    pd.read_excel(File)

def save_excel(df):
    df.to_excel(File)

@app.route("/addcolumn", methods=["PUT"])
def create_column():
    data=request.json
    column_name = data.get("name")
    df = load_file()
    df[column_name] = None
    save_excel(df)

@app.route("/rename/column", methods =["Patch"])
def rename_column():
    data=request.json
    column_old_name = data.get("old_name")
    column_new_name = data.get("new_name")
    df = load_file()
    df.rename({column_old_name: column_new_name}, inplace=True)
    save_excel(df)

@app.route("/fetchfile", methods=["GET"])
def fetch_file_data():
    df = load_file()
    return df.to_json()

@app.route("/updatefile", methods=["POST"] )
def update_file():
    data = request.json
    row=data.get("row")
    column=data.get("column")
    value=data.get("value")
    df =load_file()
    df.at[row, column]=value
    save_excel(df)

@app.route("/deletefile", mehtods=["DELETE"])
def empty_excel():
    df =pd.DataFrame()
    save_excel(df)

