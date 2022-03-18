from flask import Flask, request, render_template
import pandas as pd

app = Flask(__name__)
file_path = r"C:\Users\jorda\OneDrive\Documents\
            Work\JMM Group LLC\Subscribers\Email_list.xlsx"

@app.route('/')
def my_form():
    return render_template('subscribe.html')

@app.route('/', methods=['POST'])
def my_form_post():
    name = request.form['name']
    email = request.form['email']
    return name, email

post = my_form_post()
username = post[0]
mail = post[1]
try:
    with file_path as file:
        df = pd.read_csv(file)
        df.loc[len(df.index)] = [username, mail]
except PermissionError:
    df = pd.DataFrame(columns=["First Name", "Last Name", "Email"])
    df.loc[len(df.index)] = [username, mail]

writer = pd.ExcelWriter(file_path, engine='xlsxwriter')

df.to_excel(writer)

writer.save()
