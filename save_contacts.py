import cgi
import cgitb
import xlsxwriter
import pandas as pd

file_path = r"C:\Users\jorda\OneDrive\Documents\
            Work\JMM Group LLC\Subscribers\Email_list.xlsx"
cgitb.enable()

form = cgi.FieldStorage()

first_name = form.getvalue('first_name')
last_name = form.getvalue('last_name')
email = form.getvalue('email')


df = pd.DataFrame(columns=["First Name", "Last Name", "Email"])
df.loc[len(df.index)] = [first_name, last_name, email]

writer = pd.ExcelWriter(file_path, engine='xlsxwriter')

df.to_excel(writer)

writer.save()
