from flask import Flask, request, render_template, url_for, redirect
import pandas as pd

app = Flask(__name__)
file_path = r"C:\Users\jorda\OneDrive\Documents\
            Work\JMM Group LLC\Subscribers\Email_list.xlsx"

@app.route('/')
def home():
    return render_template("index.html/index.html")


@app.route('/subscribe.html', methods=['GET', 'POST'])
def my_form_post():
    if request.method == 'POST':
        name = request.form.get('name')
        email = request.form.get('email')
        print(name)
        print(email)
        try:
            with file_path as file:
                df = pd.read_csv(file)
                df.loc[len(df.index)] = [name, email]
        except PermissionError:
            df = pd.DataFrame(columns=["Name", "Email"])
            df.loc[len(df.index)] = [name, email]
        writer = pd.ExcelWriter(file_path, engine='xlsxwriter')

        df.to_excel(writer, sheet_name='Email List')

        writer.save()
        return redirect(url_for("user", usr=user))
    else:
        return render_template('index.html/index.html')

@app.route('/<usr>')
def user(usr=None):
    return f'<h1>{usr}</h1>'


if __name__ == "__main__":
    app.run()
