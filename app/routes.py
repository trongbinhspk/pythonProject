from flask import Flask
app = Flask(__name__)
from forms import LoginForm
from flask import render_template, flash, redirect
from config import Config


app.config['SECRET_KEY'] = 'khong-doan-noi-dau'

@app.route('/index')
def index():
     return "Hello, World!"

@app.route('/login', methods=['GET', 'POST'])
def login():
    form = LoginForm()

    return render_template('login.html', title='Sign In', form=form)

if __name__ == '__main__':
    app.run()