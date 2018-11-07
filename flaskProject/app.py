from flask import Flask, render_template, flash, redirect, url_for, session,request, logging
from flask_wtf import FlaskForm
from flask_mysqldb import MySQL
from wtforms import StringField, TextAreaField, PasswordField, validators
from passlib.hash import sha256_crypt
from data import Articles

app= Flask(__name__)
app.secret_key = 's3cr3t'

#config MySQL
app.config['MYSQL_HOST']='127.0.0.1'
app.config['MYSQL_USER']='root'
app.config['MYSQL_PASSWORD']=''
app.config['MYSQL_DB']='myflaskapp'
#app.config['MYSQL_HOST']='DictCursor'

#init MYSQL_DB

mysql = MySQL(app)

Articles = Articles()

@app.route('/')
def home():
    return render_template('home.html')

@app.route('/about')
def about():
    return render_template('about.html')

@app.route('/welcome')
def welcome():
    return render_template('welcome.html')

@app.route('/articles')
def articles():
    return render_template('articles.html', articles =Articles)

class RegisterForm(FlaskForm):
    name = StringField ('Name', [validators.Length(min=1, max=50)])
    username = StringField ('UserName', [validators.Length(min=4, max=25)])
    email = StringField ('Email', [validators.Length(min=5, max=50)])
    password = PasswordField('Password',
    [ validators.DataRequired(),
    validators.EqualTo('confirm', message ='password do not match')
    ]
    )
    confirm = PasswordField('Confirm Password')

@app.route('/register', methods=['GET', 'POST'])
def register():
    form = RegisterForm(request.form)
    if request.method =='POST' and form.validate():
        name = form.name.data
        email = form.email.data
        username = form.username.data
        password = sha256_crypt.encrypt(str(form.password.data))

        # create DictCursor
        cur = mysql.connection.cursor()
        cur.execute("INSERT INTO users (name, email, username, password) VALUES(%s, %s, %s, %s)", (name, email, username, password))

        #commit to DB
        mysql.connection.commit()

        #cursor close
        cur.close()
        flash('You are now registered', 'success')
        redirect(url_for('home'))
    return render_template('register.html', form = form)


if __name__ == '__main__':
    app.run(debug=True)
