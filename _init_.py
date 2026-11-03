from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
from flask_mysqldb import MySQL, MySQLdb
import openpyxl
from openpyxl import Workbook
from io import BytesIO

app = Flask(__name__)
app.secret_key = 'many random bytes'

app.config['MYSQL_HOST'] = 'localhost'
app.config['MYSQL_USER'] = 'root'
app.config['MYSQL_PASSWORD'] = ''
app.config['MYSQL_DB'] = 'crudapplication'

mysql = MySQL(app)

@app.route('/')
def Index():
    cur = mysql.connection.cursor()
    cur.execute("SELECT * FROM students")
    data = cur.fetchall()
    cur.close()
    return render_template('index2.html', students=data)

@app.route('/insert', methods=['POST'])
def insert():
    if request.method == "POST":
        flash("Data Inserted Successfully")
        name = request.form['name']
        email = request.form['email']
        phone = request.form['phone']
        address = request.form['address']
        state = request.form['state']
        gst_num = request.form['gst_num']
        client_type = request.form['client_type']
        cur = mysql.connection.cursor()
        cur.execute("INSERT INTO students (name, email, phone,address,state,gst_num,client_type) VALUES (%s, %s, %s, %s, %s, %s, %s)", (name, email, phone,address,state,gst_num,client_type))
        mysql.connection.commit()
        return redirect(url_for('Index'))

@app.route('/delete/<string:id_data>', methods=['GET'])
def delete(id_data):
    flash("Record Has Been Deleted Successfully")
    cur = mysql.connection.cursor()
    cur.execute("DELETE FROM students WHERE id=%s", (id_data,))
    mysql.connection.commit()
    return redirect(url_for('Index'))

@app.route('/update', methods=['POST'])
def update():
    if request.method == 'POST':
        id_data = request.form['id']
        name = request.form['name']
        email = request.form['email']
        phone = request.form['phone']
        address = request.form['address']
        state = request.form['state']
        gst_num = request.form['gst_num']
        client_type = request.form['client_type']
        cur = mysql.connection.cursor()
        cur.execute("""
               UPDATE students
               SET name=%s, email=%s, phone=%s, address=%s, state=%s, gst_num=%s, client_type=%s
               WHERE id=%s
            """, (name, email, phone, address, state, gst_num, client_type, id_data,))
        flash("Data Updated Successfully")
        mysql.connection.commit()
        return redirect(url_for('Index'))

@app.route('/export', methods=['GET'])
def export():
    cur = mysql.connection.cursor(MySQLdb.cursors.DictCursor)
    cur.execute("SELECT name, email, phone, address, state, gst_num, client_type FROM students")
    rows = cur.fetchall()
    cur.close()

    output = BytesIO()
    wb = Workbook()
    sh = wb.active
    sh.title = 'data'

    sh['A1'] = 'Name'
    sh['B1'] = 'Email'
    sh['C1'] = 'Phone'
    sh['D1'] = 'Address'
    sh['E1'] = 'State'
    sh['F1'] = 'GST Number'
    sh['G1'] = 'Client type'


    # Write data
    for idx, row in enumerate(rows, start=1):
        sh[f'A{idx+1}'] = row['name']
        sh[f'B{idx+1}'] = row['email']
        sh[f'C{idx+1}'] = row['phone']
        sh[f'D{idx+1}'] = row['address']
        sh[f'E{idx+1}'] = row['state']
        sh[f'F{idx+1}'] = row['gst_num']
        sh[f'G{idx+1}'] = row['client_type']

    wb.save(output)
    output.seek(0)
    return send_file(output, download_name='data.xlsx', as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == "__main__":
    app.run(debug=True)
