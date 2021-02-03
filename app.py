import os
import select

from flask import request, render_template, Flask, redirect, send_file
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func
from save_xlsx import SaveExcel


app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///RPC.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)


class Base(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    namePC = db.Column(db.String(100), nullable=False)
    IP = db.Column(db.String(12), nullable=False)
    MAC = db.Column(db.String(17), nullable=False)

    def __repr__(self):
        return '<Base %r>' % self.id


@app.route('/', methods=['GET', 'POST'])
def index():
    base = Base.query.order_by(Base.id.desc()).all()
    return render_template('index.html', base=base)


@app.route('/DEL', methods=['GET', 'POST'])
def DEL():
    base = Base.query.order_by(Base.id.desc()).all()
    return render_template('DEL_page.html', base=base)


@app.route('/ADD', methods=['POST', 'GET'])
def add_page():
    if request.method == "POST":
        NamePC = request.form['NAME_PC']
        ip = request.form['IP']
        mac = request.form['MAC']

        base = Base(namePC=NamePC, IP=ip, MAC=mac)

        try:
            db.session.add(base)
            db.session.commit()
            return redirect('/')

        except:
            return "При добавлении произошла ошибка"
    else:
        return render_template('ADD_page.html')


@app.route('/DEL/<int:id>')
def delete(id):
    base = Base.query.get_or_404(id)

    try:
        db.session.delete(base)
        db.session.commit()
        return redirect('/')
    except:
        return 'При удалении статьи произошла ошибка'


@app.route('/save', methods=['GET', 'POST'])
def save():
    save_exel = SaveExcel("table.xlsx")
    rows = Base.query.order_by(Base.id.desc()).all()
    i = 0
    for row in rows:
        save_exel.write_workbook(i+2, 1, row.namePC)
        save_exel.write_workbook(i+2, 2, row.IP)
        save_exel.write_workbook(i+2, 3, row.MAC)
        i += 1
    save_exel.save_excel()
    return send_file('table1.xlsx', attachment_filename='table1.xlsx')


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    app.run(host='0.0.0.0', port=port)
