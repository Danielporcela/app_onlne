from flask import Flask, render_template, request, redirect, session
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
from datetime import datetime
import os
from sqlalchemy import text

app = Flask(__name__)

app.config['SECRET_KEY'] = '123456'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///banco.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# ================= BANCO =================

class Setor(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    codigo = db.Column(db.String(20), unique=True)
    total = db.Column(db.Float, default=0)
    horas = db.Column(db.Integer, default=0)

class Planilha(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(200), unique=True)

with app.app_context():

    db.create_all()

    try:
        db.session.execute(text("ALTER TABLE setor ADD COLUMN horas INTEGER DEFAULT 0"))
        db.session.commit()
    except:
        pass


# ================= LINHAS DE HORARIO =================
# Corrigido para usar as linhas reais da planilha

def encontrar_linhas_hora(df):
    return [26, 54, 82]


# ================= CALCULO HORAS =================

def calcular_horas(valor):

    if pd.isna(valor):
        return 0

    hora = None
    minuto = None

    try:

        valor_int = int(valor)

        if valor_int >= 100:

            hora = valor_int // 100
            minuto = valor_int % 100

    except:
        pass

    if hora is None:

        try:

            if isinstance(valor,str):

                dt = datetime.strptime(valor.strip(), "%H:%M")

                hora = dt.hour
                minuto = dt.minute

        except:
            pass

    if hora is None:

        try:

            hora = valor.hour
            minuto = valor.minute

        except:
            return 0

    minutos_final = hora * 60 + minuto

    limite = 15 * 60 + 20

    minutos_extra = minutos_final - limite

    if minutos_extra < 0:
        minutos_extra = 0

    return minutos_extra


# ================= LOGIN =================

@app.route("/", methods=["GET","POST"])
def login():

    if request.method == "POST":

        usuario = request.form.get("usuario")
        senha = request.form.get("senha")

        if usuario == "admin" and senha == "123":

            session["logado"] = True
            return redirect("/painel")

    return render_template("login.html")


# ================= PAINEL =================

@app.route("/painel")
def painel():

    if not session.get("logado"):
        return redirect("/")

    setores = Setor.query.order_by(Setor.total.desc()).all()

    for s in setores:

        horas = s.horas // 60
        minutos = s.horas % 60

        s.horas_formatadas = f"{horas}:{minutos:02d}"

    return render_template("painel.html", setores=setores)


# ================= CADASTRAR SETOR =================

@app.route("/cadastrar", methods=["POST"])
def cadastrar():

    if not session.get("logado"):
        return redirect("/")

    codigo = request.form.get("codigo")

    if not codigo:
        return redirect("/painel")

    codigo = codigo.strip().upper()

    existe = Setor.query.filter_by(codigo=codigo).first()

    if not existe:

        novo = Setor(codigo=codigo)

        db.session.add(novo)
        db.session.commit()

    return redirect("/painel")


# ================= EXCLUIR =================

@app.route("/excluir/<int:id>")
def excluir(id):

    if not session.get("logado"):
        return redirect("/")

    setor = Setor.query.get(id)

    if setor:

        db.session.delete(setor)
        db.session.commit()

    return redirect("/painel")


# ================= RESETAR =================

@app.route("/resetar")
def resetar():

    if not session.get("logado"):
        return redirect("/")

    setores = Setor.query.all()

    for s in setores:

        s.total = 0
        s.horas = 0

    Planilha.query.delete()

    db.session.commit()

    return redirect("/painel")


# ================= IMPORTAR PLANILHAS =================

@app.route("/upload", methods=["POST"])
def upload():

    if not session.get("logado"):
        return redirect("/")

    arquivos = request.files.getlist("arquivos")

    linhas_codigo = [2,30,58]
    linhas_peso = [23,51,79]
    linhas_hora = encontrar_linhas_hora(None)

    for arquivo in arquivos:

        if not arquivo.filename.endswith(".xlsx"):
            continue

        existe = Planilha.query.filter_by(nome=arquivo.filename).first()

        if existe:
            continue

        try:
            df = pd.read_excel(arquivo, header=None)
        except:
            continue

        for bloco in range(3):

            lc = linhas_codigo[bloco]
            lp = linhas_peso[bloco]
            lh = linhas_hora[bloco]

            for col in range(len(df.columns)):

                try:

                    codigo = df.iloc[lc,col]

                    if pd.isna(codigo):
                        continue

                    codigo = str(codigo).replace(".0","").strip().upper()

                    peso = df.iloc[lp,col]

                    if pd.isna(peso):
                        continue

                    peso = float(str(peso).replace(",","."))

                    hora_final = df.iloc[lh,col]

                    minutos = calcular_horas(hora_final)

                    setor = Setor.query.filter_by(codigo=codigo).first()

                    if setor:

                        setor.total += peso
                        setor.horas += minutos

                except:
                    continue

        nova = Planilha(nome=arquivo.filename)

        db.session.add(nova)

    db.session.commit()

    return redirect("/painel")


# ================= EXECUTAR =================

if __name__ == "__main__":

    port = int(os.environ.get("PORT",5000))

    app.run(host="0.0.0.0", port=port, debug=True)