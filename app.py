from flask import Flask, render_template, request, redirect
from openpyxl import load_workbook
from datetime import datetime

app = Flask(__name__)

ARCHIVO_EXCEL = 'asistencia.xlsx'


@app.route('/')
def inicio():
    return render_template('index.html')


@app.route('/registrar', methods=['POST'])
def registrar():
    nombre = request.form['nombre']
    codigo = request.form['codigo']
    materia = request.form['materia']

    fecha_actual = datetime.now().strftime('%Y-%m-%d')
    hora_actual = datetime.now().strftime('%H:%M:%S')

    libro = load_workbook(ARCHIVO_EXCEL)
    hoja = libro.active

    for fila in hoja.iter_rows(min_row=2, values_only=True):
        codigo_guardado = str(fila[1])
        materia_guardada = str(fila[2])
        fecha_guardada = str(fila[3])

        if (
            codigo_guardado == codigo and
            materia_guardada.lower() == materia.lower() and
            fecha_guardada == fecha_actual
        ):
            libro.close()
            return "Este estudiante ya registró asistencia en esta materia hoy."

    hoja.append([nombre, codigo, materia, fecha_actual, hora_actual])

    libro.save(ARCHIVO_EXCEL)
    libro.close()

    return redirect('/exito')


@app.route('/exito')
def exito():
    return render_template('exito.html')


if __name__ == '__main__':
    app.run(debug=True)