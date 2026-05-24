from flask import Flask, render_template, request, redirect, send_file
from openpyxl import load_workbook
from datetime import datetime
import qrcode
import socket

app = Flask(__name__)

ARCHIVO_EXCEL = 'asistencia.xlsx'
def obtener_ip_local():
    hostname = socket.gethostname()
    ip_local = socket.gethostbyname(hostname)
    return ip_local

@app.route('/')
def inicio():
    return render_template('index.html')

@app.route('/registrar', methods=['POST'])
def registrar():
    nombre = request.form['nombre'].strip()
    codigo = request.form['codigo'].strip()
    materia = request.form['materia'].strip()

    fecha_actual = datetime.now().strftime('%Y-%m-%d')
    hora_actual = datetime.now().strftime('%H:%M:%S')

    libro = load_workbook(ARCHIVO_EXCEL)
    hoja = libro.active

    for fila in hoja.iter_rows(values_only=True):
        if fila[0] is None:
            continue

        codigo_guardado = str(fila[1]).strip()
        materia_guardada = str(fila[2]).strip().lower()
        fecha_guardada = str(fila[3]).strip()

        if (
            codigo_guardado == codigo and
            materia_guardada == materia.lower() and
            fecha_guardada == fecha_actual
        ):
            libro.close()
            return render_template('duplicado.html')

    hoja.append([nombre, codigo, materia, fecha_actual, hora_actual])

    libro.save(ARCHIVO_EXCEL)
    libro.close()

    return redirect('/exito')

@app.route('/exito')
def exito():
    return render_template('exito.html')

@app.route('/admin')
def admin():
    libro = load_workbook(ARCHIVO_EXCEL)
    hoja = libro.active

    registros = []

    for fila in hoja.iter_rows(min_row=2, values_only=True):
        registros.append({
            'nombre': fila[0],
            'codigo': fila[1],
            'materia': fila[2],
            'fecha': fila[3],
            'hora': fila[4]
        })

    libro.close()

    return render_template('admin.html', registros=registros)


@app.route('/descargar')
def descargar():
    return send_file(ARCHIVO_EXCEL, as_attachment=True)

@app.route('/qr')
def mostrar_qr():
    ip_local = obtener_ip_local()
    url_formulario = f'http://{ip_local}:5000/'

    qr = qrcode.make(url_formulario)
    qr.save('static/qr_asistencia.png')

    return render_template('qr.html', url_formulario=url_formulario)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)