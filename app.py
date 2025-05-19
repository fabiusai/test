
from flask import Flask, render_template, request, send_file
import pandas as pd
import io
from formatta_excel_logica import genera_excel_format

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/genera_excel', methods=['POST'])
def genera_excel():
    file = request.files['file']
    data_inizio = request.form['data_inizio']
    data_fine = request.form['data_fine']
    campagna = request.form['campagna']

    df = pd.read_excel(file, sheet_name='Raccolta dati')

    # Conversione e filtro date
    df['Data'] = pd.to_datetime(df['Data'], dayfirst=True, errors='coerce')
    data_inizio_dt = pd.to_datetime(data_inizio, dayfirst=True)
    data_fine_dt = pd.to_datetime(data_fine, dayfirst=True)
    df = df[(df['Data'] >= data_inizio_dt) & (df['Data'] <= data_fine_dt)]

    # Filtro per tipo campagna
    if campagna == 'editoriale':
        df = df[df['Campagna'] == 'Editoriale']
    elif campagna == 'adv':
        df = df[df['Campagna'] == 'Campagna']

    # Riformattazione/accorpamenti se necessari (es. 'Facebook')
    df['Canale'] = df['Canale'].replace({
        'facebook Postepay': 'Facebook',
        'facebook PosteMobile': 'Facebook',
        'instagram Postepay': 'Instagram',
        'instagram PosteMobile': 'Instagram'
    })

    # Chiamata alla funzione per generare Excel formattato
    output = io.BytesIO()
    genera_excel_format(df, output)
    output.seek(0)

    return send_file(output, as_attachment=True, download_name='report_editoriale_formattato.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
   app.run(host='0.0.0.0', port=10000)

