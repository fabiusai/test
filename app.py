
from flask import Flask, render_template, request, send_file
import pandas as pd
import io
from formatta_excel_logica_raggruppato import genera_excel_format

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

    # Conversione date
    df['Data'] = pd.to_datetime(df['Data'], dayfirst=True, errors='coerce')
    data_inizio_dt = pd.to_datetime(data_inizio, dayfirst=True)
    data_fine_dt = pd.to_datetime(data_fine, dayfirst=True)

    # Filtro per data
    df = df[(df['Data'] >= data_inizio_dt) & (df['Data'] <= data_fine_dt)]
    print("Righe dopo filtro:", len(df))
print("Campagne uniche:", df['Campagna'].unique())
print("Canali unici:", df['Canale'].unique())

    # Filtro per tipo campagna
    if campagna == 'editoriale':
        df = df[df['Campagna'] == 'Editoriale']
    elif campagna == 'adv':
        df = df[df['Campagna'] == 'Campagna']

    # Uniforma i canali
    df['Canale'] = df['Canale'].replace({
        'facebook Postepay': 'Facebook',
        'facebook PosteMobile': 'Facebook',
        'instagram Postepay': 'Instagram',
        'instagram PosteMobile': 'Instagram'
    })

    # Rimuove righe senza Label o Canale
    df = df.dropna(subset=['Label', 'Canale'])

    # Selezione metriche da aggregare
    metriche = ['Interazioni', 'Visualizzazioni', 'Persone raggiunte', 'Click', 'Like e reazioni', 'Condivisioni', 'Commenti']

    # Aggregazione
    df_grouped = df.groupby(['Label', 'Argomento', 'Canale'])[metriche].sum().reset_index()

    # Pivot per canali come colonne
    pivot_df = df_grouped.pivot_table(index=['Label', 'Argomento'], columns='Canale', values=metriche, aggfunc='sum', fill_value=0)

    # Riordina i canali
    ordered_channels = ['Facebook', 'Instagram', 'LinkedIn', 'Twitter', 'YouTube']
    pivot_df = pivot_df.reindex(columns=pd.MultiIndex.from_product([metriche, ordered_channels]), fill_value=0)

    # Appiattisce le colonne multi-index
    pivot_df.columns = ['{} - {}'.format(m, c) for m, c in pivot_df.columns]
    pivot_df = pivot_df.reset_index()

    output = io.BytesIO()
    genera_excel_format(pivot_df, output)
    output.seek(0)

    return send_file(output, as_attachment=True, download_name='report_editoriale_formattato.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
