from flask import Flask, request, send_file
import pandas as pd
import tempfile
from datetime import datetime
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

@app.route("/process_excel", methods=["POST"])
def process_excel():
    if 'file' not in request.files:
        return {"error": "Nessun file trovato"}, 400

    file = request.files['file']
    filename = secure_filename(file.filename)

    df = pd.read_excel(file, skiprows=7)
    df = df[['Data', 'Descrizione', 'Importo']].dropna(subset=['Data', 'Descrizione', 'Importo'])

    def convert_date(date):
        if isinstance(date, str):
            return datetime.strptime(date, "%m/%d/%Y").strftime("%d/%m/%Y")
        return date.strftime("%d/%m/%Y")

    df['Data Contabile'] = df['Data'].apply(convert_date)
    df['Data'] = df['Data'].apply(convert_date)
    df['Divisa'] = 'EUR'
    df['Causale / Descrizione'] = df['Descrizione']

    final_df = df[['Data Contabile', 'Data', 'Importo', 'Divisa', 'Causale / Descrizione']]

    output_name = f"CODIFICATO_{filename}"
    output_path = os.path.join(tempfile.gettempdir(), output_name)

    final_df.to_excel(output_path, index=False)

    return send_file(output_path, as_attachment=True, download_name=output_name)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
