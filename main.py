from flask import Flask, request, send_file
import pandas as pd
import tempfile
from datetime import datetime
import os
from werkzeug.utils import secure_filename
import traceback
from flask_cors import CORS

app = Flask(__name__)
CORS(app, origins=["https://tigullio-excel-flow.lovable.app"])  # <--- AGGIUNGI QUESTO

app = Flask(__name__)

@app.route("/process_excel", methods=["POST"])
def process_excel():
    try:
        if 'file' not in request.files:
            app.logger.error("Nessun file trovato nella richiesta")
            return {"error": "Nessun file trovato"}, 400

        file = request.files['file']
        filename = secure_filename(file.filename)
        app.logger.info(f"Ricevuto file: {filename}")

        # Legge il file Excel prendendo l'intestazione dalla riga 7 (header=6)
        df = pd.read_excel(file, header=6)
        app.logger.info(f"Colonne trovate: {df.columns.tolist()}")

        # Filtra solo le righe con Data, Descrizione e Importo validi
        df = df[['Data', 'Descrizione', 'Importo']].dropna(subset=['Data', 'Descrizione', 'Importo'])

        # Converte le date da MM/DD/YYYY a DD/MM/YYYY
        def convert_date(date):
            if isinstance(date, str):
                return datetime.strptime(date, "%m/%d/%Y").strftime("%d/%m/%Y")
            return date.strftime("%d/%m/%Y")

        df['Data Contabile'] = df['Data'].apply(convert_date)
        df['Data'] = df['Data'].apply(convert_date)
        df['Divisa'] = 'EUR'
        df['Causale / Descrizione'] = df['Descrizione']

        # Riordina e seleziona le colonne finali
        final_df = df[['Data Contabile', 'Data', 'Importo', 'Divisa', 'Causale / Descrizione']]

        # Prepara il nome del file di output con prefisso 'CODIFICATO'
        output_name = f"CODIFICATO_{filename}"
        output_path = os.path.join(tempfile.gettempdir(), output_name)

        final_df.to_excel(output_path, index=False)
        app.logger.info(f"File creato: {output_path}")

        return send_file(output_path, as_attachment=True, download_name=output_name)

    except Exception as e:
        app.logger.error("Errore durante l'elaborazione del file:")
        app.logger.error(traceback.format_exc())
        return {"error": str(e)}, 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
