"""
render_app.py — Application Flask pour Render.com
Synchronisé avec generate_pptx.py (Version Multi-Pays)
"""

import os, io, uuid, logging, tempfile
from flask import Flask, request, render_template_string, send_file, jsonify
# On importe la fonction de génération du fichier PPTX
from generate_pptx import generate_report

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

# Dossier temporaire pour stocker les fichiers durant la session
TMP_DIR = tempfile.gettempdir()

# ── HTML de l'interface (Design RECMA) ────────────────────────────────────────
HTML = """
<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>NBB Generator · RECMA</title>
  <style>
    body { font-family: sans-serif; background: #f4f7f6; padding: 50px; text-align: center; }
    .card { background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); display: inline-block; }
    h1 { color: #2D5C54; }
    input { margin: 20px 0; }
    button { background: #2D5C54; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; }
    #status { margin-top: 20px; font-weight: bold; }
  </style>
</head>
<body>
  <div class="card">
    <h1>NBB Report Generator</h1>
    <p>Uploadez votre fichier Excel (Mexique, HK, etc.)</p>
    <input type="file" id="excelFile" accept=".xlsx">
    <br>
    <button onclick="upload()">Générer le rapport</button>
    <div id="status"></div>
  </div>

  <script>
    async function upload() {
      const fileInput = document.getElementById('excelFile');
      if (!fileInput.files[0]) return alert("Sélectionnez un fichier");

      const status = document.getElementById('status');
      status.innerText = "⏳ Génération en cours... (cela peut prendre 30s)";
      
      const formData = new FormData();
      formData.append('file', fileInput.files[0]);

      try {
        const resp = await fetch('/process', { method: 'POST', body: formData });
        const data = await resp.json();
        
        if (data.error) {
          status.innerHTML = "❌ Erreur : " + data.error;
        } else {
          status.innerHTML = `✅ Terminé ! <br><br> 
            <a href="${data.url}" style="color:#2D5C54; font-weight:bold;">📥 Télécharger le PPTX</a>`;
        }
      } catch (e) {
        status.innerText = "❌ Erreur de connexion au serveur";
      }
    }
  </script>
</body>
</html>
"""

# ── Cache mémoire pour les téléchargements ────────────────────────────────────
_file_cache = {} # token -> (bytes, filename)

@app.route('/')
def index():
    return render_template_string(HTML)

@app.route('/process', methods=['POST'])
def process():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Aucun fichier reçu'}), 400
        
        file = request.files['file']
        # Sauvegarde temporaire du fichier Excel reçu
        input_path = os.path.join(TMP_DIR, "input_data.xlsx")
        file.save(input_path)

        # APPEL de la fonction de génération (C'est ici que la magie opère)
        # On passe le chemin du fichier Excel à generate_pptx.py
        output_pptx = generate_report(input_path)

        # Lecture du résultat pour le mettre en cache
        with open(output_pptx, 'rb') as f:
            data = f.read()
        
        token = str(uuid.uuid4())
        filename = os.path.basename(output_pptx)
        _file_cache[token] = (data, filename)

        return jsonify({'url': f'/download/{token}'})

    except Exception as e:
        app.logger.error(f"Erreur process: {e}", exc_info=True)
        return jsonify({'error': str(e)}), 500

@app.route('/download/<token>')
def download(token):
    if token not in _file_cache:
        return "Lien expiré", 404
    data, filename = _file_cache[token]
    return send_file(io.BytesIO(data), 
                     mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                     as_attachment=True, 
                     download_name=filename)

if __name__ == '__main__':
    # Render utilise gunicorn, mais pour le test local :
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 10000)))
