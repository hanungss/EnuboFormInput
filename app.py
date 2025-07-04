import os
import uuid
from flask import Flask, render_template, request, redirect, url_for, send_file, flash
from enubo.submit_anggota import run_bot  # pastikan file ini ada dan berisi fungsi run_bot()

app = Flask(__name__)
app.secret_key = 'rahasia_super_aman'
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or file.filename == '':
            flash("⚠️ Pilih file Excel terlebih dahulu!")
            return redirect(request.url)

        # Simpan file dengan nama unik
        unique_id = uuid.uuid4().hex
        filename = f"{unique_id}_{file.filename}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        # Jalankan proses utama
        log_path, result_path = run_bot(filepath)

        # Baca isi log untuk ditampilkan
        with open(log_path, 'r', encoding='utf-8') as f:
            log_content = f.read()

        return render_template(
            'result.html',
            log_file=os.path.basename(log_path),
            result_file=os.path.basename(result_path),
            log_content=log_content
        )

    return render_template('index.html')

@app.route('/download/<path:filename>')
def download(filename):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return "File tidak ditemukan.", 404

if __name__ == "__main__":
    app.run(debug=True)
