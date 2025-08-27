import io, os, uuid, zipfile
from urllib.parse import quote
from flask import Flask, render_template, request, send_file, redirect, url_for, abort

from sds_parser import parse_sds
from ra_writer import make_excel_bytes, render_pdf_bytes

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB/req

# メモリ上の簡易ストア（短期ダウンロード用）
STORE = {}  # token -> {"bytes": b, "mimetype": str, "filename": str}
JOBS = {}   # job_id -> list[token]


def stash_file(buf: bytes, mimetype: str, filename: str) -> str:
    token = uuid.uuid4().hex
    STORE[token] = {"bytes": buf, "mimetype": mimetype, "filename": filename}
    return token


def content_disposition(filename: str, fallback: str = "download") -> str:
    # ASCIIフォールバック + UTF-8 (RFC5987)
    return f"attachment; filename={fallback}; filename*=UTF-8''{quote(filename)}"


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    files = request.files.getlist('files')
    if not files:
        return redirect(url_for('index'))

    out_xlsx = request.form.get('out_xlsx') == 'on'
    out_pdf  = request.form.get('out_pdf') == 'on'
    if not (out_xlsx or out_pdf):
        out_xlsx = True  # デフォルトExcel

    tokens = []
    for f in files:
        raw = f.read()
        meta = parse_sds(raw, original_filename=f.filename or "")

        created = []
        if out_xlsx:
            xlsx_bytes, xlsx_name = make_excel_bytes(meta)
            t = stash_file(xlsx_bytes, \
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', xlsx_name)
            tokens.append(t)
            created.append((t, xlsx_name, 'xlsx'))

        if out_pdf:
            pdf_bytes, pdf_name = render_pdf_bytes(meta)  # wkhtmltopdf
            t = stash_file(pdf_bytes, 'application/pdf', pdf_name)
            tokens.append(t)
            created.append((t, pdf_name, 'pdf'))

    job_id = uuid.uuid4().hex
    JOBS[job_id] = tokens

    # 個別/一括ダウンロードをresult画面で提示
    # createdは最後のファイルのペアしかないため、再収集
    items = []
    for t in tokens:
        info = STORE[t]
        items.append({
            'token': t,
            'filename': info['filename'],
            'ext': info['filename'].split('.')[-1].lower(),
        })

    return render_template('result.html', items=items, job_id=job_id)


@app.route('/download/<token>')
def download(token):
    info = STORE.get(token)
    if not info:
        abort(404)
    buf = io.BytesIO(info['bytes'])
    resp = send_file(
        buf,
        mimetype=info['mimetype'],
        as_attachment=True,
        download_name=info['filename']
    )
    # Edge/古いUA向け：明示的に Content-Disposition を上書き
    resp.headers['Content-Disposition'] = content_disposition(info['filename'], fallback='file')
    resp.headers['X-Content-Type-Options'] = 'nosniff'
    return resp


@app.route('/zip/<job_id>')
def zip_all(job_id):
    tokens = JOBS.get(job_id)
    if not tokens:
        abort(404)
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, 'w', zipfile.ZIP_DEFLATED) as z:
        for t in tokens:
            info = STORE.get(t)
            if not info:
                continue
            z.writestr(info['filename'], info['bytes'])
    mem.seek(0)
    filename = 'リスクアセスメント一式.zip'
    resp = send_file(mem, mimetype='application/zip', as_attachment=True, download_name=filename)
    resp.headers['Content-Disposition'] = content_disposition(filename, fallback='bundle.zip')
    resp.headers['X-Content-Type-Options'] = 'nosniff'
    return resp


if __name__ == '__main__':
    port = int(os.getenv('PORT', '10000'))
    app.run(host='0.0.0.0', port=port, debug=False)
