"""
银行卡信息收集系统 v2 - 云端版
支持：姓名 + 银行卡号 + 银行卡照片 + Word导出
数据存储在 SQLite，照片以 base64 存入数据库
"""
from flask import Flask, render_template, request, jsonify, send_file
import sqlite3, io, datetime, os, base64
from pathlib import Path

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024  # 5MB max

DB = Path(__file__).parent / 'data.db'

def init_db():
    conn = sqlite3.connect(str(DB))
    conn.execute('''CREATE TABLE IF NOT EXISTS submissions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        card TEXT NOT NULL,
        photo TEXT DEFAULT '',
        time TEXT NOT NULL
    )''')
    conn.commit()
    conn.close()

init_db()

def get_conn():
    conn = sqlite3.connect(str(DB))
    conn.row_factory = sqlite3.Row
    return conn

# ============ 页面路由 ============

@app.route('/')
def form():
    return render_template('form.html')

@app.route('/admin')
def admin():
    return render_template('admin.html')

# ============ 数据 API ============

@app.route('/submit', methods=['POST'])
def submit():
    name = request.form.get('name', '').strip()
    card = request.form.get('card', '').strip().replace(' ', '')
    if not name or not card:
        return jsonify({'ok': False, 'msg': '姓名和卡号不能为空'})
    if len(card) < 16 or len(card) > 19:
        return jsonify({'ok': False, 'msg': '卡号位数不正确'})

    photo = ''
    if 'photo' in request.files:
        file = request.files['photo']
        if file and file.filename:
            data = file.read()
            if len(data) > 3 * 1024 * 1024:
                return jsonify({'ok': False, 'msg': '照片太大，请压缩到 3MB 以内'})
            photo = base64.b64encode(data).decode('utf-8')

    conn = get_conn()
    conn.execute(
        'INSERT INTO submissions (name, card, photo, time) VALUES (?,?,?,?)',
        [name, card, photo, datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
    )
    conn.commit()
    conn.close()
    return jsonify({'ok': True, 'msg': '提交成功！'})

@app.route('/api/data')
def get_data():
    conn = get_conn()
    rows = conn.execute('SELECT * FROM submissions ORDER BY id DESC').fetchall()
    data = [{
        'id': r['id'], 'name': r['name'], 'card': r['card'],
        'hasPhoto': bool(r['photo']), 'time': r['time']
    } for r in rows]
    conn.close()
    return jsonify(data)

@app.route('/api/photo/<int:sid>')
def get_photo(sid):
    conn = get_conn()
    row = conn.execute('SELECT photo FROM submissions WHERE id=?', [sid]).fetchone()
    conn.close()
    if row and row['photo']:
        try:
            return send_file(io.BytesIO(base64.b64decode(row['photo'])), mimetype='image/jpeg')
        except:
            pass
    return '', 404

@app.route('/api/delete/<int:sid>', methods=['POST'])
def delete(sid):
    conn = get_conn()
    conn.execute('DELETE FROM submissions WHERE id=?', [sid])
    conn.commit()
    conn.close()
    return jsonify({'ok': True})

@app.route('/api/clear', methods=['POST'])
def clear():
    conn = get_conn()
    conn.execute('DELETE FROM submissions')
    conn.commit()
    conn.close()
    return jsonify({'ok': True})

@app.route('/api/manual', methods=['POST'])
def manual():
    d = request.get_json(force=True, silent=True)
    if not d:
        return jsonify({'ok': False, 'msg': '数据格式错误'})
    name = (d.get('name') or '').strip()
    card = (d.get('card') or '').strip().replace(' ', '')
    if not name or not card:
        return jsonify({'ok': False, 'msg': '姓名和卡号不能为空'})
    if len(card) < 16 or len(card) > 19:
        return jsonify({'ok': False, 'msg': '卡号位数不正确'})
    photo = d.get('photo', '')

    conn = get_conn()
    conn.execute(
        'INSERT INTO submissions (name, card, photo, time) VALUES (?,?,?,?)',
        [name, card, photo, d.get('time') or datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
    )
    conn.commit()
    conn.close()
    return jsonify({'ok': True, 'msg': '录入成功'})

# ============ Word 导出 ============

@app.route('/export')
def export():
    conn = get_conn()
    rows = conn.execute('SELECT * FROM submissions ORDER BY id').fetchall()
    conn.close()
    if not rows:
        return '暂无数据', 404

    now = datetime.datetime.now()
    date_str = now.strftime('%Y年%m月%d日')
    fname = now.strftime('银行卡信息汇总_%Y%m%d.doc')

    has_photos = any(r['photo'] for r in rows)

    rows_html = ''
    for i, r in enumerate(rows):
        card = r['card']
        parts = [card[j:j+4] for j in range(0, len(card), 4)]
        fc = ' '.join(parts)
        nm = r['name'].replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')

        rows_html += '<tr>'
        rows_html += '<td style="border:1px solid #666;padding:8px;text-align:center">%d</td>' % (i+1)
        rows_html += '<td style="border:1px solid #666;padding:8px">%s</td>' % nm
        rows_html += '<td style="border:1px solid #666;padding:8px;font-family:\'Courier New\';letter-spacing:2px">%s</td>' % fc
        if has_photos and r['photo']:
            rows_html += '<td style="border:1px solid #666;padding:4px;text-align:center">'
            rows_html += '<img src="data:image/jpeg;base64,%s" style="max-width:140px;max-height:90px">' % r['photo']
            rows_html += '</td>'
        elif has_photos:
            rows_html += '<td style="border:1px solid #666;padding:8px;text-align:center;color:#aaa">无</td>'
        rows_html += '<td style="border:1px solid #666;padding:8px;font-size:12px;color:#888">%s</td>' % (r['time'] or '')
        rows_html += '</tr>'

    photo_header = '<th style="width:160px">银行卡照片</th>' if has_photos else ''

    html = '''<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8">
<style>
  body{font-family:"微软雅黑";}
  h1{text-align:center;color:#1a73e8;font-size:20pt;margin-bottom:10pt}
  .date{text-align:right;color:#888;font-size:11pt;margin-bottom:15pt}
  table{border-collapse:collapse;width:100%%}
  th{border:1px solid #666;padding:10px;background:#e8f0fe}
  td{border:1px solid #666;padding:8px}
  .footer{text-align:right;color:#888;font-size:10pt;margin-top:15pt}
</style></head>
<body>
<h1>银行卡信息汇总表</h1>
<p class="date">导出日期：%s</p>
<table>
  <tr><th style="width:50px">序号</th><th style="width:20%%%%">姓名</th><th>银行卡号</th>%s<th style="width:140px">提交时间</th></tr>
  %s
</table>
<p class="footer">共 %d 条记录</p>
</body></html>''' % (date_str, photo_header, rows_html, len(rows))

    bio = io.BytesIO(('\ufeff' + html).encode('utf-8-sig'))
    bio.seek(0)
    return send_file(bio, mimetype='application/msword', as_attachment=True, download_name=fname)

# ============ 启动 ============

if __name__ == '__main__':
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 1))
        ip = s.getsockname()[0]
        s.close()
    except:
        ip = '127.0.0.1'
    port = int(os.environ.get('PORT', 5000))
    print('=' * 55)
    print('  银行卡信息收集系统 v2')
    print('  管理后台: http://%s:%d/admin' % (ip, port))
    print('  填写表单: http://%s:%d/' % (ip, port))
    print('=' * 55)
    app.run(host='0.0.0.0', port=port, debug=False)
