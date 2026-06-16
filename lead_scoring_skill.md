# SKILL: AI Lead Scoring & Automation - Bất Động Sản

## 1. MỤC TIÊU

Tự động hóa quy trình chấm điểm khách hàng tiềm năng (lead scoring) trong ngành Bất Động Sản bằng AI, kết hợp giao diện Web App cho phép con người kiểm duyệt và xuất kết quả ra Excel.

---

## 2. KIẾN TRÚC TỔNG THỂ

```
Google Sheets (Nguồn dữ liệu)
        ↓ (Google Sheets API)
Python Backend (AI Scoring Engine)
        ↓ (Gemini AI / Rule-based)
Web App (Human-in-the-loop Review UI)
        ↓ (Export)
Excel Output File
```

---

## 3. CẤU TRÚC DỮ LIỆU ĐẦU VÀO (Google Sheets)

**Link:** `https://docs.google.com/spreadsheets/d/16tCAf_qqtgYZxoumYQKMEOdBhKE0wg5A/edit?gid=1542775777`

**Các cột dự kiến:**
| Cột | Mô tả |
|-----|-------|
| ID | Mã khách hàng |
| Tên khách hàng | Họ tên |
| Số điện thoại | SĐT liên lạc |
| Ghi chú / Nhu cầu | Nội dung mô tả nhu cầu (INPUT chính để AI chấm điểm) |
| Trạng thái | Trạng thái hiện tại của lead |

---

## 4. BỘ QUY TẮC CHẤM ĐIỂM

### 4.1 CỘNG 50 ĐIỂM — Khách hàng VIP / Siêu tiềm năng

AI nhận diện các từ khóa và ngữ cảnh sau trong cột **Ghi chú / Nhu cầu**:

| # | Tiêu chí | Từ khóa / Dấu hiệu |
|---|----------|---------------------|
| 1 | **Ngân sách lớn** | Số tiền ≥ 20 tỷ, "tài chính mạnh", "không thành vấn đề" |
| 2 | **Loại hình cao cấp** | "Biệt thự đơn lập", "Penthouse", "Shophouse mặt đường lớn", "Quỹ đất công nghiệp", "Sàn văn phòng diện tích lớn" |
| 3 | **Vị trí đắc địa** | "Quận 1", "Ven sông", "Vinhomes Ocean Park", "Phú Mỹ Hưng" |
| 4 | **Đối tượng VIP** | "Chủ doanh nghiệp", "Nhà đầu tư chuyên nghiệp", "Mua sỉ", "Mua số lượng lớn" |
| 5 | **Cấp thiết & Minh bạch** | "Pháp lý chuẩn 100%", "Sổ hồng riêng", "Muốn gặp trực tiếp chủ đầu tư" |

**→ Điểm ban đầu: 50 | Sau khi cộng: 100**

---

### 4.2 TRỪ 50 ĐIỂM — Khách hàng Rác / Không tiềm năng

| # | Tiêu chí | Từ khóa / Dấu hiệu |
|---|----------|---------------------|
| 1 | **Yêu cầu phi thực tế** | Nhà Quận 1 giá 1-2 tỷ, nhà có hồ bơi giá vài trăm triệu |
| 2 | **Không có nhu cầu** | "Nhầm số", "Không có nhu cầu", "Dữ liệu cũ", "Nhầm ngành" |
| 3 | **Không thiện chí** | "Hỏi giá cho vui", "Chưa có ý định mua", "Thái độ không hợp tác" |
| 4 | **Spam / Quảng cáo** | "Bảo hiểm", "Vay vốn", "Mời chào dịch vụ" |
| 5 | **Liên lạc lỗi** | "Thuê bao", "Gọi nhiều lần không bắt máy", "Không phản hồi Zalo" |

**→ Điểm ban đầu: 50 | Sau khi trừ: 0**

---

### 4.3 GIỮ NGUYÊN / CỘNG ÍT — Khách hàng Trung bình

| Dấu hiệu | Điểm |
|----------|------|
| Tìm chung cư, nhà phố tầm trung (3–10 tỷ) | 50 |
| Cần vay ngân hàng, đang cân nhắc chính sách | 40–50 |
| Nhu cầu thực nhưng cần tư vấn thêm | 45–55 |

---

### 4.4 PHÂN LOẠI KẾT QUẢ

| Điểm | Phân loại | Màu sắc | Hành động |
|------|-----------|---------|-----------|
| 80–100 | 🔥 VIP / Siêu tiềm năng | Đỏ | Ưu tiên gọi ngay |
| 50–79 | ✅ Tiềm năng | Xanh lá | Chăm sóc bình thường |
| 20–49 | ⚠️ Cần theo dõi | Vàng | Theo dõi thêm |
| 0–19 | ❌ Rác / Loại | Xám | Không cần chăm sóc |

---

## 5. CÁC BƯỚC THỰC HIỆN

### BƯỚC 1: Lấy dữ liệu từ Google Sheets

```python
# File: fetch_data.py
import gspread
from google.oauth2.service_account import Credentials

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SHEET_ID = '16tCAf_qqtgYZxoumYQKMEOdBhKE0wg5A'
SHEET_GID = '1542775777'

def fetch_leads():
    creds = Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SHEET_ID).get_worksheet_by_id(int(SHEET_GID))
    records = sheet.get_all_records()
    return records
```

---

### BƯỚC 2: Engine chấm điểm AI

```python
# File: scoring_engine.py
import google.generativeai as genai

genai.configure(api_key="YOUR_GEMINI_API_KEY")
model = genai.GenerativeModel('gemini-1.5-flash')

SCORING_PROMPT = """
Bạn là chuyên gia phân tích lead bất động sản. Dựa vào nội dung ghi chú khách hàng sau,
hãy chấm điểm tiềm năng từ 0-100 và phân loại khách hàng.

Quy tắc chấm điểm:
- CỘNG 50 điểm (VIP): ngân sách ≥20 tỷ, biệt thự/penthouse/shophouse cao cấp, vị trí đắc địa
  (Quận 1/Phú Mỹ Hưng/Vinhomes), chủ doanh nghiệp/nhà đầu tư, yêu cầu pháp lý chuẩn
- TRỪ 50 điểm (Rác): giá phi thực tế, nhầm số/không nhu cầu, hỏi cho vui, spam dịch vụ,
  không liên lạc được
- GIỮ NGUYÊN: chung cư/nhà phố 3-10 tỷ, cần vay ngân hàng, cân nhắc chính sách

Ghi chú khách hàng: "{note}"

Trả lời JSON:
{{
  "score": <số 0-100>,
  "category": "<VIP|Tiềm năng|Cần theo dõi|Rác>",
  "reason": "<lý do ngắn gọn 1-2 câu>",
  "action": "<hành động đề xuất>"
}}
"""

def score_lead(note: str) -> dict:
    if not note or note.strip() == "":
        return {"score": 0, "category": "Rác", "reason": "Không có ghi chú", "action": "Loại bỏ"}
    
    prompt = SCORING_PROMPT.format(note=note)
    response = model.generate_content(prompt)
    
    import json, re
    text = response.text
    json_match = re.search(r'\{.*\}', text, re.DOTALL)
    if json_match:
        return json.loads(json_match.group())
    return {"score": 50, "category": "Cần theo dõi", "reason": "Không xác định được", "action": "Xem xét thêm"}


def process_all_leads(leads: list) -> list:
    results = []
    for lead in leads:
        note = lead.get('Ghi chú', '') or lead.get('Nhu cầu', '') or lead.get('Note', '')
        scoring = score_lead(str(note))
        results.append({**lead, **scoring})
    return results
```

---

### BƯỚC 3: Web App Human-in-the-loop

```python
# File: app.py
from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import json, os
from fetch_data import fetch_leads
from scoring_engine import process_all_leads

app = Flask(__name__)
DATA_FILE = 'scored_leads.json'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/fetch-and-score', methods=['POST'])
def fetch_and_score():
    leads = fetch_leads()
    scored = process_all_leads(leads)
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(scored, f, ensure_ascii=False, indent=2)
    return jsonify({"status": "success", "count": len(scored), "data": scored})

@app.route('/api/leads', methods=['GET'])
def get_leads():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return jsonify(json.load(f))
    return jsonify([])

@app.route('/api/update-lead', methods=['POST'])
def update_lead():
    """Human-in-the-loop: cập nhật trạng thái/điểm từ reviewer"""
    data = request.json
    lead_id = data.get('id')
    with open(DATA_FILE, 'r', encoding='utf-8') as f:
        leads = json.load(f)
    for lead in leads:
        if str(lead.get('ID')) == str(lead_id):
            lead['score'] = data.get('score', lead['score'])
            lead['category'] = data.get('category', lead['category'])
            lead['status_reviewed'] = True
            lead['reviewer_note'] = data.get('reviewer_note', '')
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(leads, f, ensure_ascii=False, indent=2)
    return jsonify({"status": "updated"})

@app.route('/api/export-excel', methods=['GET'])
def export_excel():
    with open(DATA_FILE, 'r', encoding='utf-8') as f:
        leads = json.load(f)
    df = pd.DataFrame(leads)
    output_file = 'lead_scoring_result.xlsx'
    df.to_excel(output_file, index=False)
    return send_file(output_file, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
```

---

### BƯỚC 4: Giao diện Web App (HTML)

```html
<!-- File: templates/index.html -->
<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <title>AI Lead Scoring - BĐS</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: 'Segoe UI', sans-serif; background: #0f172a; color: #e2e8f0; }
    header { background: linear-gradient(135deg, #1e3a5f, #0ea5e9); padding: 20px 40px; }
    header h1 { font-size: 1.8rem; }
    .controls { padding: 20px 40px; display: flex; gap: 12px; flex-wrap: wrap; }
    .btn { padding: 10px 24px; border: none; border-radius: 8px; cursor: pointer; font-size: 0.95rem; font-weight: 600; transition: all 0.2s; }
    .btn-primary { background: #0ea5e9; color: white; }
    .btn-primary:hover { background: #0284c7; }
    .btn-success { background: #22c55e; color: white; }
    .btn-success:hover { background: #16a34a; }
    .stats { display: flex; gap: 16px; padding: 0 40px 20px; flex-wrap: wrap; }
    .stat-card { background: #1e293b; border-radius: 12px; padding: 16px 24px; min-width: 140px; text-align: center; }
    .stat-card .num { font-size: 2rem; font-weight: 700; }
    .stat-card .label { font-size: 0.8rem; color: #94a3b8; margin-top: 4px; }
    .red { color: #ef4444; } .green { color: #22c55e; } .yellow { color: #f59e0b; } .gray { color: #94a3b8; }
    table { width: calc(100% - 80px); margin: 0 40px 40px; border-collapse: collapse; background: #1e293b; border-radius: 12px; overflow: hidden; }
    th { background: #0f172a; padding: 14px 12px; text-align: left; font-size: 0.85rem; color: #94a3b8; text-transform: uppercase; }
    td { padding: 12px; border-bottom: 1px solid #334155; font-size: 0.9rem; }
    tr:hover { background: #263548; }
    .badge { padding: 4px 10px; border-radius: 20px; font-size: 0.78rem; font-weight: 600; }
    .badge-vip { background: #7f1d1d; color: #fca5a5; }
    .badge-potential { background: #14532d; color: #86efac; }
    .badge-watch { background: #78350f; color: #fcd34d; }
    .badge-trash { background: #1e293b; color: #94a3b8; border: 1px solid #475569; }
    input[type=number], select, input[type=text] { background: #0f172a; border: 1px solid #475569; color: #e2e8f0; padding: 4px 8px; border-radius: 6px; width: 100%; }
    .loading { text-align: center; padding: 40px; color: #94a3b8; }
  </style>
</head>
<body>
<header>
  <h1>🏢 AI Lead Scoring — Bất Động Sản</h1>
  <p style="color:#bae6fd;margin-top:6px">Human-in-the-loop Review Dashboard</p>
</header>

<div class="controls">
  <button class="btn btn-primary" onclick="fetchAndScore()">🔄 Lấy dữ liệu & Chấm điểm AI</button>
  <button class="btn btn-success" onclick="exportExcel()">📥 Xuất Excel</button>
</div>

<div class="stats" id="stats"></div>

<table id="leads-table">
  <thead>
    <tr>
      <th>ID</th><th>Tên KH</th><th>SĐT</th><th>Ghi chú</th>
      <th>Điểm AI</th><th>Phân loại</th><th>Lý do</th>
      <th>Điều chỉnh</th><th>Ghi chú reviewer</th><th>Lưu</th>
    </tr>
  </thead>
  <tbody id="leads-body"><tr><td colspan="10" class="loading">Nhấn "Lấy dữ liệu & Chấm điểm AI" để bắt đầu</td></tr></tbody>
</table>

<script>
let leadsData = [];

async function fetchAndScore() {
  document.getElementById('leads-body').innerHTML = '<tr><td colspan="10" class="loading">⏳ Đang lấy dữ liệu và chấm điểm AI...</td></tr>';
  const res = await fetch('/api/fetch-and-score', { method: 'POST' });
  const data = await res.json();
  leadsData = data.data;
  renderTable(leadsData);
}

async function loadLeads() {
  const res = await fetch('/api/leads');
  leadsData = await res.json();
  if (leadsData.length > 0) renderTable(leadsData);
}

function renderTable(leads) {
  updateStats(leads);
  const tbody = document.getElementById('leads-body');
  tbody.innerHTML = leads.map((l, i) => {
    const badgeClass = l.category === 'VIP' ? 'badge-vip' : l.category === 'Tiềm năng' ? 'badge-potential' : l.category === 'Cần theo dõi' ? 'badge-watch' : 'badge-trash';
    return `<tr>
      <td>${l.ID || i+1}</td>
      <td>${l['Tên khách hàng'] || l['Ten'] || '—'}</td>
      <td>${l['Số điện thoại'] || l['SDT'] || '—'}</td>
      <td style="max-width:200px;font-size:0.8rem">${(l['Ghi chú'] || l['Note'] || '').substring(0,100)}...</td>
      <td><strong style="font-size:1.1rem">${l.score}</strong></td>
      <td><span class="badge ${badgeClass}">${l.category}</span></td>
      <td style="font-size:0.8rem;color:#94a3b8">${l.reason || ''}</td>
      <td>
        <input type="number" id="score_${i}" value="${l.score}" min="0" max="100" style="width:60px">
        <select id="cat_${i}" style="margin-top:4px">
          <option ${l.category==='VIP'?'selected':''}>VIP</option>
          <option ${l.category==='Tiềm năng'?'selected':''}>Tiềm năng</option>
          <option ${l.category==='Cần theo dõi'?'selected':''}>Cần theo dõi</option>
          <option ${l.category==='Rác'?'selected':''}>Rác</option>
        </select>
      </td>
      <td><input type="text" id="note_${i}" placeholder="Ghi chú reviewer..."></td>
      <td><button class="btn btn-primary" onclick="saveEdit(${i})" style="padding:6px 12px;font-size:0.8rem">💾</button></td>
    </tr>`;
  }).join('');
}

function updateStats(leads) {
  const vip = leads.filter(l => l.category === 'VIP').length;
  const pot = leads.filter(l => l.category === 'Tiềm năng').length;
  const watch = leads.filter(l => l.category === 'Cần theo dõi').length;
  const trash = leads.filter(l => l.category === 'Rác').length;
  document.getElementById('stats').innerHTML = `
    <div class="stat-card"><div class="num red">${vip}</div><div class="label">🔥 VIP</div></div>
    <div class="stat-card"><div class="num green">${pot}</div><div class="label">✅ Tiềm năng</div></div>
    <div class="stat-card"><div class="num yellow">${watch}</div><div class="label">⚠️ Theo dõi</div></div>
    <div class="stat-card"><div class="num gray">${trash}</div><div class="label">❌ Rác</div></div>
    <div class="stat-card"><div class="num" style="color:#0ea5e9">${leads.length}</div><div class="label">📋 Tổng</div></div>
  `;
}

async function saveEdit(i) {
  const lead = leadsData[i];
  await fetch('/api/update-lead', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({
      id: lead.ID || i,
      score: parseInt(document.getElementById(`score_${i}`).value),
      category: document.getElementById(`cat_${i}`).value,
      reviewer_note: document.getElementById(`note_${i}`).value
    })
  });
  alert('✅ Đã lưu!');
}

function exportExcel() {
  window.open('/api/export-excel', '_blank');
}

loadLeads();
</script>
</body>
</html>
```

---

### BƯỚC 5: Xuất Excel

```python
# File: export_excel.py (gọi độc lập hoặc qua API)
import pandas as pd
import json

def export_to_excel(data_file='scored_leads.json', output_file='lead_scoring_result.xlsx'):
    with open(data_file, 'r', encoding='utf-8') as f:
        leads = json.load(f)
    
    df = pd.DataFrame(leads)
    
    # Sắp xếp theo điểm giảm dần
    if 'score' in df.columns:
        df = df.sort_values('score', ascending=False)
    
    # Tô màu theo phân loại
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Lead Scoring', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Lead Scoring']
        
        colors = {'VIP': '#FF4444', 'Tiềm năng': '#22C55E', 'Cần theo dõi': '#F59E0B', 'Rác': '#94A3B8'}
        for row_idx, (_, row) in enumerate(df.iterrows(), start=1):
            cat = row.get('category', '')
            color = colors.get(cat, '#FFFFFF')
            fmt = workbook.add_format({'bg_color': color, 'font_color': 'white', 'bold': True})
            worksheet.write(row_idx, df.columns.get_loc('category'), cat, fmt)
    
    print(f"✅ Xuất thành công: {output_file}")
    return output_file
```

---

## 6. CẤU TRÚC THƯ MỤC DỰ ÁN

```
lead_scoring/
├── app.py                    # Flask Web App chính
├── fetch_data.py             # Lấy dữ liệu từ Google Sheets
├── scoring_engine.py         # Engine AI chấm điểm (Gemini)
├── export_excel.py           # Xuất kết quả ra Excel
├── credentials.json          # Google Service Account key (KHÔNG commit)
├── scored_leads.json         # Kết quả chấm điểm (cache local)
├── lead_scoring_result.xlsx  # File Excel output
├── requirements.txt          # Dependencies
└── templates/
    └── index.html            # Giao diện Web App
```

---

## 7. DEPENDENCIES

```txt
# requirements.txt
flask>=3.0.0
gspread>=6.0.0
google-auth>=2.0.0
google-generativeai>=0.8.0
pandas>=2.0.0
openpyxl>=3.1.0
xlsxwriter>=3.1.0
```

**Cài đặt:**
```bash
pip install -r requirements.txt
```

---

## 8. HƯỚNG DẪN CHẠY

```bash
# Bước 1: Đặt credentials.json vào thư mục gốc
# Bước 2: Set API key
set GEMINI_API_KEY=your_key_here

# Bước 3: Chạy app
python app.py

# Bước 4: Mở trình duyệt
# http://localhost:5000
```

---

## 9. LUỒNG XỬ LÝ HUMAN-IN-THE-LOOP

```
1. [Reviewer] Nhấn "Lấy dữ liệu & Chấm điểm AI"
        ↓
2. [AI] Đọc Google Sheets → Chấm điểm từng khách hàng (Gemini)
        ↓
3. [Web App] Hiển thị bảng kết quả với điểm, phân loại, lý do
        ↓
4. [Reviewer] Kiểm tra, điều chỉnh điểm/phân loại nếu cần → Nhấn 💾
        ↓
5. [Reviewer] Nhấn "Xuất Excel" → Tải file kết quả đã được xét duyệt
```

---

## 10. KPI & OUTPUT MONG ĐỢI

| Metric | Mô tả |
|--------|-------|
| Tổng leads xử lý | Toàn bộ dữ liệu từ Google Sheets |
| % VIP | Tỷ lệ khách hàng điểm 80–100 |
| % Rác | Tỷ lệ cần loại bỏ |
| Thời gian chấm điểm | < 2 giây/lead |
| File output | `lead_scoring_result.xlsx` có màu sắc phân loại |

---

*Skill được xây dựng cho MINDX Gravity - Lesson 7*
*Ngày tạo: 2026-06-16*
