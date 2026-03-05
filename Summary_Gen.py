import pandas as pd

# ==================== 設定區 ====================
EXCEL_FILE = 'GCIS DBaaS Summary2_0123.xls'          # ← 你的 Excel 檔名
OUTPUT_HTML = 'dbaas_list.html'                       # 輸出的完整 HTML 檔名
SHEET_NAME = 'adhoc_list_table'

# ==================== 讀取 Excel ====================
df = pd.read_excel(EXCEL_FILE,
                   sheet_name=SHEET_NAME,
                   skiprows=0,                        # 重要：不要跳過標題行
                   engine='xlrd')

# 清除欄位名稱前後空格
df.columns = df.columns.str.strip()

# 選需要的欄位（根據你實際欄位名稱調整，如果有空格也要加）
columns_needed = ['GCISID', 'Site', 'VMName', 'DBaaSType', 'DBaaSStatus', 'MgtIP of VM']

# 檢查欄位是否存在
missing = [col for col in columns_needed if col not in df.columns]
if missing:
    print("缺少欄位，請檢查 Excel 標題：", missing)
    print("目前欄位：", df.columns.tolist())
    exit()

df = df[columns_needed].dropna(subset=['GCISID', 'MgtIP of VM'])

# 按 GCISID 分組，並排序 Site (P1→P2→T1→T2)
sort_order = {'P1': 0, 'P2': 1, 'T1': 2, 'T2': 3}
df['sort_key'] = df['Site'].map(sort_order).fillna(999)
grouped = df.sort_values(['GCISID', 'sort_key']).groupby('GCISID')

# ==================== 產生 HTML 內容 ====================
accordion_html = ''
for idx, (gcisid, group) in enumerate(grouped):
    collapse_id = gcisid.lower().replace(' ', '-').replace('_', '-')
    is_first = idx == 0

    accordion_html += f'''
<div class="accordion-item">
    <h2 class="accordion-header">
        <button class="accordion-button {'collapsed' if not is_first else ''}" type="button" 
                data-bs-toggle="collapse" data-bs-target="#collapse-{collapse_id}">
            {gcisid}
        </button>
    </h2>
    <div id="collapse-{collapse_id}" class="accordion-collapse collapse {'show' if is_first else ''}">
        <div class="accordion-body p-0">
            <div class="table-responsive">
                <table class="table table-sm table-bordered mb-0 text-center align-middle">
                    <thead class="table-light">
                        <tr>
                            <th>Site</th>
                            <th>VM Name</th>
                            <th>DBaaS Type</th>
                            <th>Status</th>
                            <th>Mgt IP</th>
                        </tr>
                    </thead>
                    <tbody>
'''

    for _, row in group.iterrows():
        ip = str(row['MgtIP of VM']).strip()
        status_class = 'status-active' if str(row['DBaaSStatus']).strip() == 'InUse' else 'text-danger'
        accordion_html += f'''
                        <tr>
                            <td>{row['Site']}</td>
                            <td>{row['VMName']}</td>
                            <td>{row['DBaaSType']}</td>
                            <td class="{status_class}">{row['DBaaSStatus']}</td>
                            <td class="ip-cell" onclick="navigator.clipboard.writeText('{ip}'); alert('已複製 IP: {ip}')">{ip}</td>
                        </tr>
'''

    accordion_html += '''
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</div>
'''

# ==================== 完整 HTML 頁面模板 ====================
full_html = f'''<!DOCTYPE html>
<html lang="zh-HK">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>DBaaS VM 清單 | DBA書簽</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="icon" type="image/x-icon" href="IMG_3755.jpg">
    
    <style>
        .btn-outline-custom {{
            --bs-btn-color: #343a40;
            --bs-btn-bg: #DCBFFF;
            --bs-btn-border-color: #DCBFFF;
            --bs-btn-hover-color: #343a40;
            --bs-btn-hover-bg: #E5D4FF;
            --bs-btn-hover-border-color: #E5D4FF;
        }}
        .table-sm td, .table-sm th {{ padding: 0.4rem 0.6rem; font-size: 0.9rem; }}
        .ip-cell {{ font-family: 'Courier New', monospace; cursor: pointer; }}
        .ip-cell:hover {{ background-color: #f0e8ff; }}
        .status-active {{ color: #198754; font-weight: bold; }}
        .text-danger {{ color: #dc3545 !important; }}
    </style>
</head>
<body>

<header class="p-3 mb-3 border-bottom" style="background-color: #D0A2F7">
    <div class="container">
        <div class="d-flex flex-wrap align-items-center justify-content-between">
            <div>
                <h4 class="mb-0">DBaaS VM 清單</h4>
                <small>每月更新 | GCISID / Site / VM / IP 一覽</small>
            </div>
            <a href="index.html" class="btn btn-outline-custom">← 回主頁</a>
        </div>
    </div>
</header>

<div class="container mb-5">

    <!-- 搜尋框 -->
    <div class="row mb-4">
        <div class="col-md-6">
            <input type="text" id="searchInput" class="form-control" placeholder="輸入 GCISID / VM 名稱 / IP 搜尋...">
        </div>
    </div>

    <div class="accordion" id="dbaar-accordion">
        {accordion_html}
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
<script>
// 簡單搜尋功能
document.getElementById('searchInput').addEventListener('input', function(e) {{
    const term = e.target.value.toLowerCase();
    document.querySelectorAll('.accordion-item').forEach(item => {{
        const text = item.textContent.toLowerCase();
        item.style.display = text.includes(term) ? '' : 'none';
    }});
}});
</script>

</body>
</html>'''

# 寫入完整 HTML 檔案
with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
    f.write(full_html)

print(f"完成！已產生完整頁面：{OUTPUT_HTML}")
print("你可以直接用瀏覽器開啟這個檔案使用。")
print("每月只要更新 Excel 檔名一致 → 重新執行腳本 → 完成！")