"""フォームの現在のタブ構造を確認"""
import sys, os, re
sys.path.insert(0, "scripts")
from auth_helper import get_session, DATAVERSE_URL

s = get_session()
API = f"{DATAVERSE_URL}/api/data/v9.2"

for entity in ["new_ticket", "new_knowledgearticle", "new_category"]:
    print(f"\n{'='*50}")
    print(f"=== {entity} ===")
    r = s.get(f"{API}/systemforms?$filter=objecttypecode eq '{entity}' and type eq 2&$select=formid,name,formxml&$top=1")
    r.raise_for_status()
    forms = r.json()["value"]
    if not forms:
        print("  フォームなし")
        continue
    form = forms[0]
    xml = form["formxml"]
    
    # タブ部分を抽出
    tabs = re.findall(r'<tab\s+name="([^"]*)"[^>]*>.*?<labels>(.*?)</labels>', xml, re.DOTALL)
    print(f"  フォーム: {form['name']} ({form['formid']})")
    print(f"  タブ数: {len(tabs)}")
    for name, labels_xml in tabs:
        # label の description を取得
        descs = re.findall(r'<label\s+description="([^"]*)"', labels_xml)
        print(f"    name=\"{name}\" → labels: {descs}")
    
    # フォーム全文（最初の2000文字）
    print(f"\n  XML ({len(xml)} chars):")
    print(f"  {xml[:2000]}")
