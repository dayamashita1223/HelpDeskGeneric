"""新環境にサンプルデータを投入（担当者 Lookup の NavProp を確認して使用）"""
import sys, json
from datetime import datetime, timedelta, timezone
sys.path.insert(0, "scripts")
from auth_helper import get_session, DATAVERSE_URL

s = get_session()
BASE = f"{DATAVERSE_URL}/api/data/v9.2"

def api_post(entity_set, body):
    r = s.post(f"{BASE}/{entity_set}", json=body)
    if r.status_code in (200, 201, 204):
        loc = r.headers.get("OData-EntityId", "")
        if "(" in loc:
            return loc.split("(")[-1].rstrip(")")
        return None
    else:
        print(f"  [ERROR] {r.status_code}: {r.text[:300]}")
        return None

# NavProp 確認
print("=== 担当者 Lookup NavProp 確認 ===")
r = s.get(f"{BASE}/EntityDefinitions(LogicalName='new_ticket')/ManyToOneRelationships?$filter=IsCustomRelationship eq true&$select=SchemaName,ReferencingAttribute,ReferencedEntity,ReferencingEntityNavigationPropertyName")
r.raise_for_status()
for rel in r.json()["value"]:
    print(f"  {rel['ReferencingAttribute']} -> {rel['ReferencedEntity']} (NavProp: {rel['ReferencingEntityNavigationPropertyName']})")
personincharge_nav = None
for rel in r.json()["value"]:
    if "personincharge" in rel["ReferencingAttribute"].lower():
        personincharge_nav = rel["ReferencingEntityNavigationPropertyName"]
        break
print(f"\n  担当者 NavProp: {personincharge_nav}")

# ユーザー取得
print("\n=== 現在のユーザー ===")
me = s.get(f"{BASE}/WhoAmI").json()
user_id = me["UserId"]
user_info = s.get(f"{BASE}/systemusers({user_id})?$select=fullname").json()
print(f"  {user_info['fullname']} ({user_id})")

now = datetime.now(timezone.utc)

# カテゴリ
print("\n=== カテゴリ作成 ===")
categories = [
    {"new_name": "IT インフラ", "new_description": "ネットワーク、サーバー、セキュリティ等", "new_sortorder": 1},
    {"new_name": "アプリケーション", "new_description": "業務アプリ、SaaS、ライセンス等", "new_sortorder": 2},
    {"new_name": "アカウント管理", "new_description": "ユーザーアカウント、権限、パスワード等", "new_sortorder": 3},
    {"new_name": "ハードウェア", "new_description": "PC、周辺機器、モバイルデバイス等", "new_sortorder": 4},
    {"new_name": "その他", "new_description": "上記以外の問い合わせ", "new_sortorder": 5},
]
cat_ids = {}
for c in categories:
    rid = api_post("new_categories", c)
    if rid:
        cat_ids[c["new_name"]] = rid
        print(f"  ✅ {c['new_name']}")

sub_categories = [
    {"new_name": "ネットワーク接続", "new_sortorder": 1, "parent": "IT インフラ"},
    {"new_name": "セキュリティ", "new_sortorder": 2, "parent": "IT インフラ"},
    {"new_name": "Microsoft 365", "new_sortorder": 1, "parent": "アプリケーション"},
    {"new_name": "業務システム", "new_sortorder": 2, "parent": "アプリケーション"},
    {"new_name": "パスワードリセット", "new_sortorder": 1, "parent": "アカウント管理"},
    {"new_name": "PC 故障・交換", "new_sortorder": 1, "parent": "ハードウェア"},
]
for sc in sub_categories:
    parent = sc.pop("parent")
    if parent in cat_ids:
        sc["new_parentcategoryid@odata.bind"] = f"/new_categories({cat_ids[parent]})"
    rid = api_post("new_categories", sc)
    if rid:
        cat_ids[sc["new_name"]] = rid
        print(f"  ✅ {sc['new_name']} (子)")

# ナレッジ記事
print("\n=== ナレッジ記事作成 ===")
articles = [
    {"new_title": "VPN に接続できない場合の対処法", "new_articlenumber": "KB-001",
     "new_question": "在宅勤務時に VPN に接続できません。", "new_answer": "1. VPN クライアントを再起動\n2. インターネット接続を確認\n3. 最新版にアップデート",
     "new_articlestatus": 4, "new_visibility": 1, "category": "ネットワーク接続"},
    {"new_title": "パスワードリセットの手順", "new_articlenumber": "KB-002",
     "new_question": "AD のパスワードを忘れました。", "new_answer": "1. SSPR ポータルにアクセス\n2. 本人確認\n3. 新パスワード設定",
     "new_articlestatus": 4, "new_visibility": 1, "category": "パスワードリセット"},
    {"new_title": "Teams 会議の録画が再生できない", "new_articlenumber": "KB-003",
     "new_question": "録画した会議が再生できません。", "new_answer": "1. 数分待つ\n2. アクセス権確認\n3. ブラウザキャッシュクリア",
     "new_articlestatus": 4, "new_visibility": 1, "category": "Microsoft 365"},
    {"new_title": "PC ブルースクリーン発生時の初期対応", "new_articlenumber": "KB-004",
     "new_question": "BSOD が発生しました。", "new_answer": "1. エラーコードを記録\n2. PC再起動\n3. 頻発ならチケット起票",
     "new_articlestatus": 3, "new_visibility": 2, "category": "PC 故障・交換"},
    {"new_title": "経費精算システムのログインエラー対処法", "new_articlenumber": "KB-005",
     "new_question": "経費精算にログインできません。", "new_answer": "1. Cookie クリア\n2. シークレットウィンドウ\n3. SSO で先にログイン",
     "new_articlestatus": 2, "new_visibility": 1, "category": "業務システム"},
]
article_ids = {}
for a in articles:
    cat = a.pop("category", None)
    body = dict(a)
    if cat and cat in cat_ids:
        body["new_categoryid@odata.bind"] = f"/new_categories({cat_ids[cat]})"
    body["new_reviewerid@odata.bind"] = f"/systemusers({user_id})"
    rid = api_post("new_knowledgearticles", body)
    if rid:
        article_ids[a["new_articlenumber"]] = rid
        print(f"  ✅ {a['new_title']}")

# チケット
print("\n=== チケット作成 ===")
tickets = [
    {"new_title": "VPN 接続が頻繁に切断される", "new_ticketnumber": "TK-001", "new_priority": 1, "new_severity": 2,
     "new_ticketstatus": 2, "new_requestedon": (now - timedelta(days=3)).isoformat(),
     "new_duedate": (now + timedelta(days=1)).isoformat(), "new_questioncontent": "在宅勤務中、VPN が10分おきに切断されます。",
     "category": "ネットワーク接続", "new_aisuggestion": "VPN クライアントを最新版にアップデートしてください。"},
    {"new_title": "Teams でファイル共有ができない", "new_ticketnumber": "TK-002", "new_priority": 2, "new_severity": 3,
     "new_ticketstatus": 1, "new_requestedon": (now - timedelta(days=1)).isoformat(),
     "new_duedate": (now + timedelta(days=3)).isoformat(), "new_questioncontent": "Teams にファイルをアップロードできません。",
     "category": "Microsoft 365"},
    {"new_title": "経費精算システムにログインできない", "new_ticketnumber": "TK-003", "new_priority": 2, "new_severity": 3,
     "new_ticketstatus": 5, "new_requestedon": (now - timedelta(days=5)).isoformat(),
     "new_resolvedon": (now - timedelta(days=4)).isoformat(),
     "new_resolution": "SSO セッション期限切れ。再ログインで解決。", "category": "業務システム", "related_kb": "KB-005"},
    {"new_title": "ノートPC のバッテリーが膨張している", "new_ticketnumber": "TK-004", "new_priority": 1, "new_severity": 1,
     "new_ticketstatus": 4, "new_requestedon": (now - timedelta(days=2)).isoformat(),
     "new_duedate": now.isoformat(), "new_questioncontent": "ThinkPad のバッテリーが膨張しています。",
     "category": "PC 故障・交換", "new_aisuggestion": "直ちに使用中止。代替PC手配を進めてください。"},
    {"new_title": "新入社員の AD アカウント作成依頼", "new_ticketnumber": "TK-005", "new_priority": 2, "new_severity": 3,
     "new_ticketstatus": 6, "new_requestedon": (now - timedelta(days=10)).isoformat(),
     "new_resolvedon": (now - timedelta(days=7)).isoformat(), "new_closedon": (now - timedelta(days=6)).isoformat(),
     "new_resolution": "5名分の AD アカウントを作成しました。", "category": "パスワードリセット"},
    {"new_title": "社内 Wi-Fi に接続できない（来客用）", "new_ticketnumber": "TK-006", "new_priority": 1, "new_severity": 2,
     "new_ticketstatus": 3, "new_requestedon": (now - timedelta(hours=6)).isoformat(),
     "new_duedate": (now + timedelta(hours=2)).isoformat(), "new_questioncontent": "来客用 Wi-Fi に接続できません。",
     "category": "ネットワーク接続"},
    {"new_title": "不審なメールを受信した", "new_ticketnumber": "TK-007", "new_priority": 1, "new_severity": 2,
     "new_ticketstatus": 2, "new_requestedon": (now - timedelta(hours=3)).isoformat(),
     "new_questioncontent": "フィッシングメールの疑い。ドメインが m1crosoft です。",
     "category": "セキュリティ", "new_aisuggestion": "フィッシングの可能性高。リンクをクリックしないでください。"},
    {"new_title": "Outlook の予定表が同期されない", "new_ticketnumber": "TK-008", "new_priority": 3, "new_severity": 4,
     "new_ticketstatus": 1, "new_requestedon": now.isoformat(),
     "new_duedate": (now + timedelta(days=7)).isoformat(), "new_questioncontent": "デスクトップとモバイルの予定表が同期されません。",
     "category": "Microsoft 365"},
]

for t in tickets:
    cat = t.pop("category", None)
    kb = t.pop("related_kb", None)
    body = dict(t)
    body["new_requestorid@odata.bind"] = f"/systemusers({user_id})"
    if personincharge_nav:
        body[f"{personincharge_nav}@odata.bind"] = f"/systemusers({user_id})"
    if cat and cat in cat_ids:
        body["new_categoryid@odata.bind"] = f"/new_categories({cat_ids[cat]})"
    if kb and kb in article_ids:
        body["new_relatedknowledgearticleid@odata.bind"] = f"/new_knowledgearticles({article_ids[kb]})"
    rid = api_post("new_tickets", body)
    if rid:
        print(f"  ✅ {t['new_title']}")
    else:
        print(f"  ❌ {t['new_title']}")

# 件数確認
print("\n=== データ件数 ===")
for eset in ["new_categories", "new_tickets", "new_knowledgearticles"]:
    r = s.get(f"{BASE}/{eset}?$select=createdon")
    print(f"  {eset}: {len(r.json().get('value', []))}")
print("\n✅ サンプルデータ投入完了")
