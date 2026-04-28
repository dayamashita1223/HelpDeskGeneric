"""
汎用ヘルプデスク チャート一括作成スクリプト

ソリューションインポート後に実行して、チケットテーブルにチャートを追加する。
.env の APP_MODULE_ID にアプリ ID を設定してから実行。

使い方:
  python scripts/create_charts.py

作成するチャート (9種):
  1. ステータス別チケット件数（横棒）
  2. 優先度別チケット件数（縦棒）
  3. カテゴリ別チケット件数（横棒）
  4. チケット起票推移（月別・折れ線）
  5. チケット起票推移（日別・折れ線）
  6. 担当者別チケット件数（横棒）
  7. 担当者×ステータス別（積み上げ横棒）
  8. 未対応チケット 担当者別（横棒）
  9. 未対応チケット ステータス×担当者（積み上げ横棒）
"""
import os, sys, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import requests
from dotenv import load_dotenv
from auth_helper import get_token as _get_token

load_dotenv()

DATAVERSE_URL = os.environ["DATAVERSE_URL"].rstrip("/")
SOLUTION_NAME = os.environ.get("SOLUTION_NAME", "HelpDeskGeneric")
APP_MODULE_ID = os.environ.get("APP_MODULE_ID", "")
API = f"{DATAVERSE_URL}/api/data/v9.2"

def hdr(sol=True, merge=False):
    token = _get_token()
    h = {
        "Authorization": f"Bearer {token}",
        "OData-MaxVersion": "4.0", "OData-Version": "4.0",
        "Accept": "application/json",
        "Content-Type": "application/json; charset=utf-8",
    }
    if merge:
        h["MSCRM.MergeLabels"] = "true"
    if sol and SOLUTION_NAME:
        h["MSCRM.SolutionName"] = SOLUTION_NAME
    return h

def extract_id(resp):
    m = re.search(r"\(([0-9a-fA-F\-]+)\)", resp.headers.get("OData-EntityId", ""))
    return m.group(1) if m else None

def upsert_chart(name, data_desc, pres_desc):
    """チャートをべき等に作成/更新"""
    r = requests.get(
        f"{API}/savedqueryvisualizations?$filter=primaryentitytypecode eq 'new_ticket' and name eq '{name}'&$select=savedqueryvisualizationid",
        headers=hdr(False)
    )
    r.raise_for_status()
    existing = r.json().get("value", [])

    if existing:
        cid = existing[0]["savedqueryvisualizationid"]
        r = requests.patch(f"{API}/savedqueryvisualizations({cid})", headers=hdr(merge=True),
            json={"datadescription": data_desc, "presentationdescription": pres_desc})
        status = "更新" if r.ok else f"更新失敗({r.status_code})"
        print(f"  {status}: {name} ({cid})")
        return cid
    else:
        r = requests.post(f"{API}/savedqueryvisualizations", headers=hdr(merge=False),
            json={"name": name, "primaryentitytypecode": "new_ticket",
                  "datadescription": data_desc, "presentationdescription": pres_desc})
        if r.ok:
            cid = extract_id(r)
            print(f"  作成: {name} ({cid})")
            return cid
        else:
            print(f"  失敗: {name} ({r.status_code}): {r.text[:200]}")
            return None

# ── チャート定義 ──

BAR_PRES = lambda colors: (
    f'<Chart Palette="None" PaletteCustomColors="{colors}">'
    '<Series><Series Name="count" ChartType="Bar" IsValueShownAsLabel="True" Font="{0}, 9.5px" LabelForeColor="59, 59, 59" /></Series>'
    '<ChartAreas><ChartArea BorderColor="White" BorderDashStyle="Solid">'
    '<AxisY LabelAutoFitStyle="DecreaseFont" TitleForeColor="59, 59, 59" TitleFont="{0}, 10.5px" LineColor="165, 172, 181"><MajorGrid LineColor="239, 242, 246" /><MajorTickMark LineColor="165, 172, 181" /><LabelStyle Font="{0}, 10.5px" ForeColor="59, 59, 59" /></AxisY>'
    '<AxisX LabelAutoFitStyle="DecreaseFont" TitleForeColor="59, 59, 59" TitleFont="{0}, 10.5px" LineColor="165, 172, 181"><MajorTickMark LineColor="165, 172, 181" /><MajorGrid LineColor="Transparent" /><LabelStyle Font="{0}, 10.5px" ForeColor="59, 59, 59" /></AxisX>'
    '<Area3DStyle Enable3D="false" /></ChartArea></ChartAreas></Chart>'
)

COL_PRES = lambda colors: (
    f'<Chart Palette="None" PaletteCustomColors="{colors}">'
    '<Series><Series Name="count" ChartType="Column" IsValueShownAsLabel="True" Font="{0}, 9.5px" LabelForeColor="59, 59, 59" /></Series>'
    '<ChartAreas><ChartArea BorderColor="White" BorderDashStyle="Solid">'
    '<AxisY LabelAutoFitStyle="DecreaseFont" TitleForeColor="59, 59, 59" TitleFont="{0}, 10.5px" LineColor="165, 172, 181"><MajorGrid LineColor="239, 242, 246" /><MajorTickMark LineColor="165, 172, 181" /><LabelStyle Font="{0}, 10.5px" ForeColor="59, 59, 59" /></AxisY>'
    '<AxisX LabelAutoFitStyle="DecreaseFont" TitleForeColor="59, 59, 59" TitleFont="{0}, 10.5px" LineColor="165, 172, 181"><MajorTickMark LineColor="165, 172, 181" /><MajorGrid LineColor="Transparent" /><LabelStyle Font="{0}, 10.5px" ForeColor="59, 59, 59" /></AxisX>'
    '<Area3DStyle Enable3D="false" /></ChartArea></ChartAreas></Chart>'
)

STACKED_BAR_PRES = lambda colors: (
    f'<Chart Palette="None" PaletteCustomColors="{colors}">'
    '<Series><Series Name="count" ChartType="StackedBar" IsValueShownAsLabel="True" Font="{0}, 9.5px" LabelForeColor="59, 59, 59" /></Series>'
    '<ChartAreas><ChartArea BorderColor="White" BorderDashStyle="Solid">'
    '<AxisY LabelAutoFitStyle="DecreaseFont" TitleForeColor="59, 59, 59" TitleFont="{0}, 10.5px" LineColor="165, 172, 181"><MajorGrid LineColor="239, 242, 246" /><MajorTickMark LineColor="165, 172, 181" /><LabelStyle Font="{0}, 10.5px" ForeColor="59, 59, 59" /></AxisY>'
    '<AxisX LabelAutoFitStyle="DecreaseFont" TitleForeColor="59, 59, 59" TitleFont="{0}, 10.5px" LineColor="165, 172, 181"><MajorTickMark LineColor="165, 172, 181" /><MajorGrid LineColor="Transparent" /><LabelStyle Font="{0}, 10.5px" ForeColor="59, 59, 59" /></AxisX>'
    '<Area3DStyle Enable3D="false" /></ChartArea></ChartAreas></Chart>'
)

LINE_PRES = lambda title_y="件数", title_x="": (
    '<Chart Palette="None" PaletteCustomColors="55,118,193">'
    '<Series><Series Name="count" ChartType="Line" BorderWidth="3" IsValueShownAsLabel="True" Font="{0}, 9.5px" LabelForeColor="59, 59, 59" MarkerStyle="Circle" MarkerSize="8" MarkerColor="55, 118, 193" /></Series>'
    '<ChartAreas><ChartArea BorderColor="White" BorderDashStyle="Solid">'
    f'<AxisY LabelAutoFitStyle="DecreaseFont" TitleForeColor="59, 59, 59" TitleFont="{{0}}, 10.5px" LineColor="165, 172, 181" Title="{title_y}"><MajorGrid LineColor="239, 242, 246" /><MajorTickMark LineColor="165, 172, 181" /><LabelStyle Font="{{0}}, 10.5px" ForeColor="59, 59, 59" /></AxisY>'
    f'<AxisX LabelAutoFitStyle="DecreaseFont" TitleForeColor="59, 59, 59" TitleFont="{{0}}, 10.5px" LineColor="165, 172, 181" Title="{title_x}"><MajorTickMark LineColor="165, 172, 181" /><MajorGrid LineColor="239, 242, 246" /><LabelStyle Font="{{0}}, 10.5px" ForeColor="59, 59, 59" /></AxisX>'
    '<Area3DStyle Enable3D="false" /></ChartArea></ChartAreas></Chart>'
)

def simple_data(group_attr, alias):
    return (
        '<datadefinition><fetchcollection>'
        '<fetch mapping="logical" aggregate="true"><entity name="new_ticket">'
        f'<attribute name="{group_attr}" groupby="true" alias="{alias}" />'
        '<attribute name="new_ticketid" aggregate="count" alias="count" />'
        '</entity></fetch></fetchcollection>'
        '<categorycollection><category><measurecollection><measure alias="count" /></measurecollection></category></categorycollection>'
        '</datadefinition>'
    )

def dual_group_data(group1, alias1, group2, alias2, filter_xml=""):
    return (
        '<datadefinition><fetchcollection>'
        '<fetch mapping="logical" aggregate="true"><entity name="new_ticket">'
        f'<attribute name="{group1}" groupby="true" alias="{alias1}" />'
        f'<attribute name="{group2}" groupby="true" alias="{alias2}" />'
        '<attribute name="new_ticketid" aggregate="count" alias="count" />'
        f'{filter_xml}'
        '</entity></fetch></fetchcollection>'
        '<categorycollection><category><measurecollection><measure alias="count" /></measurecollection></category></categorycollection>'
        '</datadefinition>'
    )

def date_data(date_attr, grouping, alias):
    return (
        '<datadefinition><fetchcollection>'
        f'<fetch mapping="logical" aggregate="true"><entity name="new_ticket">'
        f'<attribute name="{date_attr}" groupby="true" dategrouping="{grouping}" alias="{alias}" />'
        '<attribute name="new_ticketid" aggregate="count" alias="count" />'
        '</entity></fetch></fetchcollection>'
        '<categorycollection><category><measurecollection><measure alias="count" /></measurecollection></category></categorycollection>'
        '</datadefinition>'
    )

UNRESOLVED_FILTER = '<filter type="and"><condition attribute="new_ticketstatus" operator="in"><value>1</value><value>2</value><value>3</value><value>4</value></condition></filter>'

CHARTS = [
    ("ステータス別チケット件数",
     simple_data("new_ticketstatus", "status"),
     STACKED_BAR_PRES("55,118,193; 237,125,49; 255,192,0; 224,60,49; 112,173,71; 165,165,165")),

    ("優先度別チケット件数",
     simple_data("new_priority", "priority"),
     COL_PRES("55,118,193; 237,125,49; 165,165,165")),

    ("カテゴリ別チケット件数",
     simple_data("new_categoryid", "category"),
     BAR_PRES("55,118,193; 237,125,49; 165,165,165; 255,192,0; 68,114,196")),

    ("チケット起票推移（月別）",
     date_data("new_requestedon", "month", "month"),
     LINE_PRES("件数", "月")),

    ("チケット起票推移（日別）",
     date_data("new_requestedon", "day", "day"),
     LINE_PRES("件数", "日")),

    ("担当者別チケット件数",
     simple_data("crf98_personincharge", "assignee"),
     BAR_PRES("55,118,193; 237,125,49; 112,173,71; 255,192,0; 68,114,196; 165,165,165")),

    ("担当者×ステータス別",
     dual_group_data("crf98_personincharge", "assignee", "new_ticketstatus", "status"),
     STACKED_BAR_PRES("55,118,193; 237,125,49; 255,192,0; 224,60,49; 112,173,71; 165,165,165")),

    ("未対応チケット（担当者別）",
     '<datadefinition><fetchcollection>'
     '<fetch mapping="logical" aggregate="true"><entity name="new_ticket">'
     '<attribute name="crf98_personincharge" groupby="true" alias="assignee" />'
     '<attribute name="new_ticketid" aggregate="count" alias="count" />'
     + UNRESOLVED_FILTER +
     '</entity></fetch></fetchcollection>'
     '<categorycollection><category><measurecollection><measure alias="count" /></measurecollection></category></categorycollection>'
     '</datadefinition>',
     BAR_PRES("224,60,49; 237,125,49; 55,118,193; 112,173,71; 255,192,0; 165,165,165")),

    ("未対応チケット（ステータス×担当者）",
     '<datadefinition><fetchcollection>'
     '<fetch mapping="logical" aggregate="true"><entity name="new_ticket">'
     '<attribute name="crf98_personincharge" groupby="true" alias="assignee" />'
     '<attribute name="new_ticketstatus" groupby="true" alias="status" />'
     '<attribute name="new_ticketid" aggregate="count" alias="count" />'
     + UNRESOLVED_FILTER +
     '</entity></fetch></fetchcollection>'
     '<categorycollection><category><measurecollection><measure alias="count" /></measurecollection></category></categorycollection>'
     '</datadefinition>',
     STACKED_BAR_PRES("55,118,193; 237,125,49; 255,192,0; 224,60,49; 112,173,71; 165,165,165")),
]


def main():
    print("=" * 50)
    print("  チャート一括作成（9種）")
    print(f"  ソリューション: {SOLUTION_NAME}")
    print(f"  アプリ ID: {APP_MODULE_ID or '(未設定)'}")
    print("=" * 50)

    chart_ids = []
    for name, data_desc, pres_desc in CHARTS:
        cid = upsert_chart(name, data_desc, pres_desc)
        if cid:
            chart_ids.append(cid)

    # アプリに追加
    if APP_MODULE_ID and chart_ids:
        print(f"\n=== アプリに追加 ({len(chart_ids)} 件) ===")
        for cid in chart_ids:
            try:
                r = requests.post(f"{API}/AddAppComponents", headers=hdr(False),
                    json={"AppId": APP_MODULE_ID, "Components": [
                        {"savedqueryvisualizationid": cid, "@odata.type": "Microsoft.Dynamics.CRM.savedqueryvisualization"}
                    ]})
                if r.ok:
                    print(f"  ✅ {cid}")
                elif "already" in r.text.lower():
                    print(f"  ⏭️ 既に追加済み: {cid}")
                else:
                    print(f"  ⚠️ {r.status_code}")
            except:
                pass
    elif not APP_MODULE_ID:
        print("\n⚠️ .env に APP_MODULE_ID が未設定のためアプリへの追加をスキップ")
        print("  アプリ作成後に APP_MODULE_ID を設定して再実行してください")

    # 公開
    print(f"\n=== 公開 ===")
    requests.post(f"{API}/PublishXml", headers=hdr(False),
        json={"ParameterXml": '<importexportxml><entities><entity>new_ticket</entity></entities></importexportxml>'})
    if APP_MODULE_ID:
        requests.post(f"{API}/PublishXml", headers=hdr(False),
            json={"ParameterXml": f"<importexportxml><appmodules><appmodule>{APP_MODULE_ID}</appmodule></appmodules></importexportxml>"})
    print("  ✅ 公開完了")

    print(f"\n=== 作成完了: {len(chart_ids)} / {len(CHARTS)} チャート ===")


if __name__ == "__main__":
    main()
