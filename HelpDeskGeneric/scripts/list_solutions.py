"""ソリューション一覧を取得"""
import sys
sys.path.insert(0, "scripts")
from auth_helper import get_session, DATAVERSE_URL

s = get_session()
r = s.get(f"{DATAVERSE_URL}/api/data/v9.2/solutions?$filter=ismanaged eq false and isvisible eq true&$select=uniquename,friendlyname,version&$orderby=friendlyname")
r.raise_for_status()
for sol in r.json()["value"]:
    print(f"  {sol['friendlyname']} ({sol['uniquename']}) v{sol['version']}")
