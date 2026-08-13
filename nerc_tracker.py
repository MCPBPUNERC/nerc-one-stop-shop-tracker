import hashlib
import json
import os
import re
import smtplib
from collections import defaultdict
from datetime import datetime, timezone
from difflib import SequenceMatcher
from email.message import EmailMessage
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from openpyxl import load_workbook

EXCEL_URL = "https://www.nerc.com/globalassets/align-reports/one-stop-shop.xlsx"
PROFILE_PATH = Path("docs/data/profile.json")
STATE_PATH = Path("state/nerc-records.json")
DATA_PATH = Path("docs/data/nerc.json")
HISTORY_PATH = Path("docs/data/nerc-history.json")
MAX_WEB_CHANGES = 500
MAX_EMAIL_CHANGES = 25
GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_PASS = os.environ.get("GMAIL_PASS")
RECIPIENTS = [x.strip() for x in os.environ.get("RECIPIENTS", "mikep@mcphersonpower.com,tommys@mcphersonpower.com").split(",") if x.strip()]
STANDARD_RE = re.compile(r"\b([A-Z]{3})-(\d{3})(?:-[A-Za-z0-9.]+)?\b")
HEADER_WORDS = ("standard","requirement","version","status","enforcement","effective","implementation","subject","title","project","date","retirement")
IDENTITY_WORDS = ("standard","requirement","requirement number","section","subject","title","project","version","criterion","reference")
VOLATILE_WORDS = ("status","enforcement","effective","implementation","date","retirement","inactive","mandatory","future enforcement","notes","comment")

def load_json(path, default):
    try: return json.loads(path.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError): return default

def value_text(value):
    if value is None: return ""
    if isinstance(value, datetime): return value.isoformat()
    return re.sub(r"\s+", " ", str(value)).strip()

def normalize(text): return re.sub(r"[^a-z0-9]+", " ", value_text(text).lower()).strip()

def extract_standards(values):
    found=set()
    for value in values:
        for match in STANDARD_RE.finditer(value_text(value).upper()): found.add(f"{match.group(1)}-{match.group(2)}")
    return sorted(found)

def unique_headers(values):
    result=[]; seen=defaultdict(int)
    for idx,value in enumerate(values,start=1):
        name=value_text(value) or f"Column {idx}"; seen[name]+=1
        if seen[name]>1: name=f"{name} ({seen[name]})"
        result.append(name)
    return result

def find_header_row(rows):
    best_index,best_score=0,-1
    for idx,row in enumerate(rows[:30]):
        texts=[normalize(v) for v in row if value_text(v)]
        if len(texts)<3: continue
        score=sum(1 for text in texts for word in HEADER_WORDS if word in text)*5+min(len(texts),20)
        if score>best_score: best_score,best_index=score,idx
    return best_index

def stable_identity(fields, standards):
    selected=[]
    for header,value in fields.items():
        if not value: continue
        h=normalize(header)
        if any(word in h for word in VOLATILE_WORDS): continue
        if any(word in h for word in IDENTITY_WORDS): selected.append(f"{h}={normalize(value)}")
    if not selected:
        for header,value in list(fields.items())[:6]:
            h=normalize(header)
            if value and not any(word in h for word in VOLATILE_WORDS): selected.append(f"{h}={normalize(value)}")
    identity=(",".join(standards) if standards else "unclassified")+"|"+"|".join(selected[:8])
    return hashlib.sha1(identity.encode()).hexdigest()[:20], identity

def parse_workbook(path, tracked):
    wb=load_workbook(path,read_only=True,data_only=False); records=[]; sheet_meta=[]; total=0
    for sheet in wb.worksheets:
        rows=[tuple(r) for r in sheet.iter_rows(values_only=True)]
        if not rows: continue
        hi=find_header_row(rows); headers=unique_headers(rows[hi]); previous=[]; start=len(records)
        for excel_row,values in enumerate(rows[hi+1:],start=hi+2):
            texts=[value_text(v) for v in values]; populated=sum(bool(v) for v in texts); total+=populated
            if not populated: previous=[]; continue
            fields={headers[i] if i<len(headers) else f"Column {i+1}":v for i,v in enumerate(texts) if v}
            standards=extract_standards(texts)
            if standards: previous=standards
            elif previous: standards=previous
            key,identity=stable_identity(fields,standards); primary=standards[0] if standards else ""
            records.append({"key":key,"identity_text":identity,"sheet":sheet.title,"source_row":excel_row,"standards":standards,"standard":primary,"family":primary.split("-")[0] if primary else "OTHER","tracked":any(s in tracked for s in standards),"fields":fields})
        sheet_meta.append({"name":sheet.title,"rows":sheet.max_row,"columns":sheet.max_column,"records":len(records)-start})
    wb.close(); return records,sheet_meta,total

def similarity(old,new):
    if old.get("standard") and new.get("standard") and old["standard"]!=new["standard"]: return 0
    return SequenceMatcher(None,old.get("identity_text",""),new.get("identity_text","")).ratio()

def severity(change):
    if not change.get("tracked"): return "info"
    text=" ".join([change.get("field",""),change.get("old",""),change.get("new","")]).lower()
    if any(x in text for x in ("enforcement","mandatory","inactive","effective date","implementation","retirement","subject to future enforcement")): return "high"
    if any(x in text for x in ("requirement","applicability","version","status","violation","approval","ballot")): return "moderate"
    return "informational"

def field_changes(old,new):
    out=[]
    for header in sorted(set(old.get("fields",{}))|set(new.get("fields",{}))):
        a=value_text(old.get("fields",{}).get(header,"")); b=value_text(new.get("fields",{}).get(header,""))
        if a==b: continue
        typ="added" if not a else "removed" if not b else "modified"; standard=new.get("standard") or old.get("standard") or "Unclassified"
        item={"standard":standard,"family":standard.split("-")[0] if "-" in standard else "OTHER","sheet":new.get("sheet") or old.get("sheet",""),"field":header,"old":a,"new":b,"type":typ,"tracked":bool(new.get("tracked") or old.get("tracked")),"source_row":new.get("source_row") or old.get("source_row")}; item["severity"]=severity(item); out.append(item)
    return out

def compare_records(previous,current):
    old_by=defaultdict(list); new_by=defaultdict(list)
    for r in previous: old_by[r["key"]].append(r)
    for r in current: new_by[r["key"]].append(r)
    matched=[]; ou=[]; nu=[]
    for key in set(old_by)|set(new_by):
        a,b=old_by.get(key,[]),new_by.get(key,[]); n=min(len(a),len(b)); matched.extend(zip(a[:n],b[:n])); ou.extend(a[n:]); nu.extend(b[n:])
    candidates=[]
    for oi,o in enumerate(ou):
        for ni,n in enumerate(nu):
            if o.get("sheet")==n.get("sheet"):
                score=similarity(o,n)
                if score>=.72: candidates.append((score,oi,ni))
    candidates.sort(reverse=True); uo=set(); un=set()
    for _,oi,ni in candidates:
        if oi in uo or ni in un: continue
        uo.add(oi); un.add(ni); matched.append((ou[oi],nu[ni]))
    changes=[]
    for a,b in matched: changes.extend(field_changes(a,b))
    for i,r in enumerate(ou):
        if i not in uo:
            item={"standard":r.get("standard") or "Unclassified","family":r.get("family","OTHER"),"sheet":r.get("sheet",""),"field":"Record","old":"Record present","new":"Record removed","type":"record_removed","tracked":r.get("tracked",False),"source_row":r.get("source_row")}; item["severity"]=severity(item); changes.append(item)
    for i,r in enumerate(nu):
        if i not in un:
            item={"standard":r.get("standard") or "Unclassified","family":r.get("family","OTHER"),"sheet":r.get("sheet",""),"field":"Record","old":"Record absent","new":"Record added","type":"record_added","tracked":r.get("tracked",False),"source_row":r.get("source_row")}; item["severity"]=severity(item); changes.append(item)
    return changes

def send_email(subject,body):
    if not GMAIL_USER or not GMAIL_PASS: raise RuntimeError("Gmail secrets are not set")
    msg=EmailMessage(); msg["Subject"]=subject; msg["From"]=GMAIL_USER; msg["To"] = ", ".join(RECIPIENTS); msg.set_content(body)
    with smtplib.SMTP("smtp.gmail.com",587,timeout=60) as s: s.starttls(); s.login(GMAIL_USER,GMAIL_PASS); s.send_message(msg)

def main():
    profile=load_json(PROFILE_PATH,{}); tracked=set(profile.get("tracked_standards",[]))
    if not tracked: raise RuntimeError("No tracked standards configured")
    now=datetime.now(timezone.utc); ct=now.astimezone(ZoneInfo("America/Chicago")); temp=Path("current_snapshot.xlsx")
    response=requests.get(EXCEL_URL,timeout=90,headers={"User-Agent":"BPU-NERC-Regulatory-Control-Room/2.0"}); response.raise_for_status(); temp.write_bytes(response.content)
    records,sheets,cells=parse_workbook(temp,tracked); temp.unlink(missing_ok=True)
    previous=load_json(STATE_PATH,{}).get("records",[]); baseline=not bool(previous); changes=[] if baseline else compare_records(previous,records)
    tracked_changes=[c for c in changes if c.get("tracked")]; other=[c for c in changes if not c.get("tracked")]; high=sum(c["severity"]=="high" for c in tracked_changes); moderate=sum(c["severity"]=="moderate" for c in tracked_changes)
    message="Semantic baseline established" if baseline else "Tracked changes detected" if tracked_changes else "No tracked changes"
    brief=f"Version 2 established a record-aware NERC baseline across {len(records):,} records." if baseline else f"{len(tracked_changes):,} changes affect BPU-tracked standards ({high} high priority, {moderate} moderate). {len(other):,} additional changes were outside the tracked-standard list." if tracked_changes else (f"No changes were detected in BPU's tracked standards. {len(other):,} changes were detected elsewhere in the NERC workbook." if other else "No semantic record changes were detected in the NERC workbook.")
    history=load_json(HISTORY_PATH,[]); history=[{"date":ct.strftime("%b %d, %Y %I:%M %p CT"),"tracked_changes":len(tracked_changes),"other_changes":len(other),"high":high,"status_text":message}]+history
    payload={"schema_version":2,"status":"ok","message":message,"checked_at":now.isoformat(),"checked_at_display":ct.strftime("%b %d, %Y %I:%M %p"),"source_url":EXCEL_URL,"baseline":baseline,"summary":{"tracked_standards":len(tracked),"registered_functions":profile.get("registered_functions",[]),"sheets_checked":len(sheets),"records_reviewed":len(records),"cells_reviewed":cells,"tracked_changes":len(tracked_changes),"other_changes":len(other),"high":high,"moderate":moderate,"informational":max(0,len(tracked_changes)-high-moderate)},"brief":brief,"changes":tracked_changes[:MAX_WEB_CHANGES],"other_changes":other[:100],"changes_truncated":len(tracked_changes)>MAX_WEB_CHANGES,"sheets":sheets,"history":history[:45],"health":{"nerc_feed":"ok","semantic_compare":"ok","email":"pending"}}
    DATA_PATH.parent.mkdir(parents=True,exist_ok=True); STATE_PATH.parent.mkdir(parents=True,exist_ok=True); DATA_PATH.write_text(json.dumps(payload,indent=2),encoding="utf-8"); HISTORY_PATH.write_text(json.dumps(history[:180],indent=2),encoding="utf-8"); STATE_PATH.write_text(json.dumps({"records":records},indent=2),encoding="utf-8")
    details=[f"{c['standard']} | {c['field']} | {c['severity'].upper()}\n  Previous: {c['old'] or '(blank)'}\n  Current: {c['new'] or '(blank)'}" for c in tracked_changes[:MAX_EMAIL_CHANGES]]
    try:
        send_email(f"[BPU NERC Control Room] {message} - {ct:%Y-%m-%d}",brief+("\n\n"+"\n\n".join(details) if details else "")+"\n\nControl Room: https://mcpbpunerc.github.io/nerc-one-stop-shop-tracker/"); payload["health"]["email"]="ok"
    except Exception as exc:
        payload["health"]["email"]="error"; payload["health"]["email_error"]=type(exc).__name__; print(f"WARNING email failed: {exc}")
    DATA_PATH.write_text(json.dumps(payload,indent=2),encoding="utf-8")

if __name__=="__main__": main()
