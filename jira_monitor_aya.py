"""Aya-inspired Jira HDCR desktop monitor.

Install: py -m pip install -U selenium requests
Run:     py jira_monitor_aya.py

Requires authorized Jira access in an automated Microsoft Edge session.
The Outlook/email matching feature is NOT included in this version.
Already-overdue tickets on the first successful scan appear in the table
without generating a flood of historical popups. New breaches alert normally.
"""
import json
import os
import queue
import re
import shutil
import subprocess
import threading
import time
import traceback
import logging
import webbrowser
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote, urlencode, urlparse
from zipfile import ZipFile

import requests
import tkinter as tk
import customtkinter as ctk
from PIL import Image
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait

try:
    import winsound
except ImportError:
    winsound = None

# ----------------------- SETTINGS ---------------------------
JIRA_URL = "https://ayadev.atlassian.net"  # Base domain, NOT queue URL
QUEUE_URL = JIRA_URL + "/jira/servicedesk/projects/HDCR/queues/custom/1"
PROJECT_KEY = "HDCR"
PENDING_STATUSES = ("Waiting for support",)
ACTIVE_STATUS = "In Progress"
ALERT_MINUTES = 10  # alert if a ticket is pending for more than this many minutes
REFRESH_SECONDS = 30  # consulta y refresh cada 30s tras finalizar el ciclo; reloj cada segundo
MAX_PAGES = 15
PAGE_SIZE = 100
TEAM_SIZE = 14  # Total de integrantes del equipo; para calcular el promedio de tickets por agente

ROOT = Path(__file__).resolve().parent
LOGO_PATH = ROOT / "assets" / "logo_aya.png"
PROFILE_DIR = ROOT / "jira_edge_profile"
DRIVER_DIR = ROOT / "driver"
STATE_FILE = ROOT / "jira_alertas.json"
EVENTS = queue.Queue()
STOP = threading.Event()
LOG_FILE = ROOT / "jira_monitor_error.log"


# Colors inspired by Aya
NAVY = "#0D2627"
TEAL = "#00A6A6"
CYAN = "#62CEE1"
PALE = "#EAF7F9"
WHITE = "#FFFFFF"
BG = "#0D2627"
MUTED = "#627A8A"
GREEN = "#168B62"
RED = "#FF6F65"
ORANGE = "#FF913B"


def parse_date(value):
    return datetime.fromisoformat(value.replace("Z", "+00:00"))


def read_state():
    try:
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def write_state(data):
    temporary = STATE_FILE.with_suffix(".tmp")
    temporary.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    temporary.replace(STATE_FILE)


def get_edge_version():
    for hive in (r"HKLM\SOFTWARE\Microsoft\Edge\BLBeacon", r"HKCU\SOFTWARE\Microsoft\Edge\BLBeacon"):
        try:
            proc = subprocess.run(["reg", "query", hive, "/v", "version"], capture_output=True, text=True, timeout=15)
            match = re.search(r"\b\d+\.\d+\.\d+\.\d+\b", proc.stdout)
            if proc.returncode == 0 and match:
                return match.group()
        except Exception:
            pass
    for path in (shutil.which("msedge.exe"),
                 os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Microsoft", "Edge", "Application", "msedge.exe"),
                 os.path.join(os.environ.get("PROGRAMFILES", ""), "Microsoft", "Edge", "Application", "msedge.exe")):
        if path and os.path.isfile(path):
            try:
                proc = subprocess.run([path, "--version"], capture_output=True, text=True, timeout=15)
                match = re.search(r"\b\d+\.\d+\.\d+\.\d+\b", proc.stdout + proc.stderr)
                if match:
                    return match.group()
            except Exception:
                pass
    raise RuntimeError("Could not detect the installed Microsoft Edge version.")


def get_edge_driver_path():
    version = get_edge_version()
    target = DRIVER_DIR / version / "msedgedriver.exe"
    if target.exists():
        return str(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    url = f"https://msedgedriver.microsoft.com/{version}/edgedriver_win64.zip"
    zip_file = target.parent / "edgedriver.zip"
    print("Downloading Edge WebDriver:", url, flush=True)
    try:
        with requests.get(url, stream=True, timeout=60) as response:
            response.raise_for_status()
            with zip_file.open("wb") as output:
                for chunk in response.iter_content(8192):
                    if chunk:
                        output.write(chunk)
        with ZipFile(zip_file) as archive:
            member = next((name for name in archive.namelist()
                           if name.replace("\\", "/").split("/")[-1].lower() == "msedgedriver.exe"), None)
            if member is None:
                raise RuntimeError("msedgedriver.exe was not found in the official ZIP archive.")
            with archive.open(member) as source, target.open("wb") as output:
                shutil.copyfileobj(source, output)
    finally:
        if zip_file.exists():
            zip_file.unlink()
    return str(target)


def launch_edge():
    opts = Options()
    opts.add_argument(f"--user-data-dir={PROFILE_DIR}")
    opts.add_argument("--start-maximized")
    driver = webdriver.Edge(service=Service(get_edge_driver_path()), options=opts)
    driver.set_script_timeout(40)
    return driver


def jira_get(driver, endpoint):
    # Same-origin fetch in the signed-in Jira tab. No cookie extraction and no password storage.
    js = """
        const endpoint = arguments[0], done = arguments[arguments.length - 1];
        fetch(endpoint, {credentials:'same-origin', headers:{'Accept':'application/json'}})
        .then(async response => {
          const raw = await response.text(); let data = null;
          try { data = JSON.parse(raw); } catch(e) {}
          done({code:response.status, data:data, excerpt:raw.slice(0,320)});
        }).catch(e => done({code:0, data:null, excerpt:String(e)}));
    """
    result = driver.execute_async_script(js, endpoint)
    if result.get("code") != 200 or not isinstance(result.get("data"), dict):
        raise RuntimeError(f"Jira HTTP {result.get('code')}: {result.get('excerpt')}")
    return result["data"]


def fetch_issues(driver):
    # Consultar solo Waiting for support e In Progress; Pending excluido.
    quoted_statuses = ",".join('"' + s + '"' for s in (*PENDING_STATUSES, ACTIVE_STATUS))
    jql = f"project = {PROJECT_KEY} AND status IN ({quoted_statuses}) ORDER BY updated DESC"
    token = None
    result = []
    for _ in range(MAX_PAGES):
        params = {"jql": jql, "maxResults": PAGE_SIZE,
                  "fields": "summary,status,assignee,created,priority,updated"}
        if token:
            params["nextPageToken"] = token
        data = jira_get(driver, "/rest/api/3/search/jql?" + urlencode(params))
        result.extend(data.get("issues", []))
        token = data.get("nextPageToken")
        if not token:
            return result
    raise RuntimeError("Incomplete query: MAX_PAGES reached; partial results were not applied.")


def assignment_time(driver, key, created, account_id):
    """Use the latest assignee change TO the current agent; creation only if never changed."""
    start, matching_times = 0, []
    while True:
        endpoint = (f"/rest/api/3/issue/{quote(key, safe='')}/changelog?"
                    f"startAt={start}&maxResults=100")
        data = jira_get(driver, endpoint)
        rows = data.get("values", [])
        for history in rows:
            for item in history.get("items", []):
                is_assignee = (item.get("fieldId") == "assignee" or
                               str(item.get("field", "")).lower() == "assignee")
                if is_assignee and item.get("to") == account_id:
                    matching_times.append(parse_date(history["created"]))
        start += len(rows)
        if not rows or start >= data.get("total", start):
            break
    return max(matching_times) if matching_times else parse_date(created)


def watch_jira():
    driver = None
    try:
        EVENTS.put(("info", "Opening Microsoft Edge..."))
        driver = launch_edge()
        driver.get(QUEUE_URL)
        print("Edge opened; waiting for an authorized corporate session.", flush=True)
        first_cycle = True
        while not STOP.is_set():
            try:
                EVENTS.put(("info", "Connecting to Jira..." if first_cycle else "Refreshing Jira page..."))
                # No refrescar inmediatamente en el primer ciclo: el SSO y la app
                # podrían seguir iniciando después de que aparezca la interfaz.
                if not first_cycle:
                    driver.refresh()
                WebDriverWait(driver, 90).until(
                    lambda d: (
                        urlparse(d.current_url).hostname == urlparse(JIRA_URL).hostname
                        and d.execute_script("return document.readyState") == "complete"
                    )
                )
                first_cycle = False
                print("Checking Jira API access...", flush=True)
                issues = fetch_issues(driver)
                print(f"Jira API connected: {len(issues)} tickets retrieved", flush=True)
                now = datetime.now(timezone.utc)
                pending, agents = [], {}
                for issue in issues:
                    fields = issue.get("fields") or {}
                    status = ((fields.get("status") or {}).get("name") or "")
                    assigned = fields.get("assignee") or {}
                    account_id = assigned.get("accountId")
                    person = assigned.get("displayName") or "Unassigned"
                    key = issue["key"]
                    summary = fields.get("summary") or ""
                    link = JIRA_URL + "/browse/" + key
                    if account_id:
                        agents.setdefault(account_id, {"name": person, "active": 0, "pending": 0})
                    if status == ACTIVE_STATUS:
                        if account_id:
                            agents[account_id]["active"] += 1
                        continue
                    if status not in PENDING_STATUSES:
                        continue
                    if account_id:
                        agents[account_id]["pending"] += 1
                        assigned_at = assignment_time(driver, key, fields["created"], account_id)
                        item_id = f"{key}:{account_id}:{assigned_at.isoformat()}"
                        elapsed_seconds = max(0, (now - assigned_at).total_seconds())
                    else:
                        assigned_at, item_id, elapsed_seconds = None, None, 0
                    row = {"key": key, "summary": summary, "person": person, "account_id": account_id,
                           "status": status, "assigned_at": assigned_at, "elapsed_seconds": elapsed_seconds,
                           "url": link, "item_id": item_id}
                    pending.append(row)
                EVENTS.put(("snapshot", (pending, agents, now, len(issues))))
                print(f"{datetime.now():%H:%M:%S} - {len(issues)} tickets checked, "
                      f"{len(pending)} waiting.", flush=True)
            except Exception as exc:
                details = traceback.format_exc()
                print("MONITOR ERROR:", details, flush=True)
                try:
                    LOG_FILE.write_text(details, encoding="utf-8")
                except OSError:
                    pass
                EVENTS.put(("error", f"{type(exc).__name__}: {exc}"))
            # Use Event.wait so closing the window doesn't leave a long sleep.
            if STOP.wait(REFRESH_SECONDS):
                break
    except Exception as exc:
        print("EDGE ERROR:", repr(exc), flush=True)
        traceback.print_exc()
        EVENTS.put(("error", str(exc)))
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass



ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")
PANEL = "#16383A"
PANEL2 = "#1C4445"
SUBTLE = "#A9C8C4"
TEXT = "#F5FAF8"
ACCENT = "#FF923F"

class MonitorApp:
    CLOCK = ("◴", "◷", "◶", "◵")

    def __init__(self, root):
        self.root = root
        self.rows = []
        self.agents = {}
        self.alerted = read_state()
        self.has_baseline = False
        self.last_success = None
        self.last_total = 0
        self.cells = {}
        self.sound = tk.BooleanVar(value=True)
        self.status_message = tk.StringVar(value="Connecting to Jira...")
        self.next_scan = None
        self.show_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        threading.Thread(target=watch_jira, daemon=True).start()
        self.root.after(300, self.consume)
        self.root.after(1000, self.tick)

    def label(self, parent, value, size=12, color=TEXT, weight="normal", **kw):
        return ctk.CTkLabel(parent, text=value, font=("Segoe UI", size, weight),
                            text_color=color, **kw)

    def show_ui(self):
        r = self.root
        r.title("Aya Healthcare | Jira Monitor · CR IT")
        r.geometry("810x530")
        r.minsize(770, 490)
        r.configure(fg_color=BG)

        header = ctk.CTkFrame(r, fg_color=BG, corner_radius=0)
        header.pack(fill="x", padx=20, pady=(14, 10))
        # El logo oficial tiene tipografía oscura: panel claro para mantener contraste.
        logo_panel = ctk.CTkFrame(header, fg_color="#FFFFFF", corner_radius=10,
                                  width=158, height=65)
        logo_panel.pack(side="left", padx=(0, 15))
        logo_panel.pack_propagate(False)
        try:
            picture = Image.open(LOGO_PATH).convert("RGBA")
            self.logo_image = ctk.CTkImage(light_image=picture, dark_image=picture,
                                           size=(138, 60))
            ctk.CTkLabel(logo_panel, text="", image=self.logo_image).place(
                relx=.5, rely=.5, anchor="center")
        except Exception:
            self.label(logo_panel, "Aya Healthcare", 15, "#0B464A", "bold").place(
                relx=.5, rely=.5, anchor="center")
        titlebox = ctk.CTkFrame(header, fg_color="transparent")
        titlebox.pack(side="left", fill="y", pady=7)
        self.label(titlebox, "Jira Monitor", 21, TEXT, "bold").pack(anchor="w")
        self.label(titlebox, "HELP DESK  /  COSTA RICA", 10, SUBTLE).pack(anchor="w")
        self.connection = self.label(header, "● Connecting", 11, ORANGE, "bold")
        self.connection.pack(side="right", padx=(0,5))

        stats = ctk.CTkFrame(r, fg_color="transparent")
        stats.pack(fill="x", padx=20, pady=8)
        self.metric_labels = []
        for idx,(name,accent) in enumerate((
            ("Waiting for support",ACCENT),
            (f"Overdue +{ALERT_MINUTES} min",RED),
            ("Busy · In Progress",TEAL),
            (f"No In Progress / {TEAM_SIZE}", "#72D7AF")
        )):
            card=ctk.CTkFrame(stats, fg_color=PANEL, corner_radius=12)
            card.grid(row=0,column=idx,padx=(0,8) if idx<3 else (0,0),sticky="ew")
            stats.grid_columnconfigure(idx,weight=1,uniform="metrics")
            n=self.label(card,"—",25,accent,"bold")
            n.pack(anchor="w",padx=14,pady=(9,0))
            self.label(card,name,11,SUBTLE).pack(anchor="w",padx=14,pady=(0,11))
            self.metric_labels.append(n)

        content = ctk.CTkFrame(r,fg_color="transparent")
        content.pack(fill="both",expand=True,padx=20,pady=(5,5))
        titlebar=ctk.CTkFrame(content,fg_color="transparent")
        titlebar.pack(fill="x",pady=(0,6))
        self.label(titlebar,"Tickets awaiting attention",14,TEXT,"bold").pack(side="left")
        self.label(titlebar,f"◷ {ALERT_MINUTES} min from assignment  ·  Sync {REFRESH_SECONDS} s",10,SUBTLE).pack(side="right")
        self.ticket_list=ctk.CTkScrollableFrame(content,fg_color=PANEL,corner_radius=12,
                                               height=245,scrollbar_button_color="#346769")
        self.ticket_list.pack(fill="x",expand=False)
        self.empty_tickets=self.label(self.ticket_list,"Waiting for first scan...",12,SUBTLE)
        self.empty_tickets.pack(pady=27)

        note = self.label(content,
            f"{TEAM_SIZE} team members  •  No In Progress ≠ confirmed availability", 10, SUBTLE)
        note.pack(anchor="w", pady=(10,4))

        footer=ctk.CTkFrame(r,fg_color=BG,corner_radius=0)
        footer.pack(fill="x",padx=20,pady=(7,13))
        self.status_label=self.label(footer,"Connecting...",10,SUBTLE,anchor="w",wraplength=450)
        self.status_label.pack(side="left",fill="x",expand=True)
        ctk.CTkSwitch(footer,text="Sound",variable=self.sound,progress_color=ACCENT,
                      button_color=TEXT,fg_color="#396363",font=("Segoe UI",11),
                      text_color=SUBTLE,width=105).pack(side="right",padx=(8,0))
        ctk.CTkButton(footer,text="Open Jira ↗",width=105,height=30,corner_radius=8,
                      fg_color=ACCENT,hover_color="#DF752C",text_color=NAVY,
                      command=lambda:webbrowser.open(QUEUE_URL)).pack(side="right")

    def consume(self):
        try:
            while True:
                kind,payload=EVENTS.get_nowait()
                if kind=="snapshot":
                    self.rows,self.agents,self.last_success,self.last_total=payload
                    self.next_scan=time.monotonic()+REFRESH_SECONDS
                    self.connection.configure(text="● Connected · Jira",text_color="#70DCAC")
                    self.paint_tickets()
                    # The initial scan shows old overdue tickets but does not flood popups.
                    if not self.has_baseline:
                        for row in self.rows:
                            if row.get("item_id") and row.get("assigned_at"):
                                if (datetime.now(timezone.utc)-row["assigned_at"]).total_seconds() >= ALERT_MINUTES*60:
                                    self.alerted[row["item_id"]] = "baseline"
                        write_state(self.alerted)
                        self.has_baseline = True
                elif kind=="info":
                    self.status_message.set(payload)
                elif kind=="error":
                    self.connection.configure(text="● Query error",text_color=RED)
                    self.status_message.set("Error: "+payload[:180]+" · details in jira_monitor_error.log")
                    print("APP ERROR:", payload, flush=True)
                elif kind=="alert":
                    self.show_alert(payload)
        except queue.Empty:
            pass
        self.root.after(300,self.consume)

    def clear(self,parent):
        for widget in parent.winfo_children():
            widget.destroy()

    def paint_tickets(self):
        self.clear(self.ticket_list)
        self.cells={}
        if not self.rows:
            self.label(self.ticket_list,"No tickets waiting for support",12,SUBTLE).pack(pady=24)
            return
        ordered=sorted(self.rows,key=lambda x: (x["assigned_at"] is None,
                x["assigned_at"] or datetime.max.replace(tzinfo=timezone.utc)))
        self.rows=ordered
        for row in ordered:
            line=ctk.CTkFrame(self.ticket_list,fg_color=PANEL2,corner_radius=9,height=51)
            line.pack(fill="x",padx=4,pady=3)
            line.pack_propagate(False)
            a=ctk.CTkFrame(line,fg_color="transparent")
            a.pack(side="left",fill="both",expand=True,padx=10,pady=5)
            self.label(a,row["key"]+"  ·  "+row["person"],12,TEXT,"bold",anchor="w").pack(fill="x")
            self.label(a,row["summary"][:58],10,SUBTLE,anchor="w").pack(fill="x")
            ctk.CTkButton(line,text="↗",width=30,height=29,fg_color="#28595A",
                          hover_color="#367476",command=lambda u=row["url"]:webbrowser.open(u)
                          ).pack(side="right",padx=(0,10))
            clock=self.label(line,f"◴  {ALERT_MINUTES:02d}:00",13,ACCENT,"bold",width=120)
            clock.pack(side="right",padx=4)
            self.cells[row["key"]]=clock

    def tick(self):
        now=datetime.now(timezone.utc)
        late=0
        for row in self.rows:
            widget=self.cells.get(row["key"])
            if not widget:
                continue
            if row["assigned_at"] is None:
                widget.configure(text="— Unassigned",text_color=SUBTLE)
                continue
            elapsed=max(0,int((now-row["assigned_at"]).total_seconds()))
            remaining=ALERT_MINUTES*60-elapsed
            if remaining<=0:
                late+=1
                # Trigger promptly at the 10-minute boundary, independently of the Jira refresh.
                assignment_id = row.get("item_id")
                if self.has_baseline and assignment_id and assignment_id not in self.alerted:
                    self.alerted[assignment_id] = datetime.now(timezone.utc).isoformat()
                    write_state(self.alerted)
                    self.show_alert(row)
                over=-remaining
                value=f"+{over//60:02d}:{over%60:02d}"
                color=RED
            else:
                value=f"{remaining//60:02d}:{remaining%60:02d}"
                color=ACCENT
            frame=self.CLOCK[int(time.monotonic()*2)%len(self.CLOCK)]
            widget.configure(text=f"{frame}  {value}",text_color=color)
        if self.last_success:
            self.metric_labels[0].configure(text=str(len(self.rows)))
            self.metric_labels[1].configure(text=str(late))
            occupied = sum(1 for x in self.agents.values() if x["active"] > 0)
            self.metric_labels[2].configure(text=str(occupied))
            self.metric_labels[3].configure(text=str(max(0, TEAM_SIZE - occupied)))
            count=max(0,int(self.next_scan-time.monotonic())) if self.next_scan else 0
            self.status_label.configure(text=f"Jira updated: {self.last_success.astimezone():%H:%M:%S}  ·  next refresh ~{count}s")
        else:
            self.status_label.configure(text=self.status_message.get())
        self.root.after(1000,self.tick)

    def show_alert(self,item):
        if self.sound.get():
            # Use a distinct beep, rather than the possibly-muted Windows event sound.
            def play_sound():
                try:
                    if winsound:
                        winsound.Beep(1050, 240)
                        winsound.Beep(1250, 280)
                    else:
                        self.root.after(0, self.root.bell)
                except (RuntimeError, OSError):
                    self.root.after(0, self.root.bell)
            threading.Thread(target=play_sound, daemon=True).start()

        toast=ctk.CTkToplevel(self.root)
        toast.title("Aya · Unattended ticket")
        toast.geometry("365x166")
        toast.resizable(False,False)
        toast.configure(fg_color=BG)
        toast.attributes("-topmost",True)
        self.root.update_idletasks()
        x=max(0,self.root.winfo_screenwidth()-390)
        y=max(0,self.root.winfo_screenheight()-235)
        toast.geometry(f"365x166+{x}+{y}")
        self.label(toast,"⚠  Ticket awaiting attention",15,ACCENT,"bold").pack(anchor="w",padx=18,pady=(13,5))
        elapsed=int((datetime.now(timezone.utc)-item["assigned_at"]).total_seconds()/60)
        self.label(toast,f'{item["person"]} · {item["key"]} · {elapsed} min',
                   12,TEXT,"bold",anchor="w").pack(anchor="w",padx=18)
        self.label(toast,item["summary"][:54],11,SUBTLE,anchor="w").pack(anchor="w",padx=18,pady=3)
        buttons=ctk.CTkFrame(toast,fg_color="transparent")
        buttons.pack(anchor="e",padx=15,pady=7)
        ctk.CTkButton(buttons,text="Open ticket ↗",width=115,height=28,
                      fg_color=ACCENT,text_color=NAVY,hover_color="#DA752F",
                      command=lambda u=item["url"]:webbrowser.open(u)).pack(side="left",padx=4)
        ctk.CTkButton(buttons,text="Close",width=64,height=28,
                      fg_color=PANEL2,command=toast.destroy).pack(side="left",padx=4)

    def close(self):
        STOP.set()
        self.root.destroy()


if __name__=="__main__":
    print("AYA JIRA MONITOR · waiting for Jira...",flush=True)
    root=ctk.CTk()
    MonitorApp(root)
    root.mainloop()
