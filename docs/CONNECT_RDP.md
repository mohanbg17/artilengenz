# Connecting to SAP S/4HANA via Remote Desktop (No Admin Rights)

This guide covers every scenario for connecting the NLP-SAP engine to a live
SAP S/4HANA system that is only accessible via Remote Desktop (RDP).

---

## Quick Decision Tree

```
Do you know the SAP hostname and OData port?
  │
  ├─ YES ──► Can you reach it from your local machine?
  │            │
  │            ├─ YES ──► Scenario A: Direct connection (simplest)
  │            └─ NO  ──► Scenario B or C below
  │
  └─ NO  ──► Step 1: Run the discovery script inside RDP first
```

---

## Step 1 — Discover SAP Connection Details (Run Inside RDP)

1. **Open the Remote Desktop session** (RDP into the machine where SAP GUI runs)
2. **Copy `scripts/discover_sap_rdp.bat`** to the RDP desktop (e.g., via clipboard paste or shared drive)
3. **Double-click it** — no admin rights needed
4. **Screenshot or copy the output**, especially:
   - The **hostname** and **open ports**
   - The **SAPUILandscape.xml** entries (server name, system ID, client)
   - The **browser test URLs** shown at the end

### What the script finds

| Info | Where it looks |
|------|---------------|
| SAP server hostname | `SAPUILandscape.xml`, registry `SAPGUI` keys |
| ICM open ports | TCP probe of ports 44300, 443, 8000, 8080, etc. |
| SAP client number | Registry / landscape XML |
| Running SAP processes | `tasklist` (shows if SAP runs locally or remotely) |

### Browser test inside RDP
After running the script, open a browser **inside the RDP** and try:
```
https://<SAP-HOST>:44300/sap/opu/odata/IWFND/CATALOGSERVICE?$format=json
```
- **Login prompt / JSON data** → OData is up. Note the host and port.
- **Connection refused** → try port 8000 (HTTP) or ask your Basis team.
- **Works in RDP browser but not locally** → use Scenario B or C.

---

## Scenario A — Direct Connection (SAP Reachable from Your Machine)

Use this when the SAP OData port is reachable directly from wherever
the NLP engine runs (same network, VPN, etc.).

### 1. Configure .env
```bash
cp .env.s4hana.example .env
```
Edit `.env` with values from the discovery script:
```env
SAP_HOST=s4h-dev.mycompany.com   # hostname from discover script
SAP_HTTP_PORT=44300               # open HTTPS port
SAP_HTTPS=true
SAP_CLIENT=100                    # from SAP logon screen
SAP_USERNAME=NLP_QUERY
SAP_PASSWORD=YourSAPPassword1!
SAP_SYSTEM_TYPE=S4HANA
MOCK_SAP=false
ANTHROPIC_API_KEY=sk-ant-...
```

### 2. Test the connection
```bash
python scripts/test_sap_connection.py
```
Expected output:
```
✓ s4h-dev.mycompany.com:44300 is reachable
✓ OData endpoint is up — HTTP 200 on /sap/opu/odata/IWFND/CATALOGSERVICE
✓ Authenticated as 'NLP_QUERY' on client 100 (312ms)
✓ API_GLACCOUNTLINEITEM_SRV / A_GLAccountLineItem
! API_PURCHASEORDER_PROCESS_SRV — HTTP 404: Service not activated
  → Activate in SAP: /IWFND/MAINT_SERVICE → add 'API_PURCHASEORDER_PROCESS_SRV'
```

### 3. Start the engine
```bash
PYTHONPATH=src python -m uvicorn api.main:app --reload --port 8080
```

### 4. Verify via API
```bash
curl http://localhost:8080/api/v1/connect
```

---

## Scenario B — Run the Engine Inside the RDP Session

Use this when SAP is **not** reachable from your local machine.
Run the entire NLP engine inside the Remote Desktop — no admin needed.

### Inside the Remote Desktop:

**1. Install Python 3.11 (no admin — user install)**
```
https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe
```
During install → check **"Add Python to PATH"** → choose **"Install for current user only"**

**2. Copy the project into the RDP session**
Option A — Git clone (if git is available):
```cmd
git clone https://github.com/mohanbg17/artilengenz.git
cd artilengenz
```
Option B — Copy via RDP clipboard/shared drive.

**3. Install dependencies**
```cmd
cd artilengenz
pip install -e ".[dev]" --user
```

**4. Configure .env inside RDP**
```cmd
copy .env.s4hana.example .env
notepad .env
```
Set `SAP_HOST` to **localhost** or the hostname shown by `hostname` command —
since SAP runs on the same machine as the RDP session.

**5. Test connection**
```cmd
python scripts/test_sap_connection.py
```

**6. Start the API server inside RDP**
```cmd
set PYTHONPATH=src
python -m uvicorn api.main:app --host 0.0.0.0 --port 8080
```

**7. Access from your local machine**
RDP forwards TCP ports. From your local machine:
```bash
curl http://<RDP-MACHINE-IP>:8080/api/v1/connect
```
Or configure RDP to forward port 8080 (Remote Desktop Connection → Local Resources → More → Local devices and resources).

---

## Scenario C — SSH Port Forward (No Admin on RDP Box)

Use this if you have **SSH access** to the RDP machine (many enterprise Windows
machines have OpenSSH Server enabled without needing admin).

### Check if SSH is available inside RDP:
```cmd
ssh localhost
```
If it connects → SSH is available.

### From your local machine, set up a tunnel:
```bash
# Forward SAP's OData port (44300) to your local port 14430
ssh -L 14430:<SAP-HOST>:44300 <your-rdp-username>@<RDP-MACHINE-IP> -N
```
Leave this terminal open.

### Configure .env on your local machine:
```env
SAP_HOST=localhost
SAP_HTTP_PORT=14430
SAP_HTTPS=true
SAP_CLIENT=100
SAP_USERNAME=NLP_QUERY
SAP_PASSWORD=YourSAPPassword1!
MOCK_SAP=false
```

### Test and run:
```bash
python scripts/test_sap_connection.py
PYTHONPATH=src python -m uvicorn api.main:app --port 8080
```

---

## Activating OData Services in SAP (No Admin Needed)

Some SAP services need to be activated by a Basis admin once.
Ask your Basis team to run transaction `/IWFND/MAINT_SERVICE` and activate:

| Service Name | Used For |
|---|---|
| `API_GLACCOUNTLINEITEM_SRV` | GL balance queries |
| `API_PURCHASEORDER_PROCESS_SRV` | Purchase order queries |
| `API_SALES_ORDER_SRV` | Sales order queries |
| `CATALOGSERVICE` (already active usually) | Service discovery |

**One-time Basis request template:**
```
Hi [Basis Team],

Please activate the following OData services in system [SID] client [CLIENT]
via transaction /IWFND/MAINT_SERVICE:

1. API_GLACCOUNTLINEITEM_SRV
2. API_PURCHASEORDER_PROCESS_SRV
3. API_SALES_ORDER_SRV

These are standard SAP Fiori services required for a read-only analytics
tool. No custom development or transports needed.

Thanks
```

---

## SAP User Setup (Ask Basis — No Admin Needed)

You need Basis to create a read-only service user. Send them this:

```
Please create an SAP user with the following:

User ID:   NLP_QUERY
Type:      System (background)
Password:  [you choose]
Client:    [your client]

Roles to assign:
  /IWFND/RT_GW_USER          (OData Gateway access)
  SAP_BC_DWB_WBDISPLAY       (generic display)

Additional authorisation objects (if not covered by roles):
  S_SERVICE  SRV_NAME = *
  F_BKPF_BUK BUKRS = * ACTVT = 03
  F_BKPF_KOA KOART = * ACTVT = 03
  M_BEST_BSA BSART = * ACTVT = 03
  V_VBAK_VKO VKORG = * ACTVT = 03

This user needs READ-ONLY access (ACTVT=03) only.
```

---

## Troubleshooting

| Error | Likely Cause | Fix |
|-------|-------------|-----|
| `Connection refused` on port 44300 | SAP ICM HTTPS not enabled or wrong port | Try port 8000 (HTTP) or check SMICM in SAP |
| `HTTP 401` | Wrong username/password | Double-check SAP_USERNAME and SAP_PASSWORD |
| `HTTP 403` | User missing OData authorization | Ask Basis to assign `/IWFND/RT_GW_USER` role |
| `HTTP 404` on service | Service not activated | Ask Basis to activate via `/IWFND/MAINT_SERVICE` |
| SSL certificate error | Self-signed SAP certificate | The engine sets `verify=False` by default for dev |
| `sap-client header rejected` | Wrong client number | Check SAP_CLIENT matches your logon client |
| Empty result set | User has no data authorization | Ask Basis to check `F_BKPF_BUK` authorization object |
| `ICM_HTTP_CONNECTION_BROKEN` | Network timeout | Increase timeout or check SAP workload |

---

## Testing With Mock Data First

Before touching a live SAP system, always test with mock data:
```env
MOCK_SAP=true
```
All 52 unit tests run without any SAP or Anthropic connection:
```bash
PYTHONPATH=src python -m pytest tests/ -v
```
