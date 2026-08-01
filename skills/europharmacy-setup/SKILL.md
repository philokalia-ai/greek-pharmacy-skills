---
name: europharmacy-setup
description: One-time setup so a pharmacy's Europharmacy/Euromedica SQL Server can be reached securely for read-only reporting. Walks through installing Tailscale, locating the database (server, SQL instance, static port, DB name), opening the Windows Firewall ONLY over Tailscale, and filling in the .env. Use when a pharmacist says the connection/reporting isn't set up yet, or asks to connect a new machine, set up Tailscale, or "open the firewall for the database".
---

# Europharmacy — Secure Remote Setup

Goal: let a client machine run **read-only** reports against the pharmacy's live SQL Server, reachable from anywhere, **without exposing SQL Server to the public internet**. The safe channel is **Tailscale** (a private WireGuard network); the firewall is opened for the Tailscale interface only.

Work through the steps in order. Several steps change system settings or the network and **require Administrator rights** — show the command, explain it, and let the user confirm before running. Never port-forward SQL on the router.

Two machines are involved:
- **Server** = the PC that runs SQL Server (where the Europharmacy database lives).
- **Client** = the PC that will run the reports / Claude.

---

## Step 1 — Locate the database (on the SERVER)

Confirm the SQL instance, static port, and database name.

```powershell
# SQL services / instances present on this machine
Get-Service | Where-Object { $_.Name -match 'MSSQL' } | Select-Object Status, Name, DisplayName
# Instance registry names
Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL' |
  Select-Object * -Exclude PS* | Format-List
```
For the matching instance (`MSSQLxx.<INSTANCE>` from the registry), read the static TCP port:
```powershell
$inst = 'MSSQL15.EUROPHARMACY'   # <-- replace with your instance's internal name
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$inst\MSSQLServer\SuperSocketNetLib\Tcp\IPAll" |
  Select-Object TcpPort, TcpDynamicPorts
```
- If `TcpPort` is a fixed number (e.g. a custom port like 25025) → that's your port.
- If only `TcpDynamicPorts` is set → set a **static** port instead (SQL Server Configuration Manager → Protocols → TCP/IP → IPAll → TcpPort = e.g. `25025`, clear TcpDynamicPorts, restart the SQL service). A static port is required so clients don't depend on SQL Browser.
- Make sure **TCP/IP protocol is Enabled** for the instance (same Configuration Manager screen).

Confirm the database name and that a SQL login works (ask the user for the SQL user/password — the Euromedica app has them in its config):
```powershell
$cs = "Server=localhost\EUROPHARMACY;Database=Europharmacy;User ID=<user>;Password=<pw>;TrustServerCertificate=True;Connect Timeout=5"
$cn = New-Object System.Data.SqlClient.SqlConnection $cs; $cn.Open()
$cn.CreateCommand().ExecuteScalar() | Out-Null; "OK: $($cn.State)"; $cn.Close()
```

> Recommended: create a dedicated **read-only** SQL login for reporting instead of reusing the app's login. As an admin in SSMS:
> ```sql
> CREATE LOGIN report_ro WITH PASSWORD = 'choose-a-strong-one';
> USE Europharmacy; CREATE USER report_ro FOR LOGIN report_ro;
> ALTER ROLE db_datareader ADD MEMBER report_ro;   -- SELECT only
> ```

## Step 2 — Install Tailscale (on BOTH machines)

1. Install from `https://tailscale.com/download` (Windows).
2. `tailscale up` and sign in — **use the same tailnet/account on both** the server and the client.
3. On the server, note its Tailscale address:
```powershell
& "$env:ProgramFiles\Tailscale\tailscale.exe" ip -4     # the 100.x.y.z address
& "$env:ProgramFiles\Tailscale\tailscale.exe" status    # names + online state of the tailnet
```
The server also has a stable **MagicDNS** name (`<host>.<tailnet>.ts.net`). Either the `100.x` IP or that name goes into `.env`.

> Tailscale gives an encrypted private link between your machines. Prefer it over any router/port-forward. It also sidesteps DHCP LAN-IP churn.

## Step 3 — Open Windows Firewall for SQL **only over Tailscale** (on the SERVER)

This allows the SQL port on the Tailscale interface only — not on the LAN, not on the internet.

```powershell
# Requires an elevated PowerShell. Replace 25025 with your static port from Step 1.
New-NetFirewallRule `
  -DisplayName "SQL Server (Europharmacy) via Tailscale" `
  -Direction Inbound -Action Allow -Protocol TCP -LocalPort 25025 `
  -InterfaceAlias "Tailscale" -Profile Any
```
If `-InterfaceAlias "Tailscale"` doesn't match on this machine, scope by the tailnet address range instead (Tailscale uses CGNAT `100.64.0.0/10`):
```powershell
New-NetFirewallRule `
  -DisplayName "SQL Server (Europharmacy) via Tailscale" `
  -Direction Inbound -Action Allow -Protocol TCP -LocalPort 25025 `
  -RemoteAddress 100.64.0.0/10 -Profile Any
```

> 🔒 **Do NOT**: create a rule with `-RemoteAddress Any`, open the port on the public/Domain profile for the LAN, or forward the port on the router. SQL Server exposed to the internet is a serious risk. Scope to Tailscale only.

Verify from the **client** (over Tailscale):
```powershell
Test-NetConnection <server-tailscale-ip> -Port 25025   # TcpTestSucceeded should be True
```

## Step 4 — Fill in `.env` (on the CLIENT)

Copy `.env.example` → `.env` (keep it private; it's gitignored). Put the **Tailscale** address as the default and the LAN hostname as fallback:
```
EUROPHARMACY_DB_SERVER=<server-tailscale-ip>,25025
EUROPHARMACY_DB_SERVER_LAN=<server-hostname>,25025
EUROPHARMACY_DB_NAME=Europharmacy
EUROPHARMACY_DB_USER=report_ro
EUROPHARMACY_DB_PASSWORD=********
EUROPHARMACY_DB_CONNSTR=Server=<server-tailscale-ip>,25025;Database=Europharmacy;User ID=report_ro;Password=********;TrustServerCertificate=True
EUROPHARMACY_DB_CONNSTR_LAN=Server=<server-hostname>,25025;Database=Europharmacy;User ID=report_ro;Password=********;TrustServerCertificate=True
```
Recommended location: `~/.europharmacy/.env`.

## Step 5 — Test end to end (on the CLIENT)

```powershell
$envFile = "$HOME\.europharmacy\.env"   # or wherever you saved it
$cfg = @{}; Get-Content $envFile | Where-Object { $_ -match '^\s*[^#].*=' } | ForEach-Object { $k,$v = $_ -split '=',2; $cfg[$k.Trim()] = $v.Trim() }
$cn = New-Object System.Data.SqlClient.SqlConnection ($cfg['EUROPHARMACY_DB_CONNSTR'] + ";Connect Timeout=10"); $cn.Open()
$c = $cn.CreateCommand(); $c.CommandText = "SELECT @@SERVERNAME, DB_NAME(), SUSER_SNAME(), GETDATE()"
$r = $c.ExecuteReader(); $r.Read() | Out-Null
"OK -> {0} / {1} / {2} / {3}" -f $r[0],$r[1],$r[2],$r[3]; $cn.Close()
```
Success → the `europharmacy-db` skill can now run reports (it resolves the same `.env`
automatically: `$EUROPHARMACY_ENV`, then any `.env` in the current directory **or a parent**,
then `~/.europharmacy/.env`).

### If it fails
| Symptom | Likely cause |
|---|---|
| `TcpTestSucceeded: False` from the client | Firewall rule missing/mis-scoped, or wrong port (Step 1/3) |
| Connects on LAN but not Tailscale | Server not signed into the same tailnet, or rule bound to the wrong interface |
| `Login failed for user` | Wrong SQL login/password, or the login lacks access to the database |
| `login is from an untrusted domain` | You used Windows auth — this setup needs **SQL auth** |
| Works, then breaks after a reboot | The instance was on a *dynamic* port — set a static one (Step 1) |

## Security checklist
- [ ] Reporting uses a **read-only** SQL login (`db_datareader`), not an admin/app login.
- [ ] Firewall rule is scoped to **Tailscale only** (interface or `100.64.0.0/10`) — never `Any`, never router port-forward.
- [ ] `.env` holds the only copy of the password on the client, is **gitignored**, and is not shared.
- [ ] Tailscale ACLs (admin console) limit which devices may reach the server, if you want tighter control.
