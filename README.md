# greek-pharmacy-skills

Claude Code / Claude Agent skills for Greek pharmacies running the
**Euromedica / Europharmacy** management software (`EuromedicaTwoN.exe`, SQL Server backend).

They let you ask, in plain language, for **read-only** reports on your own pharmacy data —
daily sales and turnover, top products, prescriptions, suppliers/purchases, and stock-error
detection — and get answers, tables, and charts back.

> ⚠️ **Read-only by design.** Every query is a `SELECT` run under `READ UNCOMMITTED`, so it
> never writes to, nor locks, your live point-of-sale terminals. These skills never modify
> pharmacy data.

## Skills

| Skill | What it does |
|---|---|
| [`europharmacy-setup`](skills/europharmacy-setup/SKILL.md) | One-time setup: install **Tailscale**, locate the database (instance/static port/name), open the Windows Firewall **only over Tailscale**, create a read-only login, and fill in `.env`. |
| [`europharmacy-db`](skills/europharmacy-db/SKILL.md) | Connect and run the reports: sales, morning/day splits, top products (by units / value / distinct receipts), suppliers & purchases, and stock-error detection. Contains the DB schema map and ready-made SQL. |

New install → run `europharmacy-setup` first, then use `europharmacy-db`.

## Security model (read this)

- **No secrets in this repo.** Every site-specific value — server address, SQL login, password —
  lives in a local **`.env`** that is **gitignored**. The skill files contain only the shared
  database *schema* (the same for every install of this software), never credentials.
- **SQL Server is never exposed to the internet.** Remote access goes through
  [Tailscale](https://tailscale.com) (a private WireGuard network); the firewall is opened for the
  Tailscale interface only. No router port-forwarding.
- **Prefer a dedicated read-only SQL login** (`db_datareader`) for reporting, not the application's
  own login.
- The pharmacy's business data (sales/patient/financial reports and exports) is **also gitignored**
  — keep reports out of the repo.

If you ever commit a secret by accident, treat it as compromised: rotate the SQL password and
scrub history — a public git history is permanent.

## Setup

1. **Clone** this repo.
2. **Create your config:** `cp .env.example .env` and fill in your values. Recommended location:
   `~/.europharmacy/.env` (works from any project and stays out of every repo). The skills look for
   `.env` via `$EUROPHARMACY_ENV`, then `./.env`, then `~/.europharmacy/.env`.
3. **Run setup:** open the folder with Claude Code and ask it to run the `europharmacy-setup` skill.
   It walks you (and your server) through Tailscale, the firewall, and the read-only login.
4. **Report:** ask things like *"today's sales"*, *"top 3 products per day this month"*,
   *"which items have wrong stock"*.

### Making the skills discoverable

Claude Code auto-discovers skills under a `.claude/skills/` directory. Pick one:

- **Per project:** copy (or symlink) `skills/europharmacy-db` and `skills/europharmacy-setup` into
  your project's `.claude/skills/`.
- **All projects (user-level):** copy them into `~/.claude/skills/`.

Because all site-specific values are read from `.env` at runtime, the same skill files work
unchanged on any pharmacy's machine — only each site's `.env` differs.

## Requirements

- Windows (server runs SQL Server; the Euromedica app is Windows-only).
- PowerShell with `System.Data.SqlClient` (built into .NET Framework on Windows).
- [Tailscale](https://tailscale.com) on the server and any client.

## Disclaimer

Community skills, not affiliated with or endorsed by Euromedica/Europharmacy. Table and column
names are observed from the shipped database schema and may vary by version. Use read-only; verify
figures before relying on them for accounting or regulatory purposes.

## License

MIT — see [LICENSE](LICENSE).
