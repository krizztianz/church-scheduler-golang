# Schedule Generator – Church Service (Go) v2.4.11

A weekly **service staff scheduler** (07:00 & 10:00 services) based on **Master.xlsx** and **TemplateOutput.xlsx**. It prioritizes: **anti back-to-back** across Sundays, **Elder/Member composition** for **Ushers (Kolektan)** and **Duty Members (P. Jemaat)**, and limits for **Readers/Procantors/Musicians** per service. This version also adds a **one-line summary per service** when `-v` is enabled.

> Zero-setup workflow: place *Master.xlsx* and *TemplateOutput.xlsx*, run the binary/`go run`, get a ready-to-send `.xlsx` file.

---

## Feature Highlights

- **Two services** per Sunday: `07` and `10`.
- **Priority order**: 1) *Majelis Pendamping* (10:00 only), 2) composition for **Kolektan** & **P. Jemaat**, 3) **Lektor/Prokantor/Pemusik** with per-service limits, 4) other roles.
- **Anti back-to-back** preference (avoid assigning someone who just served last Sunday).
- **Relax phase** (enabled when `-noRelaxB2B=false`) to fill remaining slots.
- **Strict composition** (`-strictComposition`) to leave slots **empty** if Elder/Member quota isn't met.
- **Composition patterns** for Kolektan & P. Jemaat via compact codes `1a..4e` (see table).
- **RNG seed** for reproducible results (`-seed 42`).
- **Output .xlsx** to `~/Documents/JadwalPetugas` (default) or custom `-outdir`.
- **Verbose** mode prints selection details and a **one-line per-service summary** (Kolektan/P. Jemaat).

---

## Inputs

### 1) Master.xlsx (required)
- **Sheet `Petugas`** at minimum contains **Nama** and, optionally, **Penatua** plus eligibility columns (e.g., *Lektor*, *Prokantor*, *Pemusik*, *Kolektan*, *P. Jemaat*, etc.). Values may be `x`, `1`, `true`, or `ya`.
- **Sheet `MappingRole`** maps roles to columns in `Petugas` with:
  - **Role**
  - **Kolom Master** (alias: `Source`)
  - **Service**: `07` | `10` | `both`
  - **Slots07**, **Slots10** (optional, to override default slot counts)

### 2) TemplateOutput.xlsx (required)
- Must have a **`Jadwal Bulanan`** sheet.
- The first column A lists role labels (case-insensitive). **Majelis Pendamping** is matched by fuzzy label (contains “majel” & “pend”).

---

## Install & Build

```bash
go mod tidy
go build -o schedule-gen
# or run directly:
go run . -bulan Agustus -tahun 2025 -v
```

> Any Go version supporting modules and `excelize/v2` is fine. No database required.

---

## Files & Folders

- **Working dir**: `~/Documents/JadwalPetugas/`
  - `config/Master.xlsx` — runtime Master. If not present, the app will try to copy from `./Master.xlsx` (CWD) or from the **executable** folder.
  - Use `-forceMasterCopy` to **overwrite** `config/Master.xlsx` from (CWD/exe).
  - Use `-master "/custom/Master.xlsx"` to **directly override**.
- **Output**: defaults to `~/Documents/JadwalPetugas`, filename pattern:
  - `JadwalPetugas_<Month>_<HH>.<MM>.<SS>.xlsx`
- **Template** resolution order: current working directory → executable folder.

---

## Quick Start

```bash
# Full month (August 2025)
go run . -bulan Agustus -tahun 2025 -v

# Single date (e.g., 17 Aug 2025)
go run . -bulan 8 -tahun 2025 -tgl 17 -v

# Custom output dir + reproducible
go run . -bulan 8 -tahun 2025 -outdir "./output" -seed 42 -v
```

> `-bulan` accepts `1..12` or Indonesian names (Januari..Desember).

---

## Flags

| Flag | Type | Default | Range | Example | Description |
|---|---|---:|---|---|---|
| `-bulan` | string | *(required)* | `1..12` or `Januari..Desember` | `-bulan 8` | Month to generate (requires `-tahun`). |
| `-tahun` | int | *(required)* | > 0 | `-tahun 2025` | Year to generate (requires `-bulan`). |
| `-tgl` | int | 0 | 1..31 | `-tgl 17` | Single date mode; ignored if 0. |
| `-maxLektor` | int | 2 | 1..4 | `-maxLektor 3` | Max **Lektor** per service. |
| `-maxProkantor` | int | 2 | 1..3 | `-maxProkantor 3` | Max **Prokantor** per service. |
| `-maxPemusik` | int | 2 | 1..3 | `-maxPemusik 3` | Max **Pemusik** per service. |
| `-seed` | int64 | 0 | any | `-seed 42` | RNG seed; `0` = time-based random. |
| `-outdir` | string | `~/Documents/JadwalPetugas` | path | `-outdir "./output"` | Output folder (auto-created). |
| `-template` | string | `TemplateOutput.xlsx` | filename | `-template "Tpl.xlsx"` | Template filename to copy from. |
| `-master` | string | *(empty)* | path | `-master "/data/Master.xlsx"` | Direct **Master.xlsx** override. |
| `-forceMasterCopy` | bool | `false` | `true/false` | `-forceMasterCopy` | Overwrite `config/Master.xlsx` from (CWD/exe). |
| `-v` | bool | `false` | `true/false` | `-v` | Verbose + **one-line per-service summary**. |
| `-kolektanPattern` | string | `2b` | `1a..4e` | `-kolektanPattern 3a` | Elder/Member pattern for **Kolektan**. |
| `-pjemaatPattern` | string | `3a` | `1a..4e` | `-pjemaatPattern 2a` | Elder/Member pattern for **P. Jemaat**. |
| `-strictComposition` | bool | `false` | `true/false` | `-strictComposition` | Leave unmet quotas **empty** (no relax-any). |
| `-noRelaxB2B` | bool | `false` | `true/false` | `-noRelaxB2B` | Enforce anti back-to-back (disable relax phase). |

### Composition Codes (`1a..4e`)

Each code = total slots & **Elder (P)** vs **Member (J)** split.

| Code | P | J | Total | Code | P | J | Total |
|---|---:|---:|---:|---|---:|---:|---:|
| `1a` | 1 | 0 | 1 | `3a` | 1 | 2 | 3 |
| `1b` | 0 | 1 | 1 | `3b` | 2 | 1 | 3 |
| `2a` | 1 | 1 | 2 | `3c` | 3 | 0 | 3 |
| `2b` | 2 | 0 | 2 | `3d` | 0 | 3 | 3 |
| `2c` | 0 | 2 | 2 | `4a` | 1 | 3 | 4 |
|  |  |  |  | `4b` | 2 | 2 | 4 |
|  |  |  |  | `4c` | 3 | 1 | 4 |
|  |  |  |  | `4d` | 4 | 0 | 4 |
|  |  |  |  | `4e` | 0 | 4 | 4 |

---

## How It Works (Brief)

1. Collect all Sundays in the target month (or `-tgl` for single date).
2. Build candidate pools per role/service from `Petugas` (split **Elder/Member** for composition).
3. Prefer non-B2B assignments (avoid last-Sunday staff).
4. Fill per service in order:
   - **Majelis Pendamping** (10:00 only). If insufficient, relax by picking from those already serving at 07:00 (no double-role at 10:00).
   - **Kolektan & P. Jemaat** composition per pattern. With `-strictComposition`, leave remaining slots empty; otherwise try `relax-any`.
   - **Lektor/Prokantor/Pemusik** up to their `-max*`. If `-noRelaxB2B=false`, relax to fill.
   - Other roles follow default/overridden slot counts (`Slots07/Slots10`).
5. Write to `Jadwal Bulanan` in the template (role labels matched case-insensitively; MP matched fuzzily).

> With `-v`, the app logs **`Summary <svc>.00: Kolektan <status> | P.Jemaat <status>`** per date, plus composition and relax/strict notes.

---

## Troubleshooting

- **Missing `-bulan`/`-tahun`** → provide both flags.  
- **`Petugas`/`MappingRole` sheet missing/empty** → verify sheet names and headers.  
- **Master.xlsx not found** → place it in CWD or executable folder, or use `-master` / `-forceMasterCopy`.  
- **`role ... not found in template`** → ensure role labels in column A of `Jadwal Bulanan` match (case-insensitive). **Majelis Pendamping** uses fuzzy match.  
- **No Sundays found** → check `-bulan`; or use a valid `-tgl`.  

---

## Release Notes v2.4.11
- Added **one-line per-service summary** in verbose mode.
- Refined fill order & clearer verbose logging.

---

## License
Internal use. Adapt for your congregation's local needs.

_Last updated: 2025-08-23 04:34:21_