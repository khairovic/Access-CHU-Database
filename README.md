# Access CHU Database — Clinical Research Project Management Tool

A Microsoft Access application built for **CIC 1432** (Clinical Investigation Centre) at **CHU Dijon-Bourgogne**. It replaces scattered Excel workbooks and e-mail exchanges with one tool that follows clinical research projects from the first idea to scientific dissemination.

The tool covers three activities of the CIC:

- **Coordination**: helping a project leader with methodology, funding, regulatory files, centres and data management.
- **Investigation**: running trials for external sponsors in the CIC units.
- **Volunteers and patients**: managing a database of healthy volunteers and patients who can be enrolled in studies.

## Features

- **One project record, ten lifecycle steps.** A single tabbed form (`Coordination_form`) covers General information, Feasibility, Funding, Regulatory, Methodology, Preparation, Progress and enrolment, End of study, Analyses and Dissemination.
- **Coordination and investigation in the same screens.** A tag system (`COORD` / `INVEST` / `COORD_INVEST`) shows only the fields that apply to each pathway, so the interface isn't duplicated.
- **Volunteer and patient module.** Record, pre-screening and enrolment, plus patient visit follow-up by randomisation arm.
- **Patient file import** from Excel (`.xlsx` / `.xls`) and CSV.
- **Indicators.** Every indicator has a stored reference date, so steering figures come from queries instead of manual counts.
- **Security.** Login screen, Administrator / User profiles, a permission matrix you can change without code, temporary passwords that expire after 24 hours, and account deactivation.
- **Audit log.** Time-stamped, named, plain-language record of user actions.
- **Extensible lists.** Users can add sponsors, departments, funders, CROs, etc. directly from the dropdowns.
- **Readable project reference** in the format `YYYY-NN`, reset every year.

## Tech stack

| Component | Technology |
|---|---|
| Database and UI | Microsoft Access (ACE engine) |
| Language | VBA |
| Data access | DAO |
| Excel import | ADO + `Microsoft.ACE.OLEDB.12.0` |
| CSV import | `Scripting.FileSystemObject` |
| E-mail (temporary passwords) | Outlook automation |

## Design notes

- **Normalised data model** centred on `Projet_table`: one concept per table, lookup tables for dropdowns, junction tables for many-to-many links (for example volunteers ↔ projects), and separate tables for repeating data such as allergies or medical history.
- **Unbound forms.** Each step subform is driven by code through three methods: `ChargerXxx` (load), `ValiderXxx` (validate) and `EnregistrerXxx` (save). This avoids Access's automatic saving and makes Cancel reliable.
- **Validate before saving.** The Save button validates every tab first, then saves them all, so a project is never left half-saved.
- **Conditional display.** `RefreshConditionalFields(ClearHidden)` shows or hides fields based on earlier answers and clears data that is no longer relevant.

## Repository contents

| File | Description |
|---|---|
| `ChuDB.accdb`, `ChuDB_V1.accdb` … `ChuDB_V6.accdb` | Successive versions of the database |
| `ChuDB_V6 - Compressed.7z`, `ChuDB_V7 - Compressed.7z` | Compressed archives of later versions |

The highest version number is the most recent.

## Getting started

1. Install **Microsoft Access** (Microsoft 365 or 2016+) on Windows.
2. Download the latest version. Extract `.7z` archives with [7-Zip](https://www.7-zip.org/).
3. Open the `.accdb` file and click **Enable Content** so the VBA code can run.
4. Log in with your account on the login screen.

### Multi-user deployment

For team use, split the database with Access's **Database Splitter** (Database Tools → Access Database):

- **Back end** (tables only) goes on a shared network folder with read/write access for CIC users.
- **Front end** (forms, queries, VBA) is copied to each user's computer.

Updates then only require redistributing the front end; the data stays untouched. Access works well for about 10–15 simultaneous users. For more, the tables can be migrated to SQL Server while keeping the Access front end.

## Known limitations

- Simultaneous editing of the same project is not locked.
- Cascade deletion is handled in VBA, since Access's index limit blocked native cascading foreign keys.
- Passwords are stored in the database and are not hashed yet.
- The volunteer database contains sensitive health data. Host it on a secure hospital network space with restricted access.

## Roadmap

- Summary dashboard built on the indicator configuration
- Consolidated cross-project follow-up tab
- Front end / back end split for production
- Password hashing
- Possible migration to a server database and web application

## Author

**Khaïri Hassine**, summer internship project (ESPRIT School of Engineering), supervised by Hatem El Moncer, at CIC 1432, CHU Dijon-Bourgogne, 2026.
