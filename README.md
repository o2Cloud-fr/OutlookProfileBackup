# OutlookProfileBackup

**OutlookProfileBackup** is a fully autonomous C# console application that performs a **complete and automatic backup** of the current user's Outlook profile into a timestamped **ZIP** file saved directly to the **Desktop**. No user interaction required — just run it and let it work.

---

## 📜 Features

- 🔄 **Fully automatic** — zero user interaction required, runs from start to finish on its own
- 📧 **PST export via Outlook COM** — automatically launches Outlook in the background, copies all mailboxes (Exchange, Microsoft 365, IMAP, POP) into a portable `.pst` file
- 🗂️ **All mailboxes included** — every configured account is detected and exported automatically
- 📁 **PST file backup** — exports and includes all `.pst` data files in the ZIP
- ✍️ **Email signatures backup** — saves HTML, plain text, and associated images
- 📝 **Outlook templates backup** — backs up all `.oft` message templates
- 🖊️ **Stationery backup** — saves custom Outlook stationery files
- 📋 **Messaging rules backup** — exports `.rwz` rule files
- 🗝️ **Windows Registry export** — automatically exports `HKCU\Software\Microsoft\Office\{version}\Outlook\Profiles` to restore all accounts and profiles on a new machine
- 🔍 **Automatic Office version detection** — supports Office 2010, 2013, 2016, 2019, 2021 and Microsoft 365
- 🔒 **Outlook auto-close** — detects and closes Outlook before backup, waits for PST file to be fully unlocked before compression
- 🛡️ **Anti-overwrite protection** — automatically appends a timestamp if a ZIP with the same name already exists
- 📊 **Detailed console report** — step-by-step progress with color-coded status for every operation
- 📄 **RESTORE.txt included** — a complete restoration guide is automatically generated and included in the ZIP
- ⏱️ **Auto-close** — the tool closes itself automatically after 10 seconds
- 🧹 **Auto-cleanup** — temporary PST files are deleted after compression

---

## 🖥️ Prerequisites

- **Windows** 10 or later
- **.NET Framework** 4.8 or later
- **Microsoft Outlook** installed and configured (at least one profile)
- The tool handles opening and closing Outlook automatically

---

## 🚀 Installation

1. **Clone** this repository to your local machine:

```bash
git clone https://github.com/o2Cloud-fr/OutlookProfileBackup.git
```

2. **Open** the solution in Visual Studio.

3. Ensure the following references are present in the `.csproj`:

```xml
<Reference Include="System.IO.Compression" />
<Reference Include="System.IO.Compression.FileSystem" />
```

4. **Build** the project in `Release` mode.

5. **Run** the `.exe` — everything else is fully automatic.

---

## 🎯 Usage

Simply **double-click** the `.exe` (or run it from the command line). The tool will:

1. Detect and close Outlook if running
2. Connect to Outlook via COM and export all mailboxes to a PST file
3. Wait for the PST to be fully released before compressing
4. Create a ZIP archive on your Desktop containing everything
5. Display a full report and close automatically

> No prompts. No clicks. No configuration needed.

---

## 📂 ZIP File Structure

```
OutlookBackup_USERNAME_20260424_113333.zip
│
├── PST/
│   └── OutlookBackup_20260424_113333.pst     ← All mailboxes (Exchange, M365, IMAP...)
│
├── Signatures/
│   ├── MySignature.htm                        ← HTML signature
│   ├── MySignature.txt                        ← Plain text signature
│   └── MySignature_files/                     ← Associated images
│
├── Templates/
│   └── MyTemplate.oft                         ← Outlook message templates
│
├── Stationery/
│   └── MyStationery.html                      ← Custom stationery
│
├── Registry/
│   └── OutlookProfiles.reg                    ← All Outlook profiles and accounts
│
├── Rules/
│   └── MyRules.rwz                            ← Messaging rules
│
└── RESTORE.txt                                ← Complete step-by-step restore guide
```

---

## 🔄 Restore Guide

| Element | Procedure |
|---|---|
| **Profiles / Accounts** | Double-click `Registry\OutlookProfiles.reg` to import into Windows Registry |
| **PST Files** | Copy to `%LocalAppData%\Microsoft\Outlook\`, then in Outlook: *File → Open → Open Outlook Data File (.pst)* |
| **Signatures** | Copy contents of `Signatures\` to `%AppData%\Microsoft\Signatures\` |
| **Templates** | Copy contents of `Templates\` to `%AppData%\Microsoft\Templates\` |
| **Stationery** | Copy contents of `Stationery\` to `%AppData%\Microsoft\Stationery\` |
| **Rules** | Copy `.rwz` to `%LocalAppData%\Microsoft\Outlook\`, then in Outlook: *File → Manage Rules & Alerts → Options → Import Rules* |

> A `RESTORE.txt` file with the full step-by-step guide is automatically included in every ZIP.

---

## ⚠️ About PST Size and Exchange Accounts

If your PST appears small (e.g. 265 KB), this is expected for **Exchange / Microsoft 365** accounts — emails are stored server-side and the local copy only contains the folder structure.

To get a **full local copy** of all emails in the PST, either:

- Enable **Cached Exchange Mode** in Outlook before running the tool, or
- Manually export your mailbox via *File → Open & Export → Import/Export → Export to .pst → select mailbox → include subfolders*

---

## 🖥️ Console Output Example

```
  ╔══════════════════════════════════════════╗
  ║      OUTLOOK PROFILE BACKUP TOOL   v2    ║
  ║           Automatic Backup               ║
  ║         github.com/o2Cloud-fr            ║
  ╚══════════════════════════════════════════╝

  [INFO    ] Starting automatic backup...
  [INFO    ] Date    : 24/04/2026 11:33:33
  [INFO    ] Machine : HOSTNAME
  [INFO    ] User    : USERNAME

  ══ STEP 1/5 : CLOSING OUTLOOK ══
  [OK      ] Outlook is closed.

  ══ STEP 2/5 : PST EXPORT VIA OUTLOOK COM ══
  [INFO    ] Connecting to Outlook via COM...
  [INFO    ] Creating destination PST file...
  [INFO    ] 6 store(s) found. Copying...
  [OK      ] security@outlook.fr copied.
  [OK      ] admin@o2cloud.fr copied.
  [OK      ] PST exported (265 KB)

  ══ STEP 4/5 : COMPRESSION ══
  [PST     ] OutlookBackup_20260424_113333.pst  (265 KB copied to ZIP)
  [REG     ] OutlookProfiles.reg (Office 16.0) - 198 KB
  [TXT     ] RESTORE.txt generated.

  ══ STEP 5/5 : FINAL REPORT ══
  [SUCCESS ] Outlook backup completed successfully!
  ZIP File : C:\Users\USERNAME\Desktop\OutlookBackup_USERNAME_20260424_113333.zip
  Files    : 3
  Warnings : 0

  Closing in 10 seconds...
```

---

## 🛠 Tech Stack

- **Language** : C#
- **Framework** : .NET Framework 4.8
- **APIs used** :
  - `System.IO.Compression` — ZIP archive creation
  - `System.IO.Compression.FileSystem` — file-level ZIP operations
  - `Microsoft.Win32.Registry` — Windows Registry export
  - `System.Diagnostics.Process` — Outlook process management & `reg.exe` export
  - `System.Reflection` + `System.Runtime.InteropServices` — COM late-binding (no Interop reference needed)

---

## Authors

- [@MyAlien](https://www.github.com/MyAlien)
- [@o2Cloud](https://www.github.com/o2Cloud-fr)

---

## Badges

[![License](https://img.shields.io/badge/License-o2Cloud-yellow.svg)]()
[![C#](https://img.shields.io/badge/C%23-.NET%204.8-blue.svg)]()
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)]()
[![Outlook](https://img.shields.io/badge/Outlook-2010--365-0078D4?logo=microsoft-outlook&logoColor=white)]()

---

## Contributing

Contributions are always welcome!

See `contributing.md` for ways to get started. Please adhere to this project's `code of conduct`.

---

## Feedback

If you have any feedback, please reach out at [github@o2cloud.fr](mailto:github@o2cloud.fr)

---

## 🔗 Links

[![portfolio](https://img.shields.io/badge/my_portfolio-000?style=for-the-badge&logo=ko-fi&logoColor=white)](https://vcard.o2cloud.fr/)
[![linkedin](https://img.shields.io/badge/linkedin-0A66C2?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/remi-simier-2b30142a1/)

---

## Roadmap

- [ ] Windows Task Scheduler integration for scheduled automatic backups
- [ ] Graphical user interface (WinForms / WPF) with progress bar
- [ ] Backup to network share or custom destination path
- [ ] Windows toast notification on completion
- [ ] Multi-profile support on a single machine
- [ ] Backup encryption (password-protected ZIP)
- [ ] Email notification with backup summary

---

## Support

For support, email [github@o2cloud.fr](mailto:github@o2cloud.fr) or join our Slack channel.

---

## Used By

This project is used by the following companies:

- o2Cloud
- MyAlienTech

---

## License

[Apache-2.0 license](https://github.com/o2Cloud-fr/OutlookProfileBackup/blob/main/LICENSE)

---

![Logo](https://o2cloud.fr/logo/o2Cloud.png)
