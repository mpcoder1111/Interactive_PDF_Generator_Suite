# Security Policy

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| 1.0.x   | ✅ Yes              |

---

## Reporting a Vulnerability

This project is a **100% offline, standalone desktop application**. It does not connect to the internet, does not use a database, and does not transmit any data outside the local machine.

However, if you discover a security issue (e.g., a vulnerability in how PDF files are parsed, or in the executable packaging), please **do not open a public GitHub issue**.

Instead, report it privately by:

1. Opening a [GitHub Security Advisory](https://github.com/mpcoder1111/Interactive_PDF_Generator_Suite/security/advisories/new) on this repository (preferred).
2. Or emailing the repository owner directly via the contact on their [GitHub profile](https://github.com/mpcoder1111).

Please include:
- A description of the vulnerability
- Steps to reproduce
- Potential impact
- Any suggested mitigations (if known)

We aim to acknowledge your report within **7 days** and aim to release a fix within **30 days** for confirmed vulnerabilities.

---

## Dependency Security

Runtime dependencies and their licenses:

| Package | Version | License |
|---|---|---|
| `reportlab` | ≥ 5.0.0 | BSD-3-Clause |
| `openpyxl` | ≥ 3.1.5 | MIT |
| `pypdf` | ≥ 5.0.0 | BSD-3-Clause |
| `Pillow` | ≥ 10.0.0 | HPND (MIT-style) |

Developer-only (not shipped in the .exe runtime):

| Package | Version | License |
|---|---|---|
| `pyinstaller` | ≥ 6.0.0 | GPL-2.0 with Bootloader Exception |

We recommend keeping dependencies up to date. Run `pip install -r requirements.txt --upgrade` to update runtime dependencies.
