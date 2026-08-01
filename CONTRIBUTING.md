# Contributing to Schema-Driven Interactive PDF Form Suite

Thank you for your interest in contributing! This document explains how to get involved.

---

## 🐛 Reporting Bugs

1. Search [existing issues](https://github.com/mpcoder1111/Interactive_PDF_Generator_Suite/issues) to check if it is already reported.
2. If not, open a new issue with:
   - A clear title and description
   - Steps to reproduce the bug
   - Expected vs. actual behaviour
   - Your Python version and OS

---

## 💡 Suggesting Features

Open an issue with the `enhancement` label and describe:
- What problem you want to solve
- How you envision the feature working
- Any alternative approaches you considered

---

## 🔧 Setting Up the Development Environment

```bash
# 1. Clone the repository
git clone https://github.com/mpcoder1111/Interactive_PDF_Generator_Suite.git
cd Interactive_PDF_Generator_Suite/Interactive_PDF_Code

# 2. Install runtime dependencies
pip install -r requirements.txt

# 3. (Optional) Install developer dependencies to build .exe
pip install -r requirements-dev.txt

# 4. Run the application
python gui_app.py
```

---

## 📐 Coding Style

- Follow [PEP 8](https://pep8.org/) conventions.
- Use descriptive variable names — avoid single-letter names except for loop counters.
- Add docstrings to all functions, classes, and modules.
- Include inline comments for non-obvious logic, especially for AcroForm PDF manipulation and ReportLab canvas operations.
- Keep functions focused on a single responsibility.

---

## 🔀 Pull Request Process

1. **Fork** the repository and create a branch from `master`:
   ```bash
   git checkout -b feature/your-feature-name
   ```
2. Make your changes and ensure the application runs without errors.
3. Test with both the GUI (`python gui_app.py`) and CLI (`python pdf_generator_from_excel.py`).
4. Update `CHANGELOG.md` under `[Unreleased]` with a summary of your changes.
5. Submit a Pull Request with a clear description of what you changed and why.

---

## 📁 Project Architecture Overview

```
GUI Layer (gui_app.py)
    ↓
Configuration Layer (build_pdf_configurator_excel.py)
    ↓
PDF Engine (pdf_generator_from_excel.py)
    ↓
Extraction Engine (pdf_data_extractor.py)
```

This layered architecture keeps concerns separated — please maintain this boundary when contributing.

---

## 📄 License

By contributing, you agree that your contributions will be licensed under the [MIT License](LICENSE).
