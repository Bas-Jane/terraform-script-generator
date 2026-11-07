# Terraform Script Generator — GitHub/manual edition

Version: 2026.11.07.3. Python 3.10+; works on Windows, macOS and Linux. The core converter uses the Python standard library; the desktop GUI needs Tkinter, which may require an OS-level installation.

## What this edition does

- Launch a desktop application to browse for a workbook, select worksheets/configurations and select output grouping.
- Run headless from a terminal, Windows PowerShell, VS Code, PyCharm or another Python IDE.
- Produce separate files, one file per sheet or one combined TFVARS file with a unique timestamped execution folder.
- Preserve dynamic CIQ parsing, defaults-as-comments, metadata comments, nested child lists and multirow target-value tables.
- Generate configuration **files only**; it does **not** execute `terraform apply`.

## Installation

1. Install **Python 3.10 or newer** and enable the `python`/`python3` command in your PATH. On Windows, install Python from python.org and include Tk/Tcl when prompted.
2. Extract this ZIP. Open the extracted root folder in VS Code, PyCharm or your terminal.
3. Optional but recommended: create and activate a virtual environment; no additional packages need to be installed for core conversion.

```bash
python -m venv .venv
# Windows PowerShell: .venv\Scripts\Activate.ps1
# macOS/Linux: source .venv/bin/activate
python -m pip install -r requirements.txt
```

**Tkinter check:** `python -m tkinter` (or `python3 -m tkinter`). On Ubuntu/Debian, install Tk if missing using your system package manager (`sudo apt install python3-tk`). On headless servers use CLI mode, not GUI.

## Run the desktop application

| Operating system | How to launch |
|---|---|
| Windows | Double-click `run_windows.bat`, or use `py -3 ibr_tfvars_generator.py` |
| macOS | Open Terminal and run `./run_macos.command` or `python3 ibr_tfvars_generator.py`. First run may require `chmod +x run_macos.command` |
| Linux | Run `./run_linux.sh` or `python3 ibr_tfvars_generator.py` |
| Any OS / IDE | Run `ibr_tfvars_generator.py` with **no arguments**, or with `--gui` |

The desktop app provides **Browse Workbook**, **Browse Output Folder**, **Scan Workbook**, processing scope, grouping mode and **Generate Terraform Files**.

## CLI examples (all operating systems)

Run from the extracted directory (replace the sample workbook filename with your own):

```bash
python ibr_tfvars_generator.py "input/your_workbook.xlsx" --list
python ibr_tfvars_generator.py "input/your_workbook.xlsx" --all --combine separate --output "output"
python ibr_tfvars_generator.py "input/your_workbook.xlsx" --sheets "5G SA" --combine sheet --output "output"
python ibr_tfvars_generator.py "input/your_workbook.xlsx" --configs "selection_method_nf_based_entry" --combine all --output "output"
python test_selection_method.py
```

On macOS/Linux use `python3` if `python` is not defined. You can also use `generate_cli_windows.bat` (Windows) or `./generate_cli_unix.sh` (macOS/Linux) with the same script arguments.

## Run through Visual Studio Code

1. Install **Python** and VS Code's Python extension. Open the extracted package as a folder.
2. Run **Python: Select Interpreter** in the Command Palette, choosing your Python/virtual environment.
3. Open **Run and Debug**, then choose **GUI — Terraform Script Generator**, **CLI — List configurations**, **CLI — Whole workbook**, or **Regression tests**. These are in `.vscode/launch.json`.
4. For CLI run/debug entries, replace `input/your_workbook.xlsx` with your actual file path, or copy/edit the configuration in `launch.json`.
5. Optionally run **Terminal → Run Task** and select the regression test or CLI list task.

In PyCharm or other IDEs, set `ibr_tfvars_generator.py` as the script, choose the project directory as working directory and leave its arguments blank to open the GUI. Supply workbook plus flags for CLI execution.

## Output and verification

The selected output destination receives **one** timestamped subdirectory named `Terraform_Output_YYYYMMDD_HHMMSS/` with `.auto.tfvars`, `summary.csv` and `manifest.json`. If the selected folder itself is named `Terraform_Output`, the generator places the run folder beside it for compatibility with prior releases.

Always check generated record counts, warnings in `manifest.json`, comments/defaults and variable schemas before use. Numeric MCC/MNC may require schema-specific string overrides. Don't commit private CIQ workbooks, subscriber IDs, credentials or actual production configuration files.

## GitHub repository use — no GitHub Actions

GitHub stores/version-controls the source code; **execution is manual on your computer** (Windows, macOS or Linux), or in your IDE/terminal after cloning. No workflow, GitHub Actions, web-hosted GUI, runners, tokens or secrets are required.

1. Create a **private** repository in your organization or personal account.
2. Upload the unpacked contents (including the `.vscode` folder and launchers). You can use GitHub's **Add file → Upload files** or Git from a terminal.
3. Clone it to the local machine: `git clone https://github.com/ORG/REPO.git`, then `cd REPO`.
4. Run the **desktop GUI** or CLI as explained above (or open in VS Code and press F5). Your workbook stays on your computer; use a local path or the ignored `input/` folder.
5. To publish source updates: `git add . && git commit -m "Update generator" && git push`. Never stage private input/output content.

**Important:** Clicking a source file on github.com does not run the generator. GitHub's website has no native desktop file browser, and without workflows or a hosted service it cannot perform conversion remotely. For remote manual execution, clone/run on a computer with Python.

See `GITHUB_MANUAL_SETUP.md` for a full walkthrough.

