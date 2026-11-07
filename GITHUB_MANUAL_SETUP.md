# GitHub source repository — manual execution only

**This package intentionally has no `.github/workflows` folder.** GitHub is used for version control; the Python program runs on the user's computer after clone/download.

## Upload the repository

1. Create a **private** repository in GitHub.
2. Extract the GitHub ZIP, then upload its content to the repository root (not a nested ZIP directory).
3. Either use the GitHub website to upload files, or open an integrated terminal in VS Code:

```bash
git init
git add .
git commit -m "Add portable Terraform Script Generator"
git branch -M main
git remote add origin https://github.com/ORG/REPOSITORY.git
git push -u origin main
```

For an existing repository, clone it first and copy the package files into the clone, then add/commit/push. Replace the sample URL; do not copy it verbatim.

## Clone and run

```bash
git clone https://github.com/ORG/REPOSITORY.git
cd REPOSITORY
python ibr_tfvars_generator.py
```

On macOS/Linux, use `python3` where necessary. Or run Windows `.bat`, macOS `.command`, Linux `.sh`. GitHub login is needed only for repository access/commits, not to run the Python code.

## Visual Studio Code

Open the cloned directory using **File → Open Folder**, select your Python interpreter and use **Run and Debug → GUI — Terraform Script Generator → F5**. The GUI needs desktop Tkinter. For automated IDE execution, choose CLI debug profiles and edit the sample workbook path. Check `README.md`.

## Keeping private content out of GitHub

The `.gitignore` excludes `input/*`, `output/*`, `*.tfvars` and state files. Keep sensitive CIQ data outside the source repository. Git ignore does not untrack a file already committed. To stop tracking a committed file: `git rm --cached path/to/file` (then commit); removing it from repository history requires additional cleanup.

## No workflow execution

There are no GitHub Actions, runners, webhooks, tokens, scheduled jobs or Terraform apply commands. To run the converter, download or clone the repository and execute it manually. The GitHub website alone does not execute Python files.
