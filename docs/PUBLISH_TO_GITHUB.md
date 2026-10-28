# Publishing the demo for the client

This folder is a GitHub-ready source repository. It intentionally contains no real staff personal data or production database.

## A public code link

1. Extract the ZIP on your computer.
2. Open a terminal in the `AccessStaffSync` folder.
3. If GitHub CLI (`gh`) is installed and signed in, run:

```bash
git init
git add .
git commit -m "Prototype: Access VBA staff update sync and sample tests"
gh repo create access-staff-sync --public --source=. --remote=origin --push
```

GitHub prints the new public repository URL. **Check the files for any private data before publishing.**

If `gh` is not available, create a new GitHub repository in your browser and upload the contents of this extracted folder via `Add file` → `Upload files`.

## Optional live HTML preview

The file `preview/index.html` is an interactive *illustration of the expected test outcomes* (it does not use a database). You can enable GitHub Pages on a public repo:

- Settings → Pages → Build and deployment → Deploy from a branch.
- Pick `main` and `/docs` (or move the preview into `/docs`), OR put `preview/index.html` at the repository root and select `main` + `/ (root)`.

Another simple option: upload `preview/index.html` to your static hosting provider. **Don't label the HTML walkthrough as a running Windows Access instance**.

## Client message (replace placeholder)

> I reviewed your workflow and built a Microsoft Access/VBA proof of concept specifically for your use case, including a 4,000-record sample database generator, 60-second import timer, locked-record retries, duplicate-event protection and import receipts. It also includes two Excel examples showing how the workflow handles valid, duplicated and invalid updates. You can review the source and setup here: YOUR_GITHUB_REPO_URL. This is a working prototype design; the next step is validating it on Windows and adapting the mappings to a sanitized copy of your existing database.
