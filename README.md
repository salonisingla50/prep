# Telangana Police SI practice exams

This is a static, mobile-first practice-exam site for the Telangana Police SI preliminary exam on 29 November 2026. It uses plain HTML, CSS, and JavaScript, and runs directly on GitHub Pages.

## Deploy to GitHub Pages

1. Push these files to the repository's `main` branch.
2. In GitHub, open **Settings → Pages**.
3. Set the source to **Deploy from a branch**, choose `main` and the `/ (root)` folder, then save.
4. Open the Pages URL GitHub provides. Do not open `index.html` directly from the file system, because browsers block the workbook fetch in that mode.

## Add an exam

1. Open `questions.xlsx` and add a row for each question on the `Questions` sheet. Use the columns `Exam`, `Section`, `Question`, `A`, `B`, `C`, `D`, and `Answer`.
2. Set `Exam` to the exam number (1–30), keep the questions for each section together, and enter `A`–`D` in `Answer`. For two-option questions, leave both `C` and `D` blank.
3. Add or update that exam's `Unlock Date` and `Minutes Per Section` on the `Exams` sheet. Missing schedule rows use the app defaults.
4. Commit and push the updated workbook. The website fetches it with a cache-buster on every load.

## Attempt data

Completed attempts are intentionally stored in the learner's browser using localStorage, so the static site needs no backend and attempts remain private. The `Attempts` workbook sheet is a blank reference/archive layout; GitHub Pages cannot automatically write browser data to it.
