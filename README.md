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

## Shared attempt history

The shared Google Sheet is [Telangana SI Prep data](https://docs.google.com/spreadsheets/d/10umc1V8iAGp4wrD5IUpVuUeFgO_G8z9YFAjvj0CcRxA/edit). Its `Attempts` tab stores every completed attempt, including the question-level history used by the website.

To connect the website to it:

1. Open the Sheet, then choose **Extensions → Apps Script**.
2. Replace the default script with the contents of `Code.gs` from this repository and save it.
3. Choose **Deploy → New deployment → Web app**. Set **Execute as** to yourself and **Who has access** to anyone, then deploy and authorize it.
4. Copy the resulting `/exec` URL into `SYNC_URL` near the top of `app.js`, commit, and push that one-line change.

Both laptops will then load the same History and Progress data. The local browser copy remains as an offline fallback.
