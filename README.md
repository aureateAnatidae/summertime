# summertime
Use a Python script to sum times of check-in/check-out in Google Sheets

This is a short script for my CMPT395 teammates to help them compile the time totals for a Google Sheet, which records the time of check-in and check-out on a timesheet.

If you're reading this, hey! This is an example of an "internal tool", or a tailor-made program used by the employees of an organization and not for the client.
For example, your teachers' report card system, or their attendance sheets would be internal tools.
Use `H10:I12` for the cell range when it prompts you to provide cells to write to.

---

## Setup

Prerequisites: `uv`
See the official [installation guide](https://docs.astral.sh/uv/#installation) for your operating system.

---

1. Go to `console.cloud.google.com/projectcreate`
2. Create your project
3. Click on `APIs & Services`
4. Click on `Library`
5. Find `Google Sheets API` and enable it
6. Return to `APIs & Services`, then click `Credentials`
7. Click `CREATE CREDENTIALS` and select `OAuth Client ID`
8. Select `Desktop App`
9. Download your OAuth 2.0 Client ID by clicking the `⇓` symbol to the right of your new Client ID
10. Move the folder to this directory's root and rename it to `credentials.json`
11. Install script dependencies by running `uv sync --frozen`
12. Run the script by running `uv run __main__.py`