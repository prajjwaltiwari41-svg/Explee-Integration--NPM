# Explee hot leads → Google Sheets

Every 10 minutes, new Explee hot leads (project 50789) are added to the
**Hot Leads** tab of a Google Sheet. If a lead replies again, their row is
updated (matched on Person ID) instead of being added twice.

It runs as a Google Apps Script inside the Sheet, on Google's own timer.
You don't need a server and nothing has to stay open.

## One-time setup
1. Open the Sheet → **Extensions → Apps Script**.
2. Delete what's there and paste in `apps-script/HotLeadsSync.gs`. Put your Explee API key in `EXPLEE_API_KEY`.
3. Click **Save**, choose **setup** in the function dropdown, then click **Run**. Approve Google's permission prompt.

## Good to know
- Add your own columns (Status, Owner…) to the right of **Last Synced**. The sync never touches them.
- The first run brings in every hot lead you already have. After that, only new ones are added.
- To re-import everything, run `resetCursor` once and then `syncHotLeads`.
- If a run fails (for example, out of Explee credit), Google emails you.
