# Statement Snip for Excel: setup

Setup is a one-off and takes about 15 minutes. Afterwards, a **Statement Snip** button sits on Excel's Home tab.

## How it works and what's private

Excel add-ins have to be loaded from a web address, so the add-in's *code* (the files in this folder) is hosted on a free website. Your client data is never part of that. Screenshots are read inside the Excel side panel on your computer and written straight into your workbook. The text-reading engine (about 10 MB) downloads once from a public code library the first time you use it.

The hosted files contain no client information, so it's fine for them to be public.

## Step 1: Host the folder (free, with GitHub Pages)

1. Sign up at github.com if you don't have an account.
2. Create a new repository called `statement-snip` and set it to **Public**.
3. Click **Add file → Upload files** and drag in everything from this folder: `taskpane.html`, `commands.html`, `manifest.xml` and the `assets` folder. Commit.
4. Go to **Settings → Pages**, set Source to **Deploy from a branch**, choose **main** and **/ (root)**, then Save.
5. After a minute your site address appears, for example `https://yourname.github.io/statement-snip`.
6. Check it works by opening `https://yourname.github.io/statement-snip/taskpane.html`. You should see the panel.

## Step 2: Point the manifest at your site

Open `manifest.xml` in Notepad (Windows) or TextEdit (Mac). Use Find & Replace to change every `https://YOUR-SITE` to your site address from step 1, with no slash at the end. Save it.

(Or send me the address and I'll do it for you.)

## Step 3: Load it into Excel

### Windows (Microsoft 365 desktop)
1. Make a folder, for example `C:\ExcelAddins`, and put the edited `manifest.xml` in it.
2. Right-click the folder → **Properties → Sharing → Share**, share it with yourself, and note the **Network Path** (for example `\\YOUR-PC\ExcelAddins`).
3. In Excel go to **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**. Paste the network path into **Catalog Url**, click **Add catalog**, tick **Show in Menu**, then OK.
4. Close and reopen Excel.
5. Go to **Home → Add-ins → More Add-ins** (in some versions, **Insert → Get Add-ins**), open the **Shared Folder** tab, choose **Statement Snip**, then **Add**.

### Mac
1. In Finder press **Cmd+Shift+G** and go to `~/Library/Containers/com.microsoft.Excel/Data/Documents/`.
2. Create a folder called `wef` if it isn't there, and put the edited `manifest.xml` in it.
3. Restart Excel. Statement Snip appears under **Home → Add-ins** (or **Insert → My Add-ins**).

### Excel in a browser
**Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in**, then choose `manifest.xml`.

## Using it

1. Click **Statement Snip** on the Home tab to open the panel.
2. Snip the statement (Win+Shift+S, or Cmd+Ctrl+Shift+4 on Mac). Click the box in the panel and paste.
3. Drag a box around the part you want, if needed, then click **Read**.
4. Fix any amber cells. Press Enter on a cell to mark it as checked.
5. Add client names under **Names to strip**.
6. Click any cell in your sheet, then **Insert at selected cell**. Or use **Insert as new sheet**.

## Updating it

When I send you a new `taskpane.html`, upload it to the same GitHub repository to replace the old one. Excel picks it up the next time you open the panel, and you don't need to reinstall.

## If something goes wrong

- **The button doesn't appear.** Restart Excel. On Windows, check that **Show in Menu** is ticked for the catalog.
- **A blank panel or an error.** Open the taskpane.html address in a browser to check the site is live, and that the manifest addresses match it exactly.
- **On a managed work computer.** IT policies often block add-ins you install yourself. Use it on your own business machine instead.

## Servicing workbook mode

Open a Pavilion servicing workbook and Statement Snip shows a **Servicing workbook** card:

1. Fill in the Structure tab, then click **Build tabs from Structure**. One tab is created per party, and unused rows on the Servicing tab are hidden.
2. Snip a P&L, balance sheet or tax return and click **Read**.
3. Under **Send to workbook**, check the party, the statement and the year columns. Each line is matched to a row on the party's tab, and any unmatched line is highlighted.
4. Click **Write to tab**. Only figures are written, never names. For a P&L, the panel checks the result against the reported net profit.
