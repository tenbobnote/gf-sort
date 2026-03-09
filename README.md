# GF Sort — Word Add-in for German Foundations

Sorts noun-group sections in the Resource Guide by the German Foundations sorting rules (ending → gender → mono/poly → alphabetical).

## Install (Mac — one command)

Open **Terminal** (press `Cmd + Space`, type "Terminal", hit Enter) and paste this:

```
curl -sL https://raw.githubusercontent.com/tenbobnote/gf-sort/main/install.sh | bash
```

This downloads the add-in and installs it into Word. That's it — no accounts, no passwords, nothing else to install.

### After installing

1. Open **Word**
2. Open the **Resource Guide** document
3. Look for **"GF Sort"** on the right side of the **Home** tab in the ribbon

If the button doesn't appear on the Home tab:
- Go to **Insert → Get Add-ins → MY ADD-INS → Shared Folder** → select **GF Sort**

## Install (Windows)

1. Download [`manifest-prod.xml`](https://raw.githubusercontent.com/tenbobnote/gf-sort/main/manifest-prod.xml) (right-click → Save Link As)
2. Open Word
3. Go to **Insert → Get Add-ins → MY ADD-INS → Upload My Add-in**
4. Browse to the downloaded `manifest-prod.xml` and click **Upload**

Note: On Windows, you'll need to re-upload the manifest each time you open Word unless your admin sets up a shared catalog folder.

## How to use

1. Click in any **noun-group table** in the Resource Guide (e.g., Animals, Beverages, Finance, etc.)
2. Click the **GF Sort** button in the ribbon (or open it from Insert → Add-ins)
3. The panel shows which section you're in (e.g., "Animals — Predictable")
4. Click **Sort Section** — the words are re-sorted by the German Foundations hierarchy
5. Comments are preserved automatically

## Updating

The add-in loads from the internet, so you always get the latest version automatically. No need to reinstall after updates.

If you experience issues after an update on Mac, clear the Word cache:
1. Quit Word
2. Open Terminal and run:
   ```
   rm -rf ~/Library/Containers/com.microsoft.Word/Data/Library/Caches/WebKit
   ```
3. Reopen Word

## Sortable sections

The add-in works on standard noun-group tables: Animals, Beverages, Chemical Elements, Colors, Fabrics, Finance, Fruits & Nuts, Gerunds, Landscape, Measurements, Metals & Materials, Numbers, Plants, Rocks, Weather.
