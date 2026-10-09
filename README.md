# Hotcam Proposals CC check

Outlook add-in. When you send an email whose subject or new text mentions
**quote / quotation / budget / proposal**, and proposals@hotcam.tv isn't on it,
it adds proposals@ to CC and asks:

- **Send Anyway** – sends with proposals@ copied in
- **Don't Send** – back to the draft. Remove the CC if not needed and send again; it won't ask twice for the same email.

The quoted reply chain is ignored, so replying to an old thread that mentioned a budget won't trigger it.
Works in New Outlook (Mac and Windows), classic Outlook for Windows, and Outlook on the web.
If anything goes wrong inside the add-in, it lets the email send rather than blocking it.

## Files

| File | Purpose |
| --- | --- |
| manifest.xml | Uploaded to the M365 admin centre |
| commands.html, launchevent.js | The send check (hosted) |
| assets/icon-*.png | Icons (hosted) |
| index.html | Support page (hosted) |
| set-host.sh | Writes your hosting URL into manifest.xml |

## 1. Host the files (GitHub Pages, free)

1. Create a **public** GitHub repo called `hotcam-proposals-cc`.
2. Upload everything in this folder (except you can skip README/set-host.sh).
3. Repo **Settings › Pages › Build and deployment**: Source = *Deploy from a branch*, branch = `main`, folder = `/ (root)`. Save.
4. After a minute, check `https://<your-username>.github.io/hotcam-proposals-cc/launchevent.js` loads in a browser.

## 2. Point the manifest at it

```sh
cd hotcam-proposals-cc
chmod +x set-host.sh
./set-host.sh https://<your-username>.github.io/hotcam-proposals-cc
npx office-addin-manifest validate manifest.xml   # optional, Microsoft's validator
```

Commit the updated manifest.xml too (not required for hosting, but keeps the repo in sync).

## 3. Deploy to you and Pete

1. Go to **admin.microsoft.com › Settings › Integrated apps › Upload custom apps**.
2. App type **Office Add-in** › **Upload manifest file (.xml) from device** › pick `manifest.xml`.
3. Assign users: **Specific users/groups** › add yourself and Pete.
4. Accept permissions › **Finish deployment**.

It usually shows up within a few hours, but allow up to 24. Quit and reopen Outlook once it's there.
Pete doesn't need to install anything.

## 4. Test

New email to yourself, subject "Test quote". Hit Send: proposals@ should be added and the prompt shown.

## Changing it later

- Keywords: edit the `KEYWORDS` line in `launchevent.js` and push. Takes effect without redeploying (Outlook may cache for a few hours).
- Adding people: Integrated apps › Hotcam Proposals CC check › Users.
- Manifest changes (rare): bump `<Version>` in manifest.xml and use **Update** in Integrated apps.
