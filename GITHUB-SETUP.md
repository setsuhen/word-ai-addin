# GitHub Pages Setup – Step by Step

Get your Word AI add-in running on Word Online. No coding required.

---

## Step 1: Create a GitHub account

If you don't have one: go to [github.com](https://github.com) → Sign up

---

## Step 2: Create a new repository

1. Log in to GitHub
2. Click the **+** (top right) → **New repository**
3. Set:
   - **Repository name:** `word-ai-addin` (or any name you like)
   - **Visibility:** Public
   - **Do NOT** check "Add a README" or any other files
4. Click **Create repository**
5. Right away, go to **Settings** → **Pages** (left sidebar)
6. Under **Build and deployment**, set **Source** to **GitHub Actions**

---

## Step 3: Configure Git (first time only)

If you've never used Git on this computer, run these once in PowerShell (use your name and email):

```powershell
git config --global user.name "Your Name"
git config --global user.email "your-email@example.com"
```

---

## Step 4: Push your project

Open **PowerShell**, go to the word-ai-addin folder, and run (replace `YOUR-USERNAME` with your GitHub username):

```powershell
cd "c:\Users\aholl\OneDrive\Desktop\cursor projects\word-ai-addin"
```

```powershell
git init
```

```powershell
git add .
```

```powershell
git commit -m "Initial commit - Word AI add-in"
```

```powershell
git branch -M main
```

```powershell
git remote add origin https://github.com/YOUR-USERNAME/word-ai-addin.git
```

```powershell
git push -u origin main
```

**If GitHub asks for a password:** Use a [Personal Access Token](https://github.com/settings/tokens) (create one with `repo` access) instead of your password.

---

## Step 5: Wait for the deploy

1. Go to your repo → **Actions** tab
2. Wait for "Deploy to GitHub Pages" to finish (green checkmark)
3. Your add-in will be live at: `https://YOUR-USERNAME.github.io/word-ai-addin/`

---

## Step 6: Add the add-in in Word Online

1. Go to [Word on the web](https://www.office.com/launch/word)
2. Open or create a document
3. **Insert** → **Add-ins** → **Add from URL**
4. Paste this (replace `YOUR-USERNAME` and `word-ai-addin` if you used different names):

   ```
   https://YOUR-USERNAME.github.io/word-ai-addin/manifest.xml
   ```

5. Click **Add**
6. The add-in appears in the **Home** tab as **AI Assistant**
7. Click it, enter your OpenAI API key, and start using it

---

## Sharing with colleagues

Send them the manifest URL:

```
https://YOUR-USERNAME.github.io/word-ai-addin/manifest.xml
```

They add it the same way: **Insert** → **Add-ins** → **Add from URL** → paste the link.

---

## Troubleshooting

**"Workflow failed"** – Check the Actions tab for details. Ensure Pages is set to **GitHub Actions** in Settings → Pages.

**Add-in doesn't load** – Wait a few minutes after the workflow finishes. Confirm the manifest URL opens in a browser.

**Git push asks for login** – Use a Personal Access Token instead of your password, or try [GitHub Desktop](https://desktop.github.com/).
