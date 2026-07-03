# SQL Script Generator — file structure

## Why split one file into five?

The original was one 650-line file doing everything: colors, SQL-building
logic, two reusable UI pieces, and the main app, all tangled together.
That's fine for a quick prototype, but it gets hard to navigate once it
grows — you end up scrolling past 200 lines of unrelated JSX just to find
the one function you need to fix.

Splitting by **responsibility** (not by size) is the standard way real
front-end codebases are organized. Each file below answers one question.

```
sql-script-generator/
└── src/
    ├── theme.js                       "What colors does this app use?"
    ├── utils/
    │   └── sqlHelpers.js               "How do I turn a row into SQL?"
    ├── components/
    │   ├── MultiSelectDropdown.jsx     "How does the column picker work?"
    │   └── DataPreviewModal.jsx        "How does the data-preview popup work?"
    └── SqlScriptGenerator.jsx          "How do all the pieces fit together?"
```

## How the files connect (the `import` chain)

```
SqlScriptGenerator.jsx
   ├── import { theme } from "./theme"
   ├── import { buildTableScript, downloadTextFile } from "./utils/sqlHelpers"
   ├── import { MultiSelectDropdown } from "./components/MultiSelectDropdown"
   └── import { DataPreviewModal } from "./components/DataPreviewModal"

MultiSelectDropdown.jsx
   └── import { theme } from "../theme"      (note: "../" — one folder up)

DataPreviewModal.jsx
   └── import { theme } from "../theme"
```

Two things worth noticing as a fresher, since this trips everyone up early:

1. **`export` / `import` must match.** `theme.js` does
   `export const theme = {...}`, so anywhere that wants it writes
   `import { theme } from "./theme"` — same name, in curly braces, because
   it's a *named* export. `SqlScriptGenerator.jsx` instead does
   `export default function SqlScriptGenerator() {...}` — a *default*
   export — so it's imported without curly braces:
   `import SqlScriptGenerator from "./SqlScriptGenerator"`.

2. **Relative paths (`./`, `../`) are literal folder navigation.**
   `./theme` means "theme.js in the same folder as me."
   `../theme` (used inside `components/`) means "go up one folder, then
   find theme.js." If you move a file to a different folder, every
   relative import inside it (and every file that imports *it*) needs
   updating — this is the #1 cause of "module not found" errors when
   restructuring a project.

## Running it locally

This is now a complete Vite + React project — `package.json`,
`vite.config.js`, `index.html`, and `src/main.jsx` were added so it can be
installed and run on its own, no extra setup needed.

```bash
npm install
npm run dev
```

`npm run dev` starts a local dev server (Vite prints the URL, usually
`http://localhost:5173`) with hot-reload — save any file and the browser
updates instantly.

To produce the optimized production build (what actually gets deployed):

```bash
npm run build
```

This writes static files into a `dist/` folder — plain HTML/CSS/JS that
any static host (Netlify, Vercel, GitHub Pages, S3...) can serve. Sanity-
check it locally before deploying:

```bash
npm run preview
```

## Deploying to Netlify

Two ways to do this. Both end with the same live site.

### Option A — drag and drop (fastest, no Git needed)

1. Run `npm run build` locally. This creates a `dist/` folder.
2. Go to [app.netlify.com/drop](https://app.netlify.com/drop).
3. Drag the `dist/` folder onto the page.
4. Netlify uploads it and gives you a live URL in seconds.

Downside: if you change the code, you have to rebuild and re-drag the
folder every time. Fine for a one-off demo, tedious for ongoing work.

### Option B — connect a Git repo (recommended)

Every `git push` then auto-rebuilds and redeploys — no manual steps.

1. Push this project to a GitHub repo:
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<repo-name>.git
   git push -u origin main
   ```
   The `.gitignore` here already excludes `node_modules/` and `dist/`, so
   the repo stays small — Netlify installs dependencies and builds fresh
   on its own servers.
2. Go to [app.netlify.com](https://app.netlify.com) → **Add new site** →
   **Import an existing project** → pick your GitHub repo.
3. Netlify auto-detects the build settings from `netlify.toml`:
   - Build command: `npm run build`
   - Publish directory: `dist`
4. Click **Deploy**. You get a live URL like
   `https://your-site-name.netlify.app`, which updates automatically on
   every future push.

### What `netlify.toml` does

It's what makes step 3 automatic instead of manual — Netlify reads it and
configures the build itself. It also adds a redirect rule
(`/* -> /index.html`), standard practice for any single-page React app,
so refreshing a non-root URL won't 404 if you add client-side routing
later.

## Where to make common changes

| You want to...                                   | Edit this file                          |
|----------------------------------------------------|------------------------------------------|
| Change any color, including the header blue         | `theme.js`                                |
| Change how the SQL text is formatted (e.g. quoting) | `utils/sqlHelpers.js`                     |
| Change how the column-picker dropdown looks/behaves | `components/MultiSelectDropdown.jsx`      |
| Change the "view uploaded data" popup               | `components/DataPreviewModal.jsx`         |
| Change page layout, add new state, wire new features| `SqlScriptGenerator.jsx`                  |

That table alone is the real payoff of splitting the file: instead of
"where in 650 lines is this?", it's "which file owns this concern?"
