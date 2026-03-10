# Google Docs → dotCMS Sync

A Google Apps Script sidebar add-on that pushes Google Doc content to dotCMS as structured content.


<img width="2524" height="1372" alt="Screenshot 2026-03-10 at 10 06 12 AM" src="https://github.com/user-attachments/assets/710d9033-bbb7-4df7-9bb2-c31370bae0f8" />






## How It Works

1. **Configure** your dotCMS host URL and API token in the Settings tab
2. **Select a content type** — the plugin auto-generates a two-column metadata table (`Field | Value`) in your Google Doc with the required fields
3. **Fill in field values** — use the sidebar's Field Editor for dropdowns, booleans, and relationship lookups, or type directly in the table
4. **Write your content** below the metadata table as normal Google Doc content
5. **Click Sync** — the plugin pushes everything to dotCMS

## Features

- **Arbitrary content types** — works with any dotCMS content type; search and select with a chip-style picker
- **Metadata table** — two-column table with a `Field | Value` marker header row, auto-populated with required fields from the Content Type API
- **Field Editor** — sidebar controls for select/dropdown fields, boolean fields (true/false), and relationship field lookups (single and multi-select via Lucene queries)
- **Body field selector** — choose which content type field receives the document body; persisted as `dotBodyField` in the metadata table
- **Code formatting** — monospace text in Google Docs is automatically wrapped in `<code>` (inline) or `<pre><code>` (block) tags during HTML export
- **Clean HTML export** — body content exported as HTML with Google's inline styles and unnecessary markup stripped
- **Image handling** — inline images uploaded via dotCMS Temp API → dotAsset flow, with MD5 hash-based deduplication across syncs
- **Image/file metadata fields** — inline images in image or file field cells are uploaded as dotAssets; the identifier is written back below the image so subsequent syncs skip re-uploading
- **Embedded drawings/charts** — exported as PNG and uploaded as dotAssets
- **Relationship fields** — relationship values are sent as Lucene query strings (`+identifier:(id1 OR id2)`) for proper dotCMS persistence
- **Smart host detection** — automatically detects custom `HostFolderField` variables on the content type; falls back to `host` only when no custom field exists
- **Create or update** — first sync creates content, subsequent syncs update it (identifier and inode tracked in the metadata table)
- **Site/folder picker** — select target site and folder from the sidebar; auto-selects from metadata on load
- **Language picker** — push content to any language version
- **Draft/Publish toggle** — uses `/api/v1/workflow/actions/default/fire/EDIT` or `/PUBLISH`
- **Persistent sync log** — shows last sync timestamp, status, and any failed items
- **Progress indicator** — progress bar during sync
- **Partial failure handling** — if some images fail, content still syncs and failures are reported

## Project Structure

```
├── appsscript.json     Manifest with OAuth scopes
├── Code.gs             Entry point, menu, server-side function routing
├── Settings.gs         Per-user host URL + API token (UserProperties)
├── DotCMSApi.gs        dotCMS REST API wrapper
├── DocParser.gs        Metadata table detection, field extraction, body HTML export, image collection
├── HtmlCleaner.gs      Strips Google inline styles from exported HTML
├── ImageHandler.gs     Image hash tracking, dedup, Temp API → dotAsset upload
├── SyncEngine.gs       Sync orchestration (create/update, partial failure handling)
├── Sidebar.html        Two-tab sidebar UI (Sync + Settings)
├── .claspignore        Files excluded from clasp push
└── .clasp.json         Local clasp config (git-ignored)
```

## Setup

### Prerequisites

- [Node.js](https://nodejs.org/)
- [clasp](https://github.com/google/clasp) — `npm install -g @google/clasp`
- A dotCMS instance with API access

### Deploy for Testing

```bash
# Login to Google
clasp login

# Create an Apps Script project bound to a Google Doc
clasp create --type docs --title "dotCMS Sync"

# Push files
clasp push

# Open the script editor
clasp open-script
```

Then open the bound Google Doc, refresh, and go to **Extensions → dotCMS Sync → Open Sidebar**.

### Domain-Wide Deployment

1. In the Apps Script editor: **Deploy → New deployment** (type: Editor Add-on)
2. In the [Google Cloud Console](https://console.cloud.google.com), enable the **Google Workspace Marketplace SDK** for your GCP project
3. Go to the Marketplace SDK **App Configuration** tab and enter the deployment ID
4. Under **Store Listing**, fill in the app name, description, and required assets (icon, screenshot)
5. Set visibility to **Private** (or Public with Internal OAuth consent screen) and publish

## dotCMS API Endpoints Used

| Endpoint | Purpose |
|----------|---------|
| `GET /api/v1/contenttype` | List content types |
| `GET /api/v1/contenttype/{var}/fields` | Get content type fields |
| `GET /api/v1/site` | List sites |
| `GET /api/v1/folder/siteid/{id}` | List folders |
| `GET /api/v2/languages` | List languages |
| `GET /api/content/_search` | Search content (relationship lookups) |
| `POST /api/v1/temp` | Upload temp file (images) |
| `POST /api/v1/workflow/actions/default/fire/SAVE` | Save as draft |
| `POST /api/v1/workflow/actions/default/fire/PUBLISH` | Publish content |

## Google Doc Structure

```
┌─────────────────────────────┐
│  Field     │  Value         │  ← Marker header row (required)
├─────────────────────────────┤
│  title     │  My Article    │  ← Required fields auto-generated
│  author    │  af884e69...   │  ← Relationship field (identifier)
│  image     │  🖼️ + id       │  ← Image field (inline image + identifier)
│  category  │  Tech          │  ← Optional fields added via dropdown
│  dotBodyField │ blogContent │  ← Body field mapping (auto-persisted)
└─────────────────────────────┘

Body content goes here — including `inline code` and:

    code blocks in monospace fonts

These become <code> and <pre><code> in the exported HTML.
Images are automatically uploaded as dotAssets.
```
