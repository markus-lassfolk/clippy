# Word and PowerPoint: editing and workflows (m365-agent-cli)

**Purpose:** Describe how to **edit**, **review**, and **automate** Word (`.docx`) and PowerPoint (`.pptx`) files using this CLI and Microsoft Graph, including **checkout**, **versions**, **convert**, collaboration links, and **OOXML** round-trips.

**Related:** [`GRAPH_PRODUCT_PARITY_MATRIX.md`](./GRAPH_PRODUCT_PARITY_MATRIX.md) (workload status), [`GRAPH_API_GAPS.md`](./GRAPH_API_GAPS.md) (platform limits), [`AGENT_WORKFLOWS.md`](./AGENT_WORKFLOWS.md) §7 (short agent checklist), [`CLI_REFERENCE.md`](./CLI_REFERENCE.md) (flags and read-only mode), [`GRAPH_SCOPES.md`](./GRAPH_SCOPES.md) / [`GRAPH_PERMISSION_FEATURE_MATRIX.md`](./GRAPH_PERMISSION_FEATURE_MATRIX.md) (permissions), [`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md) (raw **`graph invoke`** when needed).

---

## 1. What Graph covers (and what it does not)

| Area | Support | CLI |
| --- | --- | --- |
| **Drive item lifecycle** | Upload, download, copy, move, delete, share, permissions, checkout/checkin, version list/restore | **`files`** or **`word`** / **`powerpoint`** (mirrored per-item verbs) |
| **Preview & thumbnails** | Embeddable preview session; thumbnail URLs | **`word preview`**, **`powerpoint preview`**, **`thumbnails`** |
| **Format conversion** | e.g. PDF via Graph convert job | **`convert`** (see subcommand help for `--format` and output path) |
| **Compliance & library** | List item metadata, MIP sensitivity, retention labels, follow, permanent delete | **`list-item`**, **`sensitivity-*`**, **`retention-label*`**, **`follow`**, **`permanent-delete`** |
| **Discovery on the drive** | List, search, delta | **`files`** only (not duplicated under **`word`** / **`powerpoint`**) |
| **In-document Word comments** (threaded, on the file content) | **Not** a first-class Graph drive-item API comparable to Excel **`workbook/comments`** | **`graph invoke`** only if Microsoft documents a path for your scenario; often **desktop Office** or **OOXML** |
| **PowerPoint slide / shape object model** | **No** Excel-style **`…/workbook/…`** REST for decks | Same as above: **OOXML**, **desktop Office**, or documented **beta** via **`graph invoke`** |

The CLI prepares **URLs**, **bytes**, and **file state**; **Office Online** and **desktop Office** perform actual editing and **live co-authoring**. There is no separate “start co-authoring” API in this tool beyond sharing, checkout discipline, and uploading new revisions.

---

## 2. `word` / `powerpoint` vs `files`

- **`word`** and **`powerpoint`** call the **same Graph endpoints** as **`files`** for every **mirrored** subcommand (upload, checkout, convert, …). Use them when the intent is “this drive item is a Word doc” or “this is a deck”; use **`files`** for **folder-level** work (**`list`**, **`search`**, **`delta`**, **`shared-with-me`**, …).
- **Drive location** is identical: pick **at most one** of **`--user`**, **`--drive-id`**, **`--site-id`**, or **`--site-id`** + **`--library-drive-id`** (default **`/me/drive`**). See [`AGENT_WORKFLOWS.md`](./AGENT_WORKFLOWS.md) §3.

---

## 3. Human or browser editing (Office Online)

1. Get a link: from **`word meta`** / **`powerpoint meta`** use **`webUrl`**, or create a preview with **`word preview`** / **`powerpoint preview`** (optional **`--json-file`** for body shapes like `chromeless`).
2. Optional **exclusive** workflow: **`word checkout`** / **`powerpoint checkout`** before asking someone to edit, then **`checkin`** with **`--comment`** when done (pair with your org’s SharePoint check-out policy).
3. Optional **sharing**: **`share`**, **`invite`**, **`permissions`** (same patterns as **`files`**; **`--collab`** where supported — see **`--help`**).

---

## 4. Checkout, checkin, and co-authoring expectations

- **Checkout** requests an **exclusive** edit lock on the item where the service supports it; **checkin** releases it and can carry a comment. Not every library enforces check-out; behavior is **tenant/library** dependent.
- **Co-authoring** in Office Online does not require a special Graph “mode” from this CLI beyond normal sharing and file access. If you need a **strict** “only one editor” period, prefer **checkout** + communicate, then **checkin** after upload or after browser edits conclude.

---

## 5. Versions and restore

- **`versions <fileId>`** — list stored versions (JSON with **`--json`** where supported).
- **`restore <fileId> <versionId>`** — restore a prior version per Graph semantics.

Use these for **audit** and **rollback** after automated uploads or user mistakes.

---

## 6. Convert

- **`convert`** — produces another format (commonly **PDF**); use subcommand **`--help`** for **`--format`**, output path, and drive flags.

Handy when agents must deliver a **read-only** artifact without editing Office XML.

---

## 7. Agent automation: download → edit → upload

1. **`word download`** / **`powerpoint download`** — save `.docx` / `.pptx` bytes locally.
2. Edit with any tool that outputs a **valid** Office Open XML package (see §9).
3. **`word upload`** / **`powerpoint upload`** or **`upload-large`** for big files.

**Idempotence:** Prefer targeting a **folder** + filename or explicit replace semantics per **`upload`** help; avoid blind overwrite without reading **`meta`** first if collisions matter.

**Legacy binary formats** (e.g. `.doc`, `.ppt`) are not modern OOXML; convert them in Office or another converter **before** a Graph-centric pipeline expects `.docx`/`.pptx`.

---

## 8. Compliance, library columns, and signals

- **`list-item`** — SharePoint **library columns** and metadata (**often 404 on personal OneDrive**).
- **`sensitivity-assign`** / **`sensitivity-extract`** — Microsoft Purview / MIP; body via **`--json-file`** per Graph docs; tenant/licensing dependent.
- **`retention-label`** / **`retention-label-remove`** — item retention labeling.
- **`follow`** / **`unfollow`** — OneDrive for Business “followed” item.
- **`permanent-delete`** — destructive; bypasses recycle bin where policy allows.
- **`activities`** / **`analytics`** — item activity feed and usage-style analytics where enabled (see **`--help`**).

All of the above exist on **`files`** under the **same** subcommand names.

---

## 9. OOXML and local editing (practical pointers)

Office **`.docx`** and **`.pptx`** files are **ZIP** archives containing **XML** parts and **relationships** (ECMA-376 / ISO Open XML).

| Format | Main content types | Typical approaches |
| --- | --- | --- |
| **Word** | WordprocessingML (`word/document.xml`, styles, headers, …) | **Microsoft Word**, **LibreOffice**, **python-docx** (limited schema), **Open XML SDK** (.NET), direct ZIP+XML with care for **`[Content_Types].xml`** and **`.rels`** |
| **PowerPoint** | PresentationML (slides under `ppt/slides/`, layouts, theme) | **PowerPoint**, **LibreOffice Impress**, **Open XML SDK**, programmatic ZIP/XML (higher complexity than Word for shapes/media) |

**Guidance for agents**

- Prefer **high-level libraries** that maintain valid packages; hand-editing XML is error-prone.
- After local edits, **re-upload** in one transaction when possible; use **`upload-large`** for big decks with media.
- If validation fails in Office after upload, compare **before/after** package structure (missing `rels` or content types is the usual culprit).

---

## 10. Discovery and read-only safety

- Find items: **`graph-search`**, **`files list`** / **`search`**, **`files recent`**, or **`files delta --state-file`** for durable sync ([`AGENT_WORKFLOWS.md`](./AGENT_WORKFLOWS.md) §4).
- **`--read-only`** blocks documented mutating commands; confirm [CLI_REFERENCE.md](./CLI_REFERENCE.md) § Read-Only Mode. Read/query helpers such as **`meta`**, **`download`**, **`versions`**, **`convert`** (depending on implementation) may or may not be gated—**check help** before automation.

---

## 11. When you need `graph invoke`

Use **`m365-agent-cli graph invoke`** (or **`graph batch`**) for:

- Undocumented or **tenant-specific** Graph paths.
- **Beta** endpoints for Word/PowerPoint that are not wrapped yet (see [`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md)).

Always verify **method**, **URL**, **body**, and **required scopes** against current Microsoft Graph documentation.

---

*Last updated: 2026-05-05 — Initial guide aligned with `word` / `powerpoint` / `files` parity and Graph platform limits.*
