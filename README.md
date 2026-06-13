# ImageRight 7.1 — Email / Excel Duplicate-Detection Automation

> **Deliverable:** Approaches and implementation design for an automation that detects duplicate
> inbound emails (and their Excel attachments) arriving into ImageRight, and **flags/annotates**
> them in place rather than reprocessing them.
> **Target system:** ImageRight 7.1, Revision 8 (Vertafore).
> **Status:** Design / approaches document (no code written yet).

---

## 1. Context

**The problem.** Clients submit business by email. Each email typically carries an Excel attachment
(e.g. a bordereau, schedule, or loss run). Those emails are captured into ImageRight (Vertafore's
insurance Enterprise Content Management / document-management system). The same submission is
frequently sent more than once — resent threads (`RE:` / `FW:`), accidental re-sends, or a periodic
file delivered again. Today nothing automatically notices that a newly-arrived item is the *same* as
one already in ImageRight, so duplicates get re-handled, re-indexed, and waste underwriting/claims time.

**The goal.** When an email arrives into ImageRight, automatically determine whether the **same email
has already been seen before** and, if so, **flag/annotate it as a suspected duplicate while leaving
it in place** (no auto-delete, no silent discard). A human still makes the final call.

**Duplicate is defined by (confirmed with stakeholder):**
1. **Exact attachment file** — byte-for-byte identical attachment (file hash).
2. **Normalized email subject** — subject with `RE:`/`FW:`, whitespace and case stripped.
3. **Sender + subject + date composite** — same sender, same normalized subject, same send date/period.
   *(Optional later enhancement: hashing the Excel **data/content** so re-saved/renamed files with
   identical rows are still caught — see §6.4.)*

**Decision on a match:** flag/annotate and keep (set a metadata flag + note, optionally raise a
low-priority task), do **not** remove the document.

---

## 2. Is this automatable? — Short answer

**Yes.** ImageRight is not a black box; Vertafore exposes several supported automation surfaces, and the
duplicate-check logic itself is straightforward. The only real branch point is *how we connect* to
ImageRight, which depends on whether we get **API access** (the clean path) or must **drive the desktop
UI** (the fallback). This document covers both, recommends the API-first path, and specifies the
fallback so the project is not blocked while API access is confirmed.

---

## 3. What kind of automation does ImageRight support?

ImageRight is a Windows-based ECM with a server, a desktop client, a web client, and a workflow engine.
For programmatic automation it offers, roughly from most to least "official":

| # | Surface | What it is | Good for | Notes |
|---|---------|-----------|----------|-------|
| 1 | **ImageRight REST API** | Modern HTTP/JSON API ("500+ APIs"), JWT auth, resources for **Files, Folders, Documents, Pages, Tasks/Diaries, Search**, plus **SignalR push notifications** for real-time events. Distributed via the **Vertafore Developer Portal** (Sandbox + Live), with a downloadable **C# SDK (V2)**. | Headless, event-driven integrations. **The right tool for this job.** | Confirm the exact API version/endpoints licensed for **7.1 rev 8** (public docs reference 6.8; 7.x ships the same family). Needs Developer Portal registration + credentials. |
| 2 | **Workflow / Workflow Studio** | ImageRight's task-based workflow engine. Define **steps**, **criteria**, and **automatic/manual routing**. Items (tasks) flow through queues. | Making the dup-check a native step every incoming item passes through, and routing/flagging the result. | Configuration-driven; custom logic typically lives in an API/SDK-backed step or external service the step calls. |
| 3 | **Intelligent / Advanced Capture** | Capture + OCR + auto-index. Ingests from **email and network folders**, batch tools, templates, auto-index to eliminate manual matching. | Extracting subject/sender and Excel cell values at ingestion; applying dedup at capture time. | Licensed module — confirm availability. Strongest if you later want **Excel-content** dedup. |
| 4 | **Outlook Add-In / Email Import** | Desktop import of emails + attachments. Configurable body format (**EML / Native / PDF**) and file-type conversion (`.xlsx`, `.doc`, …). Index fields captured on import (see §4). | Understanding *how* items enter ImageRight and what metadata is available to match on. | This is the **ingestion path**, not an automation API by itself. |
| 5 | **ImageRight.NET / DMU integration** | Configuration-based integration (e.g. AIM/DMU → ImageRight) for auto file creation and drawer/file-type mapping. | Auto-filing from other Vertafore products. | **Not** a general-purpose programmable API — config-driven. Mentioned for completeness. |
| 6 | **RPA / UI automation** (3rd-party) | Drive the **Windows desktop client** with UiPath / Power Automate Desktop / Blue Prism, or scripting (pywinauto, AutoIt, FlaUI). | **Fallback when no API access.** | Brittle (UI changes break it), needs a logged-in session, hardest to scale — but unblocks delivery. |

**Key takeaway:** items 1 + 2 (REST API + Workflow) are the supported, robust combination. The REST
API's **SignalR push** is what lets us react "whenever a mail arrives" without constant polling. Item 6
is the safety net if API access cannot be obtained.

---

## 4. Data we can match on (the "fingerprint" inputs)

From the email itself and from ImageRight's index/metadata captured at import:

- **From the email:** sender, subject, body, send date/time, attachment file name(s), attachment bytes.
- **From ImageRight index fields (captured on import):** Drawer, File Type, File Number, Document Type,
  Folder, Document Description, **Document Date**, **Received Date/Time**, Page descriptions, plus
  org-specific custom attributes (state, policy type, agent name, etc.).

These give us everything needed to build the duplicate signatures in §6 without OCR. (OCR/Capture is
only needed if we later want to dedup on the Excel's *internal data*.)

---

## 5. Approaches (with implementation outline)

Five approaches, ordered by recommendation. The recommended solution (§7) is a **hybrid of A + B**, with
**D** as the documented fallback.

### Approach A — API-first integration service *(Recommended primary)*

A small **headless service** (Windows Service / .NET worker) that:

1. **Subscribes** to ImageRight events via **SignalR** (new document/email in the target drawer/queue),
   or polls a workflow queue via REST if SignalR isn't available for 7.1 rev 8.
2. **Pulls** the new item's metadata (subject/sender/dates) and the Excel attachment bytes via the REST API.
3. **Computes** the three duplicate signatures (§6).
4. **Looks up** prior occurrences in a **dedup registry** (§6.3) — a small dedicated table — and/or
   queries the ImageRight **Search API** by metadata.
5. **On match:** sets a metadata flag + adds a note/annotation, optionally raises a low-priority task
   (§8). **On new:** records the signatures and lets the item continue normally.

- **Pros:** supported, headless (no UI session), reliable, testable, scales, easy to log/audit.
- **Cons:** needs API access + Developer Portal credentials + sandbox.
- **Tech:** **C# / .NET 8** worker service (native fit for ImageRight + official C# SDK). Excel reading
  via **ClosedXML / EPPlus / OpenXML SDK**. Dedup store in **SQL Server** (ImageRight already runs on SQL).

### Approach B — Workflow custom step *(Recommended companion)*

Add a **"Duplicate Check" step** in **Workflow Studio** that every inbound item passes through. The step
invokes the same dedup logic (calling the service from Approach A or an API endpoint) and uses workflow
**criteria/routing** to set the flag and keep the item in its normal queue.

- **Pros:** native to ImageRight's process model; visible, governable, auditable; no separate trigger needed.
- **Cons:** logic still lives in API/SDK code behind the step; needs workflow admin rights.
- **Use it as:** the *trigger + routing layer* that delegates the actual comparison to Approach A.

### Approach C — Intelligent / Advanced Capture auto-index

Use the Capture module to extract subject/sender (and, if desired, Excel cell values) at ingestion and
apply dedup during auto-indexing.

- **Pros:** best path if you later want **Excel-content** dedup; processing happens at the front door.
- **Cons:** requires the licensed Capture module; more config; heavier.
- **Use it for:** the optional Excel-data dedup enhancement (§6.4), not the first release.

### Approach D — RPA / UI automation *(Documented fallback — only if no API access)*

Drive the ImageRight **desktop client** with **UiPath** or **Power Automate Desktop** (both strong on
Windows desktop automation; PAD is low/no-cost with Windows). The bot watches the inbox/queue, reads
the on-screen subject/sender, exports/opens the attachment, computes signatures, checks the registry,
and uses the UI to set the flag/note.

- **Pros:** works with **zero API access**; unblocks delivery.
- **Cons:** brittle (breaks on UI/version changes), needs an always-logged-in interactive session,
  slower, harder to audit, doesn't scale well.
- **Tech:** **UiPath** (enterprise-grade, ImageRight automations exist in the wild) or **Power Automate
  Desktop** (lower cost). Excel via the same .NET/Python libraries. **Pair with E** to dedup before
  the desktop step where possible.

### Approach E — Email-side pre-filter *(Optional complement)*

Catch duplicates **before** ImageRight using **Microsoft Graph / Exchange** on the monitored mailbox:
compute the signatures at the mailbox and tag/move/skip obvious resends.

- **Pros:** stops duplicates at the source; reduces load on ImageRight; works alongside A or D.
- **Cons:** only sees the email side, not what's already filed in ImageRight; needs mailbox API access.
- **Use it as:** a cheap first line of defense, especially valuable in the RPA fallback scenario.

---

## 6. Duplicate-detection logic (the core algorithm)

Independent of the connection approach. Compute **three signatures** per incoming email; a match on
**any** flags it as a suspected duplicate (tunable to "any" vs "all").

### 6.1 Signatures

| Signature | How it's computed | Catches |
|-----------|-------------------|---------|
| **`attachment_hash`** | `SHA-256` over each attachment's raw bytes | Identical file resent (most precise) |
| **`subject_norm_hash`** | `SHA-256` of the normalized subject: lowercase, trim, collapse whitespace, strip leading `RE:`/`FW:`/`FWD:` (repeatedly), strip surrounding punctuation | Resent threads / replies with same subject |
| **`composite_key`** | `SHA-256` of `normalize(sender) \| subject_norm \| date_bucket` where `date_bucket` = send date (or period, e.g. ISO week/month for periodic files) | Periodic re-submissions from same sender |

> Normalization rules should be documented and unit-tested — they are where most false positives/negatives come from.

### 6.2 Decision

```
incoming email
   ├─ compute attachment_hash[], subject_norm_hash, composite_key
   ├─ query dedup registry (and/or ImageRight Search) for any match
   │     match on attachment_hash      → DUPLICATE (high confidence)
   │     match on composite_key        → DUPLICATE (high confidence)
   │     match on subject_norm_hash    → SUSPECTED  (lower confidence)
   ├─ if DUPLICATE/SUSPECTED → flag + annotate (keep in place), link to original DocID
   └─ else (new) → record all signatures in registry, continue normal processing
```

### 6.3 Dedup registry (recommended)

A dedicated SQL table is the simplest reliable "memory." Querying the ImageRight Search API alone is
possible but slower and depends on the right index fields being populated.

```sql
CREATE TABLE ir_dedup_registry (
    id                BIGINT IDENTITY PRIMARY KEY,
    imageright_doc_id NVARCHAR(64)  NOT NULL,   -- original document/file id in ImageRight
    file_number       NVARCHAR(64)  NULL,
    sender            NVARCHAR(320) NULL,
    subject_raw       NVARCHAR(998) NULL,
    subject_norm_hash CHAR(64)      NULL,
    composite_key     CHAR(64)      NULL,
    attachment_hash   CHAR(64)      NULL,       -- one row per attachment
    attachment_name   NVARCHAR(260) NULL,
    received_at       DATETIME2     NOT NULL,
    created_at        DATETIME2     NOT NULL DEFAULT SYSUTCDATETIME()
);
CREATE INDEX ix_attach   ON ir_dedup_registry(attachment_hash);
CREATE INDEX ix_subject  ON ir_dedup_registry(subject_norm_hash);
CREATE INDEX ix_composite ON ir_dedup_registry(composite_key);
```

### 6.4 Optional enhancement — Excel **content** dedup

You originally mentioned matching on *Excel content*. Not in scope for v1 (you selected file/subject/sender
signals), but easy to add: open the workbook (ClosedXML/EPPlus/openpyxl), read each sheet's used range,
normalize (trim, type-coerce, ignore formatting/empty trailing rows), serialize deterministically, and
`SHA-256` it into a `content_hash` column. This catches re-saved/renamed files with identical data. Pairs
naturally with **Approach C** (Capture) if you want it done at ingestion.

---

## 7. Recommended architecture

**Primary:** Approach **A (API service)** as the brain, triggered and governed by Approach **B (workflow
step)**, backed by the **dedup registry** (§6.3). **Fallback:** Approach **D (RPA)** if API access is
denied, optionally fronted by **E (email pre-filter)**.

```
                 ┌──────────────────────────────────────────────┐
   inbound email │  ImageRight 7.1  (Drawer / Workflow queue)     │
  ──────────────▶│                                                │
                 │  [Workflow step: "Duplicate Check"] ── calls ──┼──┐
                 └──────────────────────────────────────────────┘  │
                          ▲  flag/annotate (keep in place)          │ REST + SignalR
                          │                                         ▼
                 ┌────────┴───────────────┐        ┌───────────────────────────────┐
                 │  Dedup Service (.NET)   │◀──────▶│  Dedup Registry (SQL Server)   │
                 │  • subscribe events     │        └───────────────────────────────┘
                 │  • pull email + xlsx    │
                 │  • compute signatures   │        (Fallback if no API:
                 │  • lookup + decide      │         UiPath/PAD bot drives desktop UI,
                 │  • write flag + note    │         + optional MS Graph pre-filter)
                 └─────────────────────────┘
```

**Why this is the recommendation**
- **.NET/C#** is the native ecosystem for ImageRight (official C# SDK; runs alongside the SQL backend).
- **SignalR events** give true "on arrival" reaction without polling load.
- **Workflow step** keeps the behavior visible and governable inside ImageRight, not hidden in a script.
- **Dedicated registry** makes lookups fast, the logic testable, and the audit trail clean.
- The **RPA fallback** means the project isn't blocked while API access is being arranged.

---

## 8. "Flag/annotate, keep" — what the action looks like

On a match, **do not move or delete** the document. Instead:

1. **Set a metadata flag** — a custom index field, e.g. `DuplicateStatus = "Suspected Duplicate"` and
   `DuplicateOf = <original DocID / File Number>`.
2. **Add a note / annotation** on the document (e.g. a sticky note on page 1): *"Possible duplicate of
   document #<id> received <date> from <sender> — flagged by automation on <timestamp>."*
3. **(Optional) Raise a low-priority Task/Diary** to the relevant queue so a human reviews and confirms.
4. **Log** the decision (incoming id, matched original id, which signature matched, confidence) for audit.

The item continues through its normal workflow — the flag is informational, the human decides.

---

## 9. Approach comparison (decision matrix)

| Criterion | A: API service | B: Workflow step | C: Capture | D: RPA/UI | E: Email pre-filter |
|-----------|:---:|:---:|:---:|:---:|:---:|
| Needs API access | ✅ yes | ✅ yes | partial | ❌ no | mailbox API |
| Reliability | High | High | High | Low–Med | Med |
| Reacts on arrival | ✅ (SignalR) | ✅ (queue) | ✅ | polling | ✅ |
| Excel-content dedup | add-on | add-on | ✅ native | add-on | limited |
| Effort to build | Med | Low (on top of A) | Med–High | Med–High | Low–Med |
| Maintenance risk | Low | Low | Med | **High** | Low |
| Recommendation | **Primary** | **Companion** | Later/optional | Fallback only | Optional complement |

---

## 10. Open questions / dependencies to confirm (with Vertafore & internal IT)

1. **API availability for 7.1 rev 8** — is the REST API + Developer Portal access licensed/enabled? Does
   this revision support **SignalR** push, or must we poll a workflow queue?
2. **Ingestion path** — do emails enter via server-side import/Capture, a monitored mailbox, or users
   manually importing via the Outlook Add-In? (Determines the trigger.)
3. **Where do they land** — which Drawer / File Type / Workflow queue should the automation watch?
4. **Credentials & environments** — service account, Developer Portal Sandbox, and a non-prod ImageRight
   instance to test against.
5. **Custom index fields** — can we add `DuplicateStatus` / `DuplicateOf` fields (IEMC/admin rights)?
6. **Volume** — emails/day and attachment sizes (sizing the service + registry).
7. **"Match" policy** — flag on **any** signature match (higher recall) vs **all** (higher precision)?
8. **Capture module** — is Intelligent/Advanced Capture licensed (for the optional Excel-content dedup)?
9. **RPA tooling** — if we go the fallback route, is UiPath available, or should we use Power Automate Desktop?

---

## 11. Risks & mitigations

| Risk | Mitigation |
|------|-----------|
| API not available on 7.1 rev 8 | Fallback to Approach D (RPA); front with Approach E to reduce load |
| False positives flag legitimate new mail | "Flag, don't delete" + human review; tune match policy; start in **log-only/shadow mode** |
| Subject normalization misses edge cases | Unit-test normalization with real subject samples; keep rules configurable |
| UI changes break RPA (fallback) | Use stable selectors, pin client version, monitor with alerts |
| Registry drifts from ImageRight reality | Backfill registry from historical docs; reconcile periodically via Search API |
| Re-saved Excel evades file-hash | Add content-hash (§6.4) in phase 2 |

---

## 12. Phased roadmap

- **Phase 0 — Discovery (1–2 wks):** confirm §10 items; obtain sandbox + credentials; collect ~200 real
  sample emails (incl. known duplicates) for tuning.
- **Phase 1 — Logic core (1 wk):** implement signatures + normalization + registry as a standalone,
  fully unit-tested library (no ImageRight dependency yet).
- **Phase 2 — Integration (2–3 wks):** wire Approach A (REST/SignalR pull + flag/annotate) and the
  Approach B workflow step; run in **shadow/log-only mode** (decides but doesn't flag).
- **Phase 3 — Pilot (1–2 wks):** enable flagging on one drawer/queue; measure precision/recall; tune.
- **Phase 4 — Rollout & ops:** widen scope, add dashboards/alerts, document runbook.
- *(Fallback track if no API):* Phases 1 + 3 reuse as-is; swap Phase 2 for the RPA bot (Approach D + E).
- *(Optional later):* Excel-content dedup (§6.4) / Capture integration (Approach C).

---

## 13. Verification & testing plan

1. **Unit tests (logic core):** feed crafted emails — identical attachment, `RE:`/`FW:` variants, same
   sender+subject different day, genuinely-new email — assert the right signature fires and confidence
   level. Cover normalization edge cases explicitly.
2. **Registry tests:** insert/lookup correctness; index performance at expected volume.
3. **Sandbox integration test:** in the ImageRight Sandbox, import a known email, then re-import it;
   confirm the service detects the duplicate, sets `DuplicateStatus`/`DuplicateOf`, adds the note, and
   **leaves the document in place**. Confirm a brand-new email is *not* flagged.
4. **Shadow mode in pilot:** run live for a period writing decisions only to the log; compare against
   manual review to measure false positives/negatives before enabling real flagging.
5. **(Fallback) RPA dry-run:** run the bot against a test queue with screen-recording; verify it reads
   the right fields and applies the flag without disrupting normal handling.
6. **Acceptance:** agreed precision/recall threshold on the sample set; sign-off from the underwriting/
   claims team that flag + note + (optional) task is actionable.

---

## 14. Sources

- [ImageRight — Modern Insurance Document Management (Vertafore)](https://www.vertafore.com/products/insurance-document-management-system/imageright)
- [ImageRight — Insurance document management (ECM)](https://www.vertafore.com/products/enterprise-content-management/imageright)
- [ImageRight REST API documentation (6.8 PDF — API family reference)](http://support2.vertafore.com/Repository/Documentation/DWN311/ImageRight6.8RestAPI.pdf)
- [Getting Started with the Vertafore API Developer Portal](https://help.vertafore.com/devportal/content/getstarted/gettingstarted.htm)
- [Developer Portal — Download SDK (C# V2)](https://help.vertafore.com/devportal/content/howto/downloadsdkbutton.htm)
- [ImageRight Microsoft Outlook Add-In (email import)](https://help.vertafore.com/imagerightdesktop/content/microsoft_outlook_add-in.htm)
- [ImageRight Import Dialog Box Explained (index/metadata fields)](https://help.vertafore.com/imagerightdesktop/content/import_dialog_box_explained.htm)
- [ImageRight Import Dialog Options (convert types / attachments)](https://help.vertafore.com/imagerightdesktop/content/outlookinterface/import_dialog_options.htm)
- [ImageRight Flow and Step Display Options (Workflow)](https://help.vertafore.com/imagerightdesktop/content/flow_and_step_display_options.htm)
- [ImageRight Manual Route (Workflow routing)](https://help.vertafore.com/imagerightdesktop/content/outlookinterface/manual_route.htm)
- [ImageRight.NET Setup (DMU integration)](https://help.vertafore.com/datamaintenanceutility/content/imageright.net_setup.htm)
- [ImageRight 7.1 Client Feature Comparison (PDF)](https://online.vertafore.com/rs/920-PGU-682/images/External%20-%20ImageRight%207.1%20Client%20Feature%20Comparison.pdf)
- [Community: Vertafore ImageRight API — Create Operation](https://developer.sailpoint.com/discuss/t/vertafore-imageright-api-create-operation/188083)
- [GitHub: manikumarkv/imageright-apis (ImageRight 7.2 API wrapper)](https://github.com/manikumarkv/imageright-apis)
