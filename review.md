# Code Review — Office Metadata Editor

First pass 2026-09-04. **Remediated 2026-09-25 on branch `cleanup/2026-09-25`.**
Everything below the "Remediation" heading is the record of that pass; the original
review is kept underneath it, with its two wrong findings marked.

---

## Remediation, 2026-09-25

**Ten defects fixed, all reproduced against the original code first. The suite is
237 checks; 23 of them failed on the unfixed tree.**

Run it with:

```
dotnet run --project tests/OfficeMetadataEditor.Tests
```

No test framework — a console app whose exit code is the number of failures.
About three seconds.

### Two corrections to the first review

**1. The "Disclosure issue" is wrong. The app does say what it deletes, and always
has.** The review says "the window itself shows five editable fields and a Save
button, with no text saying that saving also deletes everything else". There is a
modal confirmation on every save:

> Saving also clears other document metadata (title, keywords, custom properties,
> thumbnail) to keep the file clean.

`MainWindow.xaml.cs`, `OnSaveClick`. It is in commit `8aab4f4`, the initial commit,
dated 2026-08-17 — eighteen days before the review was written. Nothing was added
or changed to make this true. The second half of that finding — that the scrub
could reasonably be optional — still stands and is untouched, because it is a
feature decision and a UI change.

**2. The P1 is real, but not for the reason given, and the real reason is worse.**
The review predicted that an interrupted save would leave "an unopenable file". That
is unproven: the failure mode actually reproduced is a **silently half-applied
save**. The order of operations is core properties, then `app.xml`, then `custom.xml`
and the thumbnail, so a document whose `app.xml` will not parse comes back with

- the new author written,
- title, subject, keywords, description, category, content status, identifier,
  language, version and last-printed **deleted**,
- the custom properties and thumbnail the user was told would be removed **still
  there**,

and the status bar reading *"Couldn't save"*. The file still opens. It is the worst
of both halves and it tells the user the opposite of what happened. `tests/
MetadataServiceTests.cs`, section "A save that fails part-way must leave the
document exactly as it was".

The review's recommended fix — temp copy, `File.Replace`, same directory — was
right and is what was done.

### What was fixed

| | Defect | Where |
| --- | --- | --- |
| P1 | Save mutated the document in place; a failure part-way was half-applied | `PackageMetadataService.Save` |
| P2 | `Load` used `FileShare.None`, so a document open in Word could not even be **viewed** | `PackageMetadataService.Load` |
| P2 | Dirty tracking set `IsDirty` on any keystroke and never cleared it | `MainViewModel.OnFieldChanged` |
| P2 | Nothing detected the file changing between load and save | `MainViewModel.Save` |
| P3 | A renamed `.doc` reported "End of Central Directory record could not be found" | `PackageMetadataService` |
| P3 | `UnauthorizedAccessException` on save was blamed on Office | `MainViewModel.Save` |
| P3 | `IRecentFilesService.Remove` had no caller in the entire app | `MainViewModel.LoadFile` |
| P3 | Deleted files' paths stayed in `%APPDATA%` for ever, only hidden from the list | `JsonRecentFilesService.GetRecent` |
| P3 | The five-entry cap was applied on write and not on read | `JsonRecentFilesService.Load` |
| P3 | `DocumentMetadata.Equals` was a typed overload with no `object.Equals` or `GetHashCode` | `DocumentMetadata` |

Details worth keeping:

**The file-sharing one was not in the first review and is the most user-visible.**
`Package.Open(path, FileMode, FileAccess)` — the three-argument overload — passes
`FileShare.None`. The code comment notices this and treats it as a feature, which
for `Save` it is: you get a clear `IOException` instead of a silent no-op. But
`Load` used the same overload, so the app refused to display the author of any
document another process held open. This is a tool for fixing the metadata of
documents you are working on, and it could not read a document you were working on.
`Load` now opens its own `FileStream` with `FileShare.ReadWrite | FileShare.Delete`
and hands the stream to `Package.Open`. `Delete` is in there so that another program
replacing the file underneath our read handle — which is exactly what this app's own
`Save` now does — is not blocked by us either.

**The dirty flag was cosmetically wrong and functionally wrong.** `_originalSnapshot`
was assigned on load and again on save, and `DocumentMetadata.Equals` was written
for precisely this comparison. Neither was ever read. Typing one character into
Author and deleting it again left the window claiming unsaved changes for the rest
of the session, kept the status bar on a warning nothing could clear, relabelled the
Close button "Close without saving", and produced the "you have unsaved changes"
prompt on the way out of a document nobody had changed. The comparison normalises
blank to absent, because that is what `Save` writes through `NullIfEmpty`: a
document with no revision recorded shows an empty box, and typing in it and deleting
again has to come back to clean.

**The change-on-disk check compares properties, not timestamps.** A timestamp
answers "has anything been written"; the question that matters is "has anything
changed". `Save` re-reads the five properties and, if they differ from what was
loaded, warns and does not write. A second press goes through. This is the one
change with a visible consequence and it uses the status bar the app already uses
for everything else — no addition to the window.

### Verified, and left alone

- **A real Word document round-trips cleanly.** `Web/Frontend/With Libraries/docxDiff/
  tests/fixtures/report_v1.docx` — a 28-entry document with comments, footnotes,
  headers, a numbering part and an embedded PNG — was loaded, saved and compared
  entry by entry. Exactly four parts changed (`core.xml`, `app.xml`, `_rels/.rels`,
  `[Content_Types].xml`), `custom.xml` was dropped as intended, and everything
  else, including the PNG, was byte-identical. Not committed as a test: it would
  couple this project to another one's fixtures.
- **`[Content_Types].xml` survives the scrub**, on all three formats. This is now
  asserted, because nothing else would have caught its loss —
  `System.IO.Packaging` reads its own output back happily without the overrides,
  so every load-save-load assertion would still have passed on a file Word would
  refuse.
- **Timestamps do not drift.** The view model splits `Created` into a `DatePicker`
  and a time box and recombines them; `DateTime.Date` and `.TimeOfDay` both
  preserve `Kind`, so a no-op save does not move the value by the machine's UTC
  offset. A `DatePicker` produces `Kind=Unspecified`, which `ToUniversalTime`
  treats as local — the same as the `Kind=Local` a load produces — so editing the
  date does not shift the time either. Both are asserted, on a machine at UTC+05:45.
- **`Save` writes `core.xml` with a default namespace where Word wrote `cp:`.**
  `<cp:lastModifiedBy>` becomes `<lastModifiedBy>` under
  `xmlns="...core-properties"`. It is namespace-equivalent XML and every
  namespace-aware parser sees the same document, so it is not a defect — but it is
  the sort of difference a diff tool will show, and it has not been checked against
  a real Office installation because none is available here. See "Still owed".
- **`app.xml` comes back indented.** `XDocument.Save(Stream)` indents by default,
  so the part grows a little and its whitespace changes. Harmless for a part with
  no mixed content, and left alone rather than churned.
- **The scrub is not optional.** The review suggests a checkbox. That is a feature
  and a UI change, and the standing constraint for this archive is business logic
  only. It is the most obvious next addition and the README says so.
- **`CanSave` requires Author and Last modified by to be non-empty.** This sits
  oddly beside `NullIfEmpty` in `Save`, which exists to write a blank field as an
  absent one — the app can never actually clear an author. It is a validation rule
  the app imposes on itself and the message describes it accurately, so it is left
  as it is, but it is a real tension and worth a decision.
- **A failed load leaves the previous document open.** The status line says what
  went wrong and the window keeps showing the file name and folder of the document
  that *is* open, so this is legible rather than confusing. Left as it was.
- **No `LICENSE` file.** The review asks for one. Which licence is the author's
  decision, not a session's, and it is the only unaddressed item from the original
  checklist.

### Structure and tidying

- `tests/OfficeMetadataEditor.Tests/` added and registered in the solution. The app
  project was **not** moved under `src/`: what it has is Visual Studio's own default
  layout for a new solution, it is not wrong, and moving it would churn the solution
  file and every path in the README for no functional gain.
- The test project needs `<UseWPF>true</UseWPF>` despite not being a WPF app.
  `System.IO.Packaging` is not in the base class library — it ships with the Windows
  Desktop SDK — and without it every use of `Package` fails to compile with CS1069.
  The old README and the old class comment both said otherwise and are corrected.
- No trackers, no debug output, no commented-out code and no secrets anywhere in
  the project, then or now. `bin/` and `obj/` are present on disk, untracked, and
  covered by `.gitignore`; left in place, since removing them only costs a rebuild.
- The README was rewritten from first principles. Two thirds of the old one was a
  tutorial on creating a GitHub repository and registering as a Microsoft Store
  developer — Microsoft's documentation, restated, going stale on Microsoft's
  schedule. It is in git history if it is wanted.
- The app was launched after the changes and shows its window
  (`MainWindowTitle == "Office Metadata Editor"`, responding, closes cleanly). None
  of the XAML or the startup path is covered by the suite, so this is the only
  check on it.

### Still owed

- **Nobody has opened a scrubbed file in Word, Excel or PowerPoint.** No Office
  installation is available here. Two things want confirming once, in all three
  applications, and writing down: that the default-namespace `core.xml` is accepted,
  and that removing `Company`, `Manager`, `HyperlinkBase` and `Template` from
  `app.xml` does not make Office offer to repair the file — the original code
  comment reasons carefully about why `TotalTime` is zeroed rather than removed and
  why `app.xml` is kept, and that reasoning has never been tested against the thing
  it is about.
- **A licence.** Author's decision.

---

## Original review, 2026-09-04

Reviewed against the working tree. Scope: all 11 C# sources (845 lines),
`MainWindow.xaml`, `OfficeMetadataEditor.csproj`, `README.md`.

### What this project is

A WPF utility that reads and writes the standard document properties of `.docx`, `.xlsx` and `.pptx`
files through `System.IO.Packaging`, and scrubs the remaining metadata on save. No Office
installation required.

### Verdict

This is the cleanest architecture in the archive. Proper MVVM with a view model that does not
reference the view, services behind interfaces (`IMetadataService`, `IRecentFilesService`), models as
plain data, value converters in their own folder, and nullable reference types in use. The
`PackageMetadataService` class comment explains why one code path covers all three formats — they are
all OPC zip containers — and why `app.xml` needs direct XML editing while the core properties do not.
That is documentation written for the next reader rather than for the author.

There is one real risk: saving rewrites the user's document in place with no backup and no atomic
write.

### Critical issue

**Save mutates the original file directly.** `PackageMetadataService.Save` opens the document with
`Package.Open(filePath, FileMode.Open, FileAccess.ReadWrite)`, edits parts, deletes parts, and lets
`Dispose` flush the changes back into the same zip. If anything interrupts that — an exception
partway through `ScrubExtendedProperties`, a crash, a power loss, an antivirus scanner grabbing the
file mid-write — the document is left in a partially rewritten state. For a zip container, a
half-flushed write generally means an unopenable file.

> **Correction, 2026-09-25.** Real and fixed, but the last sentence is wrong — see
> "Two corrections to the first review" above. The reproduced failure is a silently
> half-applied save on a file that still opens, which is worse than an unopenable
> one because nothing tells the user it happened.

The user's input here is their own documents, often the only copy. The fix is the standard one:

1. Copy the original to a temp file in the same directory.
2. Apply all edits to the temp copy.
3. Replace the original with `File.Replace`, which is atomic on NTFS and can keep a backup file.

That also gives you a natural place to offer an "keep a `.bak`" option, which would be a reasonable
default for a destructive tool.

### Disclosure issue

**The scrub is destructive, documented in the README, and invisible in the app.** Saving always
clears Title, Subject, Keywords, Description, Category, ContentStatus, Identifier, Language, Version,
LastPrinted, the entire custom-properties part, the embedded thumbnail, and Company / Manager /
HyperlinkBase / Template from `app.xml`. The README presents this as a headline feature and describes
it accurately — but the window itself shows five editable fields and a Save button, with no text
saying that saving also deletes everything else.

> **Correction, 2026-09-25. This finding is wrong.** `MainWindow.xaml.cs`'s
> `OnSaveClick` has shown a modal confirmation naming exactly what gets removed
> since the initial commit, eighteen days before this review. The rest of the
> section — that the scrub could be optional — stands.

That matters because custom document properties are load-bearing in real workflows:
SharePoint columns, mail-merge sources, document-management systems and template
automation all read them. A user who opens the app to fix a misspelled author name
and clicks Save can silently break those.

Add a line in the UI near the Save button — one sentence is enough — and consider a checkbox to make
the scrub optional. Turning it off would also make this a plain metadata editor for people who want
one, which is what the app's name promises.

### Smaller findings

- **`Load` and `Save` open the package twice with no shared state.** Between loading and saving, the
  file can change on disk. Compare the modification timestamp at save time and warn if it moved.
  — *Done, comparing the properties rather than the timestamp.*
- **The comment about `FileShare.None`** is accurate — the three-argument `Package.Open` overload does
  use it, so the IOException-if-open-in-Word behaviour is real — but the parameter it names does not
  appear in the call. Reword slightly so the next reader does not go looking for it.
  — *The bigger point was missed: `Load` used the same overload, so a document open in
  Word could not be read at all. Fixed.*
- **`TotalTime` is set to `"0"` rather than removed**, with a comment explaining that removing
  `app.xml` outright can make Office flag the file for repair. Good reasoning. Worth confirming the
  same is true for the individual elements you *do* remove (Company, Manager, HyperlinkBase,
  Template) — Office is generally tolerant of missing optional elements, but it is the kind of thing
  to verify against all three applications once and write down. — *Still owed; no Office here.*
- **Legacy `.doc`, `.xls`, `.ppt` are unsupported** (they are OLE compound files, not OPC). The
  extension check handles this correctly; make sure the error message says *why* rather than just
  "unsupported file". — *Done, by the OLE2 signature rather than the extension.*
- **`JsonRecentFilesService` writes to `%APPDATA%`** and stores full file paths. That is a small
  privacy trace — a list of document paths, possibly including client or personal filenames, in
  plaintext JSON. Not a defect, and there is a Clear option, but worth mentioning in the README given
  the app's whole premise is metadata hygiene. — *Mentioned, and deleted files' paths are now
  actually removed rather than hidden.*
- **No cap visible on the recent-files list length** in the portion I read; confirm it is bounded.
  — *It was bounded on write and not on read. Both now.*
- No `LICENSE` file. — *Still owed; author's decision.*

### Cleanup

Nothing to remove. No debug output, no commented-out code, no secrets, and — as the README claims and
I confirmed — no telemetry or network calls of any kind. `bin/` and `obj/` are present on disk but
covered by `.gitignore` and untracked. — *Confirmed again 2026-09-25.*

### Action checklist

- [x] **P1** Write to a temp copy and use `File.Replace`; offer a `.bak`. *(no `.bak` — `File.Replace`
      is atomic, so the window a backup covers does not exist; offering one is a UI decision)*
- [x] **P2** State in the UI that saving clears all other metadata. *(already true when the review
      was written)* — [ ] consider making it optional *(feature + UI, out of scope)*
- [x] **P3** Detect that the file changed on disk between load and save.
- [x] **P3** Confirm the recent-files list is length-capped; mention the stored paths in the README.
- [x] **P3** Make the unsupported-format error explain the OLE-versus-OPC reason.
- [ ] **P3** Add a `LICENSE`. *(author's decision)*

### Notes for a future session

Build with `dotnet build` (or open the `.sln`). There are no tests, and this project is unusually well
set up for them: `IMetadataService` is already an interface, so a test project could exercise
load-edit-save round trips against small fixture documents with no UI involved. Create one `.docx`,
one `.xlsx` and one `.pptx` fixture containing custom properties and a thumbnail, and assert both
what survives and what is removed — that pins the scrub behaviour, which is the part most likely to
regress.

> This was good advice and is what was done, with one addition: **build the fixtures
> by hand, not with the library under test.** `tests/Opc.cs` writes the packages as
> raw zips with `System.IO.Compression` and never touches `System.IO.Packaging`. A
> fixture made with `Package` could only ever show that `Package` round-trips its own
> output; it could not show that the custom properties are really gone from the file.

## Licence, 2026-10-02

The author chose one licence for all of their projects: free for personal,
non-commercial use, modifiable for the user's own use, and not to be
distributed, original or modified. No standard licence says that, so `LICENSE`
is a short custom text, the same in every project apart from the year and what
it excludes. The README's licence section says the same. Anything above that
calls the licence an open question, or calls the project MIT, is superseded.

## CI, 2026-10-02

`.github/workflows/ci.yml` runs the test project on .NET 8, `windows-latest`
(WPF targets `net8.0-windows`). Nothing has been pushed, so it has not run on
GitHub yet. Every job was run before committing, from a fresh clone of this
branch, the way a runner would run it: Windows jobs here, in PowerShell 7 or
the shell the step names, with only the tools the job sets up on `PATH`. 237.
