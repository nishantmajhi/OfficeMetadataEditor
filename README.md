# Office Metadata Editor

A Windows desktop app for reading and rewriting the document properties of
`.docx`, `.xlsx` and `.pptx` files. It edits five fields — Author, Last
modified by, Revision number, Created and Last modified — and, on save,
deletes everything else the file was carrying about you.

Office is not required. Nothing is sent anywhere.

## What it does to a file

Opening a document shows five values. Saving writes those five back and
removes the rest:

| Where | What goes |
| --- | --- |
| `docProps/core.xml` | Title, subject, keywords, description, category, content status, identifier, language, version, last-printed date |
| `docProps/app.xml` | Company, manager, hyperlink base, template path; total editing time reset to zero |
| `docProps/custom.xml` | Removed entirely — every custom document property |
| Thumbnail | The embedded preview image, removed |

That is roughly the same ground as Word's own **File → Info → Check for
Issues → Inspect Document → Remove Properties and Personal Information**,
minus anything inside the document body. Comments, tracked changes and
revision history are *not* touched: those live in format-specific parts and
removing them safely is a different job from this one.

**The deletion is not optional and there is no undo.** Custom document
properties in particular are load-bearing in real workflows — SharePoint
columns, mail-merge sources, document-management systems and template
automation all read them — so a document that depends on them will stop
working. The app asks for confirmation before every save and names what it
is about to remove. If you only want to fix a misspelled author name on a
document like that, copy it first.

## Why one editor handles all three formats

`.docx`, `.xlsx` and `.pptx` are the same thing wearing three hats: an
[OPC](https://en.wikipedia.org/wiki/Open_Packaging_Conventions) package, which
is a zip file with a manifest. Rename one to `.zip` and look inside — the
document properties are in `docProps/core.xml`, in the same schema, whichever
application wrote it.

So there is no Word path, Excel path and PowerPoint path in this codebase.
There is one, in `Services/PackageMetadataService.cs`, built on
`System.IO.Packaging.Package`, which reads and writes that part directly. The
file type is used for one thing only: which letter and which colour to put in
the badge.

What this does *not* cover is the pre-2007 binary formats. A `.doc`, `.xls` or
`.ppt` is an OLE2 compound file, a completely different container, and the app
says so rather than letting the zip reader complain about a missing central
directory.

## Saving is all-or-nothing

A save is applied to a copy of the document in the same folder and swapped in
with `File.Replace`, which is atomic on one volume and preserves the original
file's attributes, permissions and creation time.

This is not incidental. Writing metadata means rewriting the zip container, and
doing that in place means that any interruption — a malformed part, a crash,
an antivirus scanner, a power cut — leaves a document with half a save applied
to it. The input to this app is somebody's own documents, often the only copy.

Two consequences worth knowing:

- Saving needs room for a second copy of the document, briefly, and needs
  permission to create a file in the document's folder. A folder where you may
  change existing files but not create new ones will refuse the save.
- If the document is open in Word, Excel or PowerPoint, the save fails
  immediately and says so, before any copy is made. Reading is unaffected —
  you can open and view the properties of a document that is open elsewhere.

## Other behaviour worth knowing

**It notices if the file changed while you had it open.** The five values on
screen were read when you opened the document. If something else has written to
it since, the first press of Save warns you instead of overwriting, and a
second press goes ahead.

**Timestamps are shown in your local time zone and stored as UTC**, which is
what the format requires. A document created at 09:15 UTC shows as 09:15 in
London and 15:00 in Kathmandu, and saving it back without touching it does not
move it.

**The recent-files list is kept in
`%APPDATA%\OfficeMetadataEditor\recent.json`** — five full paths, in plain
text. Those paths may themselves say something about you, which is worth
mentioning in an app about metadata hygiene. "Clear recently opened" empties
it, and entries whose file has been deleted are dropped from the file, not just
hidden from the menu.

**Revision on `.xlsx`**: Excel does not always show `cp:revision` in its own
properties dialog even though it is a valid core property. The value is
written correctly; Excel just may not show it back to you.

## Building and running

Needs the [.NET 8 SDK](https://dotnet.microsoft.com/download) and Windows.
WPF does not run on Linux or macOS.

```
dotnet build
dotnet run --project OfficeMetadataEditor
```

Or open `OfficeMetadataEditor.sln` in Visual Studio 2022 and press F5.

`WindowBackdropType="Mica"` needs Windows 11. On Windows 10 the window loses
the backdrop effect and is otherwise fine.

## Tests

```
dotnet run --project tests/OfficeMetadataEditor.Tests
```

237 checks, no test framework — a console app that prints a line per assertion
and exits with the number of failures. There is nothing to install.

`tests/Opc.cs` builds `.docx`, `.xlsx` and `.pptx` fixtures as raw zips using
`System.IO.Compression` and reads them back the same way, deliberately never
touching `System.IO.Packaging`. A fixture built with the library under test
could only ever show that the library round-trips its own output; building them
independently is what makes "the custom properties really are gone" a claim
about the file rather than about the API.

The fixtures carry a title, subject, keywords, description, category, content
status, version, language and identifier, a Company / Manager / HyperlinkBase /
Template in `app.xml`, a custom-properties part, a thumbnail, and an embedded
binary part that is not valid UTF-8 — so "the document body survived" is a
claim about bytes.

## How it is put together

```
OfficeMetadataEditor.sln
├── OfficeMetadataEditor/               the app
│   ├── App.xaml(.cs)                   composition root: builds the services by hand
│   ├── MainWindow.xaml(.cs)            the only window
│   ├── Models/
│   │   ├── DocumentMetadata.cs         the five fields, plus equality and normalisation
│   │   └── OfficeFileType.cs           Word/Excel/PowerPoint: badge letter and accent colour
│   ├── Services/
│   │   ├── IMetadataService.cs         load / save / is-this-supported
│   │   ├── PackageMetadataService.cs   the whole of the OPC work
│   │   ├── IRecentFilesService.cs
│   │   └── JsonRecentFilesService.cs   the %APPDATA% list
│   ├── ViewModels/
│   │   └── MainViewModel.cs            state, commands, dirty tracking, validation
│   └── Converters/                     file type → badge/brush, bool → visibility
└── tests/OfficeMetadataEditor.Tests/   console app, no dependencies
```

The view model has no reference to the window, and the services are behind
interfaces — which is why the tests can drive the whole app with two fakes and
no UI. There is no DI container; `App.OnStartup` builds three objects by hand,
which is the right amount of machinery for three objects.

## Dependencies

| Package | Why |
| --- | --- |
| [WPF-UI](https://github.com/lepoco/wpfui) | Fluent/Windows 11 controls, Mica backdrop, and `SystemThemeWatcher` for live light/dark switching. MIT. |
| [CommunityToolkit.Mvvm](https://github.com/CommunityToolkit/dotnet) | Microsoft's MVVM source generators — `[ObservableProperty]`, `[RelayCommand]` — instead of hand-written `INotifyPropertyChanged`. MIT. |

Everything else — file dialogs, JSON, the package reading and writing — is the
.NET base class library. `System.IO.Packaging` is worth a note: it is *not* in
the base class library proper. It ships with the Windows Desktop SDK, so a
plain `net8.0` project referencing it fails to compile until it is added as a
NuGet package. This app targets `net8.0-windows` with `UseWPF`, so it is
already there.

## Limits

- Five fields, not every property. Anything not in that list is deleted on
  save, not edited.
- No way to keep the custom properties or the thumbnail. Turning the scrub off
  would make this a plain metadata editor as well as a cleaner, and is the most
  obvious thing to add next.
- Document-body metadata — comments, tracked changes, embedded personal
  information inside the content — is out of scope.
- If you need to *read or write* custom properties rather than delete them, or
  need word and slide counts, swap `System.IO.Packaging` for the
  [Open XML SDK](https://github.com/dotnet/Open-XML-SDK). For what this app
  does, `Package` is smaller and adds no dependency.

## Licence

Free for personal, non-commercial use. You may use, study and modify this
software for yourself, but you may not use it commercially, and you may not
share or publish it, in its original form or modified. See [LICENSE](LICENSE)
for the full terms.
