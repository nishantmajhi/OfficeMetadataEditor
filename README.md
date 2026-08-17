# Office Metadata Editor

A small, focused WPF app for viewing and editing the standard document
properties (author, last modified by, revision number, created/modified
timestamps) of `.docx`, `.xlsx`, and `.pptx` files - without needing Word,
Excel, or PowerPoint installed.

It follows your Windows light/dark setting automatically and matches
Office's own brand colors per file type.

## Why this exists

There isn't a good metadata editor for Office files in the Microsoft Store -
the few that exist are dated, non-English-only, or have rough UI/UX. This
fills that gap with a small, native, Fluent-styled utility.

## Features

- Edit **Author**, **Last modified by**, **Revision number**, **Created**,
  and **Last modified** for `.docx`, `.xlsx`, and `.pptx` files
- **Cleans the file on save**: clears every other piece of metadata Office
  quietly keeps around - title, subject, keywords, description, category,
  content status, custom document properties, the embedded thumbnail
  preview, and identifying fields in `app.xml` like company, manager, and
  hyperlink base - so what you get back is just the file plus the five
  fields you chose to keep
- Automatic light/dark theme, synced live with Windows' setting
- Recent files list (with a "Clear recently opened" option)
- Clear, specific error messages (e.g. when a file is still open in Office)
- No telemetry, no accounts, no internet connection required

## Why one editor works for all three formats

`.docx`, `.xlsx`, and `.pptx` are all **OPC packages** - plain zip files that
share a `docProps/core.xml` part for author, last-modified-by, revision,
created, and modified. `System.IO.Packaging.Package` (built into .NET, no
extra NuGet package needed) reads and writes that part directly through
`Package.PackageProperties`. That's the whole trick - see
`Services/PackageMetadataService.cs`.

Saving does more than edit those five fields, though - it's a real cleanup
pass. It also clears the rest of `docProps/core.xml` (title, subject,
keywords, etc.), scrubs identifying fields out of `docProps/app.xml`
(company, manager, hyperlink base, template path), removes the
`docProps/custom.xml` part entirely, and strips any embedded thumbnail
preview. That mirrors what Word/Excel/PowerPoint's own "Inspect Document ->
Remove Properties and Personal Information" does - minus touching document
_content_ (comments, tracked changes), which is format-specific and out of
scope for this app.

If you later want to touch custom properties instead of just deleting them,
or need word/slide counts, swap this out for the DocumentFormat.OpenXml SDK
(Microsoft, actively maintained) - but for what this app does,
`System.IO.Packaging` is simpler and adds zero dependencies.

## Tech stack

| Package                                                                       | Why                                                                                                                                                                                                                |
| ----------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| [WPF-UI](https://www.nuget.org/packages/WPF-UI)                               | Fluent/Windows 11-style controls, Mica backdrop, and `SystemThemeWatcher.Watch(this)` for automatic light/dark theme syncing. Actively maintained ([lepoco/wpfui](https://github.com/lepoco/wpfui)), MIT licensed. |
| [CommunityToolkit.Mvvm](https://www.nuget.org/packages/CommunityToolkit.Mvvm) | Microsoft's own MVVM toolkit - `[ObservableProperty]` / `[RelayCommand]` source generators instead of hand-written `INotifyPropertyChanged` boilerplate.                                                           |

Everything else (file dialogs, JSON persistence, metadata read/write) uses
only the .NET base class library - no other dependencies.

## Project layout

```
OfficeMetadataEditor/
├── OfficeMetadataEditor.sln
├── README.md
├── .gitignore
└── OfficeMetadataEditor/
    ├── OfficeMetadataEditor.csproj
    ├── App.xaml                        # Composition root - builds services + MainViewModel
    ├── App.xaml.cs
    ├── MainWindow.xaml                 # Titlebar, menu, empty/loaded states, status bar
    ├── MainWindow.xaml.cs
    ├── Models/
    │   ├── DocumentMetadata.cs         # Creator / LastModifiedBy / Revision / Created / Modified
    │   └── OfficeFileType.cs           # Word/Excel/PowerPoint classification, badge, accent color
    ├── Services/
    │   ├── IMetadataService.cs
    │   ├── PackageMetadataService.cs   # Reads/writes docProps/core.xml, scrubs everything else
    │   ├── IRecentFilesService.cs
    │   └── JsonRecentFilesService.cs   # Recent-files list persisted to %AppData%
    ├── ViewModels/
    │   └── MainViewModel.cs            # All state + commands, dirty tracking, validation
    └── Converters/                     # File-type -> badge/brush, bool -> visibility, status -> color
        ├── OfficeFileTypeConverters.cs
        └── BooleanToVisibilityConverters.cs

```

## Prerequisites

- Windows 10/11 (WPF doesn't run on Linux/macOS)
- [.NET 8 SDK](https://dotnet.microsoft.com/download)
- Visual Studio 2022 (optional, but recommended for XAML editing/debugging)

## Building and running

```
git clone https://github.com/<your-username>/OfficeMetadataEditor.git
cd OfficeMetadataEditor
dotnet restore
dotnet build
dotnet run --project OfficeMetadataEditor
```

Or open `OfficeMetadataEditor.sln` in Visual Studio 2022 and press F5.

## Known limitations / possible next steps

- **Saving is destructive by design**: every save clears title, subject,
  keywords, custom properties, and the thumbnail - there's no "keep these
  too" option yet. If you need any of that preserved, make a copy of the
  file first.
- **Content-level metadata isn't touched**: comments, tracked changes, and
  revision history inside the document body aren't part of this app's
  cleanup - only the package-level properties described above are.
- **File locking**: `Save()` opens the package with `FileAccess.ReadWrite`
  and no sharing, so it fails fast with a clear status-bar message if the
  file is still open in Word/Excel/PowerPoint, rather than silently no-oping.
- **Revision on .xlsx**: Excel doesn't always surface `cp:revision` in its
  own UI even though it's a valid core-properties field, so edits here may
  not be visible inside Excel itself - the value is still written correctly
  to `docProps/core.xml`.
- No automated tests yet - `PackageMetadataService` is the one piece worth
  covering (round-trip a save against a real sample file of each type and
  confirm the scrubbed fields are actually gone).

## Publishing this project to GitHub

1. Create a new (empty) repository on GitHub - don't initialize it with a
   README, license, or `.gitignore`, since you already have this project
   locally.
2. Add a `.gitignore` for Visual Studio/.NET before your first commit (a
   ready-made one is included in this project as `.gitignore`; if you don't
   have it, GitHub's own [`VisualStudio.gitignore`
   template](https://github.com/github/gitignore/blob/main/VisualStudio.gitignore)
   works too). This keeps `bin/`, `obj/`, and user-specific files out of
   version control.
3. From the solution folder:
   ```
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/<your-username>/OfficeMetadataEditor.git
   git push -u origin main
   ```
4. Consider adding a `LICENSE` file (e.g. MIT, matching the libraries this
   project depends on) so others know how they can use your code.

## Publishing to the Microsoft Store (optional, for later)

You don't need to do any of this to build and run the app locally - this
section is only relevant if you decide you want it listed in the Store.

### 1. Register as a developer

Go to [Microsoft Store Developer
Center](https://partner.microsoft.com/dashboard) and sign up. As of 2026,
registration is free for both individual developers and companies (Microsoft
dropped the old $19/$99 registration fees). You'll verify your identity with
a Microsoft account (personal) or a Microsoft Entra ID (work/organization)
account.

### 2. Reserve your app's name

In Partner Center, create a new app submission and reserve a name (e.g.
"Office Metadata Editor"). This also gives you the Store identity values
(Package/Identity `Name` and `Publisher`) you'll need for packaging.

### 3. Choose a packaging path

You have two options - pick whichever fits how much control you want over
the build:

**Option A - Packaged (MSIX), the traditional route**

1. In Visual Studio, right-click the solution -> **Add** -> **New
   Project...** -> search for "Packaging" -> choose **Windows Application
   Packaging Project**.
2. In the new packaging project, right-click **Applications** -> **Add
   Reference...** -> select `OfficeMetadataEditor` as the entry point.
3. Open the packaging project's `Package.appxmanifest` and set the
   `Identity` `Name`/`Publisher` to match what Partner Center reserved for
   you, plus your app's display name, description, and logo assets (a
   300x300 Store logo is the minimum requirement).
4. Right-click the packaging project -> **Publish** -> **Create App
   Packages...** -> choose **Microsoft Store using a new app name** (or an
   existing reservation) -> follow the wizard. This produces an
   `.msixupload` file.
5. In Partner Center, upload that `.msixupload` under your submission's
   Packages section. Microsoft signs and certifies it for you - you don't
   need to buy your own code-signing certificate for this path.

**Option B - Unpackaged (installer link), the simpler route**

Since 2021, the Store also accepts traditional Win32 installers (no MSIX
required):

1. Build a signed installer (e.g. with
   [Inno Setup](https://jrsoftware.org/isinfo.php),
   [WiX](https://wixtoolset.org/), or `dotnet publish` with
   `-p:PublishSingleFile=true --self-contained true`).
2. Get a code-signing certificate that chains to a CA in the [Microsoft
   Trusted Root Program](https://learn.microsoft.com/en-us/security/trusted-root/participants-list)
   - self-signed certificates aren't accepted for this path - and sign your
     installer with it.
3. Host the signed installer at a stable URL that won't change after
   submission (your own site, GitHub Releases, etc.).
4. In Partner Center, choose the unpackaged app submission flow and point it
   at that URL instead of uploading an `.msixupload`.

### 4. Fill out the Store listing and submit

Add screenshots, a description, age rating, and pricing (free is fine), then
submit for certification. Certification typically checks that the app
launches, behaves as described, and doesn't violate content/security
policies - for a small offline utility like this one, that's usually
straightforward. You'll get an email once it's approved (or with specific
feedback if it isn't).

### A note on Mica/WindowsBackdropType

This app uses `WindowBackdropType="Mica"`, which requires Windows 11. On
Windows 10 it falls back gracefully (no crash), just without the backdrop
effect - worth mentioning in your Store listing's system requirements if you
want to be precise about minimum OS version.
