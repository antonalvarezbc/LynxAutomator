# LynxAutomator user manual

[Español](manual.md) · [Português](manual.pt.md) · [Installation and packages](../README.md#linux-macos-and-windows-builds)

This revised manual describes the full application on `feature/cross-platform-builds`.
Older release executables may behave differently. The original
[DOCX guide](../WIP%20LynxAutomator%20GUIDE%20.docx) is retained as a historical
reference, including its screenshots; current instructions are maintained in Markdown.

## 1. Getting started and editions

LynxAutomator prepares spreadsheets, downloads authorized Wildlife Insights images
and provides camera-trap utilities. It does not automatically upload spreadsheets
or photos to Wildbook. Select Spanish, Portuguese or English in the interface;
some messages remain in English. Save results before changing language: doing so
rebuilds the forms. Language changes are blocked during any active task.

| Feature | Full | Historical alpha mini |
| --- | --- | --- |
| Wildbook folder import and catalog | Yes | Yes |
| WI download and CSV conversion | Yes | Yes, with fewer UI options |
| Iberian lynx spreadsheets | Yes | Yes |
| Video frame extraction | Yes | Yes |
| Dedicated original-file date correction | Yes | No |
| Bulk original-file renaming | Yes | No |

Mini still creates files, invokes gsutil and changes timestamps on extracted
frames. It is not read-only and is not signed simply because it is smaller. The
current CI workflow builds the full app only. See the [architecture review](architecture.md).

Follow the README to download and extract the package for your OS. Builds from
this branch still need successful CI runs on their target systems. Use copies of
original media when testing date correction or renaming.

### Long tasks, progress and cancellation

Video extraction, spreadsheet reads/processing/saves, catalogs, lynx data, date
scanning/correction, renaming and downloads run in the background in both editions.
The bottom panel shows the task, stage/current file and **Cancel**. An activity
indicator is used when a reliable completion percentage is unavailable.

Input fields and operation buttons are disabled during a task to keep parameters
stable and prevent conflicting file operations. The window continues handling
events. Wait or cancel before starting another task or changing language.

Cancellation is checked between frames, files and processing stages. An Excel
read/write, copy, decoder call or active transfer may need to finish first; this is
not an instant interruption. Each gsutil transfer has a five-minute timeout.
Completed files remain; cancelling does not undo original-file date changes or
renames. Frames, copies and Excel use temporary output; a workbook replaces its
destination only after a complete write and cancellation check.

Closing the window requests cancellation and waits for the worker to finish before
exiting. After original-file date changes, dates are rescanned even on cancellation
or failure to refresh the form's reference values.

## 2. Initial Wildbook spreadsheet

Use your project's `.xlsx` template with the exact column names required by your
Wildbook instance. Common fields include:

| Encounter.locationID | Encounter.country | Encounter.genus | Encounter.specificEpithet | Encounter.submitterID |
| --- | --- | --- | --- | --- |
| Andújar-Cardeña | Spain | Lynx | pardinus | your_username |

This is a partial example, not a complete import template. Include one initial row
with shared values. Modules may propagate values from that row. Leave calculated
photo/date/individual fields empty and review coordinates and other project fields.
Output is tabular data: original formatting, formulas and additional worksheets
are not necessarily preserved. Save to a new filename.

## 3. Wildbook → BIWbE from Folder

1. Select the folder directly containing PNG/JPG/JPEG images.
2. Select the initial spreadsheet.
3. Optionally enable multiple images per encounter and enter a nonnegative
   integer threshold in **seconds**.
4. Process, then download/save the resulting spreadsheet.

This module does not traverse subfolders. Only images with a readable EXIF
`DateTimeOriginal` are included; filesystem dates are not used as a fallback.
Check the output image count.

Grouping compares consecutive photos in time order. At a 60-second threshold,
10:00:00, 10:00:40 and 10:01:20 can all belong to one encounter. This module has no
project/camera separation; use a coherent folder per camera/deployment.

## 4. Wildbook → BIWbE Catalog

1. Select the catalog folder; subfolders are included.
2. Select the initial spreadsheet.
3. Choose whether to capitalize individual IDs and collapse rows sharing an ID.
4. Process and save.

The first space-delimited word of each filename becomes
`MarkedIndividual.individualID`: `Nube lateral.jpg` gives `Nube`, whereas
`Nube_lateral.jpg` gives `Nube_lateral`. Capitalization may change the remaining
letters of an ID too. Image references use paths relative to the selected folder.

If IDs exist only in folder names, use the renamer on copies first. Keep the space
separating the ID from the rest of the name; replacing it with an underscore
changes how this catalog module interprets it.

## 5. Wildlife Insights → WI Downloader

Request the desired filtered data export in Wildlife Insights, extract the package
and find `images.csv`. Follow the image-access instructions in the export's Data
Use and Citation Guide. See the official
[private download guide](https://www.wildlifeinsights.org/get-started/download/private).
Install Google Cloud CLI with `gsutil` on PATH and configure an authorized account.
`gsutil version` checks availability; installation alone does not grant bucket access.

1. Select a CSV with nonempty `location` and `deployment_id` columns.
2. Choose whether to use separate deployment folders.
3. Click Download and select the destination.
4. Review the downloaded, skipped and failed counts.

Locations must be `gs://` JPEG images (`.jpg`/`.jpeg`). Content is checked, unsupported
filename characters are replaced, and the extension becomes `.JPG`. Other formats
are not converted to JPEG. Existing files are skipped without revalidation.
Transfers use a temporary directory so incomplete downloads do not occupy the final
filename. Distinct locations mapping to the same output name during a run are
reported as conflicts.

Cancel in the bottom task panel takes effect after the current transfer, which has a five-minute timeout.
Wait for completion before starting again.

## 6. Wildlife Insights–Wildbook → WI CSVs to BIWbE

Select the initial `.xlsx` (first worksheet), images CSV and deployments CSV.
Both CSVs need `project_id` and `deployment_id`. After merging, data must include
`latitude`, `longitude`, `placename`, `location`, `timestamp`, `project_id`,
`deployment_id` and `subproject_name`. If `number_of_objects` is absent it defaults
to 1. Overlapping data-column names in the CSVs may cause an ambiguous/missing-column
error and need review.

Choose the options, process, resolve any reported errors and save:

- **Multiple images per encounter** groups consecutive images by a nonnegative
  integer threshold in seconds, separately for each project and deployment.
- **Separate images with more than one object** puts each image with
  `number_of_objects > 1` in its own row. It does not create one row per animal.

**Neither option is “multispecies”.** The code does not identify or compare species.
With both options enabled, images with more than one object remain separate while
the remaining images are grouped by time.

Empty identifiers, duplicate deployments, unmatched images and invalid timestamps
are rejected. Correct the inputs and retry. A failed run disables output from a
previous run.

Image references use basenames with `.JPG`. Match them against the actual upload
files, especially if downloads changed special characters or names repeat across
deployments. `Occurrence.occurrenceID` remains `project-deployment`, not a unique
burst ID; confirm this fits your project's import semantics.

## 7. Iberian Lynx module

Select the root **containing the farms**, then configure Settings to match:

| Revision folder | Linces folder | Path under selected root |
| --- | --- | --- |
| No | No | `Farm/Station/Individual/image.jpg` |
| Yes | No | `Farm/Station/Revision/Individual/image.jpg` |
| No | Yes | `Farm/Station/Linces/Individual/image.jpg` |
| Yes | Yes | `Farm/Station/Revision/Linces/Individual/image.jpg` |

Set grouping minutes to `0` to disable grouping or a positive integer to group.
Optional spreadsheets can enrich the data: stations require an `Estacion` column,
individuals a `Lince` column. Avoid duplicate keys, which can multiply rows.
Generate, review, then download the spreadsheet.

An individual folder such as `Nube y Brisa` creates records for both individuals.
Grouping combines files and individuals within farm/station/revision. File paths
are separated by `;`. `Número de Fotos` counts those entries, which may include
videos or repeated entries for multiple individuals; it is not a verified unique
photo count.

Capture dates come from EXIF when available. Files without a readable capture
date, commonly videos, can remain undated. Review these before temporal grouping.
Folder positions determine the fields, so extra nesting or incorrect checkboxes
can mislabel records. Verify farm, station, revision and individual in the output.

## 8. Date Changer

This applies the same clock offset to files directly in a folder, without recursion.

1. Select the folder.
2. Choose the oldest, newest or a custom reference date representing the camera's
   incorrect time.
3. Enter the corresponding real time as `YYYY-MM-DD HH:MM:SS`.
4. Choose Copy to Folder for corrected copies, or Rewrite Dates for originals.

If the camera showed `2024-05-01 10:00:00` when it was actually noon, all files
receive a +2-hour adjustment; they are not all assigned the same timestamp.

| Data | Updated behavior |
| --- | --- |
| JPEG with EXIF DateTimeOriginal | Capture, digitized and DateTime tags receive the shifted date |
| Other formats/JPEG without a capture tag | No embedded capture date is created or corrected |
| Windows filesystem timestamps | Creation, modification and access |
| Linux/macOS filesystem timestamps | Modification and access, not creation |

Oldest/newest uses readable EXIF, then creation time on Windows or modification
time on Unix. There is no automatic timezone conversion. Review the reference
when mixing formats or cameras; use a custom date when appropriate.

Copies take their reference from the source and use alternative filenames on
collisions. Per-file errors can leave a partially completed batch. Do not repeat
an adjustment on already corrected files without checking the new offset.

## 9. Video Frame Extractor

Select a folder directly containing MP4/AVI/MOV/MKV/FLV videos, enter a positive
finite interval in seconds and choose an output folder after clicking Extract.
Subfolders are not scanned. JPEG frames are sampled at whole-frame intervals,
with a minimum step of one frame. Output names include the video and frame number;
existing files are preserved using suffixes.

Before extraction, **ffprobe** (from FFmpeg, available on PATH) reads embedded
dates. Review the candidates and their source for every video; confirm or correct
them against the camera overlay. Missing/conflicting metadata or missing ffprobe
requires manual entry. Filesystem dates are never used automatically.

Enter `2024-07-15 14:30:00`, optionally with a UTC offset such as `+02:00`.
Without an offset, EXIF keeps camera clock time and filesystem dates are unchanged.
With an offset, EXIF offset tags and access/modification times are also set;
Windows sets creation time with millisecond precision, Linux/macOS do not.
Each frame adds its position (frame number/FPS). Capture, digitized, modification
and subsecond EXIF tags are written. Variable-frame-rate timing is approximate.
The camera overlay is not read automatically using OCR.

## 10. Image Renamer

Select a source folder (including subfolders), choose filename components, then
rename. Supported image extensions are JPG/JPEG/PNG/GIF. Enable Copy Photos and
choose a different destination to preserve originals; otherwise files are renamed
in their existing locations.

Options include the first word of the folder name, original filename, EXIF date,
custom text and replacing spaces with underscores. Folder/original-name options
apply capitalization. Custom text cannot contain `/` or `\`. If EXIF is missing, the date component is omitted.

Copying collects images into one destination instead of preserving the directory
tree. Name collisions receive suffixes. A destination nested under the source is
excluded from recursive processing.

## 11. Troubleshooting and limits

- Unknown-publisher/security warning: verify the package source and the computer's
  policy. Mini does not replace signing; do not disable protection just to try it.
- Missing Tkinter: use Python with Tk support as described in the README.
- Missing gsutil/access denied: check CLI availability, the launching process's PATH,
  authorized account and the instructions accompanying your WI export.
- Rejected CSV: check columns, identifiers, duplicate/unmatched deployments and dates.
- Missing photos in folder import: check readable capture EXIF and nonrecursive input.
- Cannot save Excel: close the workbook and choose a writable folder/new filename.
- Cancelling: wait for the current read/write, frame or transfer to finish.
- Disabled controls: wait for the active task or use Cancel in the bottom panel.

Save before closing or changing language. There is no universal undo history.
Generating a spreadsheet does not certify all Wildbook rules or dataset integrity.
See the [logic review](logic-review.md) for remaining improvements.

## Camtrap DP photographs

Open **Camtrap DP**, load a local `datapackage.json` or ZIP (1.x, CSV/CSV.gz), then
select species using the search box and checkboxes. No Wildbook registration is
required. Review the unique image counts before choosing the output directory.
Event-associated images are included by default from the same deployment/time
interval, including empty frames for detection in Wildbook. This does not imply
that all animals are the same individual. Local-only mode avoids network access.

A new `camtrap-…` directory holds validated images and `manifest.csv` linking IDs,
species, filenames and status. Private media and videos are skipped. Accessible
HTTP URLs are downloaded without additional credentials. Retry failed downloads
creates a new batch. Cancellation preserves completed files and cleans partials;
a pending HTTP read may take up to its 20-second timeout.

The optional local synthetic lynx fixture (excluded from Git) has 366 observations, 247 directly associated images,
300 including events, and 10 local JPEGs. Its images are not actual lynxes.
The reader performs basic structural checks, not full schema validation. Wildbook
Excel export and direct Agouti API access are not yet available in this tab.

## Shared Bulk Import: Wildlife Insights and Camtrap DP

For Wildlife Insights, click **Bulk Import from ZIP or CSV**. Choose a ZIP containing `images.csv` or `images_<project>.csv`, and exactly one `deployments.csv` (subfolders supported), or choose both CSVs separately. No local photographs or Excel template are required. Metadata and photo filenames, including extensions, come from the CSVs (`location`). Supply the photos later when uploading the output Excel to Wildbook. Ambiguous duplicate filenames are rejected.

The previous Excel template workflow remains under Advanced for compatibility. Excel is the output format. For Camtrap DP, select species and obtain photographs before opening Bulk Import; successful batches and retries are combined during the session.

Edit, add, disable, delete and reorder fields with ↑. `fixed` uses a constant; other sources use record metadata. Use field names supported by your Wildbook. Optional grouping uses event/species/individual in Camtrap, or interval/project/deployment/species in WI. Multi-animal photographs remain separate. Click **Preview / validate**, then **Save Excel**. Nothing is uploaded automatically.

### Optional locality profiles

Enter a name, such as “Doñana”, and click **Save profile**. Use **Load profile** to reuse location, country, submitter and custom fields. Saving the same name updates that profile. Working without saving is supported: no locality is loaded automatically, and preview/export never overwrite profiles.

Profiles are shared across sources in `.local-settings/bulk-import.json`, excluded from Git. The previous profile appears as `Default`. Packaged apps use `LynxAutomator` under `LOCALAPPDATA` or `XDG_CONFIG_HOME`/`~/.config`.

Validation is local and does not check server configuration. Dates preserve source clock time. Grouping events does not establish individual identity.

### Field catalog and metadata comments

Wildlife Insights ZIP input accepts `images.csv` and `images_<project>.csv`, combining image fragments in the same folder with one `deployments.csv`. Separate export folders are rejected to avoid mixing datasets.

Column names have an editable dropdown based on the [Wildbook documentation](https://wildbook.docs.wildme.org/data/bulk-import-beta.html), covering encounters, sightings, projects and other documented categories. Custom names and indexed families remain editable. Photo columns are automatic; subfields such as `.keywords` can be configured. Server support may vary.

Use **Metadata comments** to search and select several source fields, then choose `Sighting.comments`, `Encounter.sightingRemarks` or `Encounter.researcherComments`. Existing configured text is preserved. Example using the `template` source:

```text
Camera: {deployment.cameraID}; Model: {deployment.cameraModel}; Setup: {deployment.setupBy}
```

Available metadata uses `deployment.`, `media.` and `observation.` prefixes. Distinct values are retained when grouping, separated by ` | `. Unknown fields produce an explicit error; blank values stay blank. Templates persist in local profiles. Selected metadata is preserved as notes; keep the original CSVs to preserve their full structure and relationships.

New profiles use `Encounter.sightingID` to link sightings. `Encounter.sightingRemarks` is suitable for comments that persist on cloned encounters. Validation names missing required fields: genus, epithet, year, first photo and location (text, locationID or both coordinates). Comments over Excel's 32767-character limit are rejected instead of silently truncated.

### Shared configuration and additional metadata

WI and Camtrap DP share the editor, validation, profiles, grouping engine and Excel writer. Set the maximum gap in seconds inside the editor when enabling grouping. Explicit Camtrap events are preserved; sequences without an event use the interval. Changes require a fresh preview.

WI accepts additional CSVs such as `projects.csv` and automatically reads extra CSVs from ZIP exports. Rows are linked by project, camera, deployment or image identifiers. A single global row without identifiers can provide common metadata. Preview warns about unmatched tables. PDFs are not converted into fields. Camtrap also exposes `package.*` and additional declared CSV resources. Use these in comments or columns, for example `{projects.project_name}` or `{cameras.camera_model}`.

### Wildbook and multiple location assignments

The normal picker shows only Wildbook names. **Advanced options → Fetch GitHub branches** retrieves the current WildMeOrg/Wildbook branches for selection by name, with URLs kept internal. Local JSON is available only in advanced options. Lists and catalogs are cached; development branches may lack a valid catalog.

Select several deployments or encounters with **Ctrl/Shift** and apply one location. Encounter overrides deployment, then common value; `*` applies to all rows and removes overrides. Save assignments in a profile. Review encounter assignments after changing grouping. Camera coordinates are preserved; GitHub may differ from the deployed server. Dialogs are attached to their owning window to appear in front.
