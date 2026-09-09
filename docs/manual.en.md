# LynxAutomator user manual

[Español](manual.md) · [Installation and packages](../README.md#linux-macos-and-windows-builds)

This revised manual describes the full application on `feature/cross-platform-builds`.
Older release executables may behave differently. The original
[DOCX guide](../WIP%20LynxAutomator%20GUIDE%20.docx) is retained as a historical
reference, including its screenshots; current instructions are maintained in Markdown.

## 1. Getting started and editions

LynxAutomator prepares spreadsheets, downloads authorized Wildlife Insights images
and provides camera-trap utilities. It does not automatically upload spreadsheets
or photos to Wildbook. Select Spanish, Portuguese or English in the interface;
some messages remain in English. Save results before changing language: doing so
rebuilds the forms. Language changes are blocked during an active download.

| Feature | Full | Historical alpha mini |
| --- | --- | --- |
| Wildbook folder import and catalog | Yes | Yes |
| WI download and CSV conversion | Yes | Yes, with logic differences |
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

Stop takes effect after the current transfer, which has a five-minute timeout.
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

All extracted frames receive the same base timestamp: the video's filesystem
creation time on Windows or modification time on Linux/macOS. The app does not
read the embedded video recording date or add each frame's offset. Copied videos
may therefore produce dates that do not represent capture time. This qualifies
the old manual's statement about preserving capture dates.

## 10. Image Renamer

Select a source folder (including subfolders), choose filename components, then
rename. Supported image extensions are JPG/JPEG/PNG/GIF. Enable Copy Photos and
choose a different destination to preserve originals; otherwise files are renamed
in their existing locations.

Options include the first word of the folder name, original filename, EXIF date,
custom text and replacing spaces with underscores. Folder/original-name options
apply capitalization. Custom text cannot contain `/` or `\`. Missing EXIF may
appear as `None` when the date option is enabled; disable that option for such files.

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
- Busy window: video, Excel and renaming still run in the GUI thread. Try a small batch.

Save before closing or changing language. There is no universal undo history.
Generating a spreadsheet does not certify all Wildbook rules or dataset integrity.
See the [logic review](logic-review.md) for remaining improvements.
