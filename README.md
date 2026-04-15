# Modern SP Military Org Chart

## Summary

A fully-featured SharePoint Framework (SPFx) Org Chart web part built for military and government organizations. Renders a hierarchical organization chart from a custom SharePoint list, with support for specialty code color coding, drag-and-drop reparenting, inline position management, search/filter, and clean image export.

## Compatibility

| :warning: Important          |
|:---------------------------|
| Every SPFx version is only compatible with specific version(s) of Node.js. In order to be able to build this sample, please ensure that the version of Node on your workstation matches one of the versions listed in this section. This sample will not work on a different version of Node.|
|Refer to <https://aka.ms/spfx-matrix> for more information on SPFx compatibility.   |

![SPFx 1.20.0](https://img.shields.io/badge/SPFx-1.20.0-green.svg)
![Node.js v18](https://img.shields.io/badge/Node.js-v18-green.svg)
![Compatible with SharePoint Online](https://img.shields.io/badge/SharePoint%20Online-Compatible-green.svg)
![Does not work with SharePoint 2019](https://img.shields.io/badge/SharePoint%20Server%202019-Incompatible-red.svg)
![Does not work with SharePoint 2016 (Feature Pack 2)](https://img.shields.io/badge/SharePoint%20Server%202016%20(Feature%20Pack%202)-Incompatible-red.svg)
![Local Workbench Incompatible](https://img.shields.io/badge/Local%20Workbench-Incompatible-red.svg)
![Hosted Workbench Compatible](https://img.shields.io/badge/Hosted%20Workbench-Compatible-green.svg)

## Prerequisites

A SharePoint list with the following columns (internal names configurable in web part settings):

| Column | Default Internal Name | Type | Description |
|---|---|---|---|
| Title | `Title` | Single line of text | Job/role title |
| Unit | `Unit` | Single line of text | Organizational unit |
| Office | `Office` | Single line of text | Office symbol |
| Rank | `Rank` | Single line of text | Rank or grade abbreviation |
| Name | `Name` | Single line of text | Person's full name — leave blank for Vacant |
| Billet | `Billet` | Single line of text | Unique position ID (node key) |
| Owner | `Owner` | Single line of text | Billet ID of supervisor — leave blank for root |
| Code | `Code` | Single line of text | Specialty code for card color coding |

All column names are fully remappable from the web part property pane.

## Features

### Chart Display
- **Hierarchical tree layout** — renders parent/child relationships using the Billet/Owner fields
- **Multi-unit support** — groups root nodes by unit into separate rows; unit/office sort priority is configurable
- **Vacant positions** — nodes with no name display as "Vacant" with a dashed border and reduced opacity
- **Specialty code color coding** — cards colored by specialty code using configurable gradient palettes with a color key legend
- **SharePoint theme aware** — card backgrounds and panel colors adapt to the site's light/dark theme

### Search & Filter
- **Name search** — filters the chart to matching nodes, keeping leadership offices (configurable) always visible
- **Unit filter** — dropdown to scope the chart to a single unit; can be locked to a fixed unit via web part settings
- **Office filter** — secondary dropdown to further narrow by office symbol

### Position Management *(requires list edit permissions)*
- **Add position** — form to create a new position directly from the chart; supports "Save & Add Another"
- **Edit position** — in-place editing of all fields; changing a Billet ID automatically cascades to all direct reports
- **Remove position** — confirmation dialog; direct reports are automatically reparented to the removed node's parent
- **Drag-and-drop reparenting** — drag a card onto another card or the root drop zone to move a position in the hierarchy; cross-unit moves are blocked

### Detail Panel
- Click any card to open a detail panel showing direct reports, total subordinates, unit/office, reports-to, and configurable extra fields
- **Panel columns** — configure which list columns appear in the panel via JSON (supports phone formatting for `Commercial`/`DSN` columns, mailto links for `Email`)

### Card Customization
- **Additional card lines** — configure extra lines shown on each card via JSON (e.g. billet number, email, phone)
- **Full color control** — connector line color, card border color, selected card gradient (start/end/border/font), and all button colors are individually configurable from the property pane

### Export
- **Save as Image** — exports the full chart to a PNG file using `html2canvas`; captures the entire scrollable area regardless of current scroll position, with box-shadows removed for a clean output

### Layout
- **Horizontal drag-to-scroll** — click and drag anywhere on the chart to pan horizontally
- **Full-width modern page support** — optional toggles to hide SharePoint's white space margins, left nav, and top navigation bar for a clean full-width display

## Applies to

- [SharePoint Framework](https://learn.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview)
- [Microsoft 365 developer tenant](https://learn.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Contributors

- [Austin Livengood](https://github.com/TehGoodLivin)

## Version history

| Version | Date | Comments |
|---|---|---|
| 1.0 | April 20, 2025 | Initial release |
| 1.1 | April 15, 2026 | Added specialty code colors, drag-and-drop reparenting, position add/edit/remove, detail panel, search/filter, image export, full-width layout toggles |

## Minimal Path to Awesome

### Build and Test

1. Clone this repo
2. In the command line run:
   ```
   npm i
   gulp build
   gulp serve --nobrowser
   ```
3. Create a SharePoint list with the columns listed in [Prerequisites](#prerequisites)
4. Navigate to the hosted workbench: `https://<tenant>.sharepoint.com/sites/<site>/_layouts/15/workbench.aspx`
5. Add the web part and configure it in the property pane

### Package and Deploy

1. In the command line run:
   ```
   gulp bundle --ship
   gulp package-solution --ship
   ```
2. Upload the generated `.sppkg` from `sharepoint/solution/` to your tenant App Catalog
3. Add the app to your SharePoint site
4. Add the web part to a page and configure the data source settings in the property pane

## Property Pane Reference

### Web Part Settings
| Setting | Description |
|---|---|
| Web Part Title | Title displayed above the chart |
| Hide Excess White Space | Removes SharePoint canvas padding for full-width display |
| Hide Menus | Hides the SharePoint left navigation |
| Hide Top Navigation | Hides the SharePoint site header |

### Data Source Settings
| Setting | Description |
|---|---|
| Site URL | Full URL of the SharePoint site; leave blank for current site |
| List Name | Internal or display name of the list |
| Unit | Lock to a single unit; leave blank for all-units dropdown |
| Title / Unit / Office / Rank / Name / Billet / Owner / Specialty Code Field | Maps each data field to the correct list column |

### Display Settings
| Setting | Description |
|---|---|
| Line Color | Color of connector lines between nodes |
| Card Border Color | Default card border color |
| Selected Gradient Start / End | Gradient for a selected (clicked) card |
| Selected Card Border / Font Color | Colors for selected card border and text |
| Button Colors | Individual color controls for Add, Save, Delete, Cancel, Move, and OK buttons |
| Specialty Code Colors (JSON) | Array of `{Code, GradientStart, GradientEnd, FontColor}` objects that map specialty codes to card colors. Use `"Other"` as a fallback entry. |

### Card Settings
| Setting | Description |
|---|---|
| Additional Card Lines (JSON) | Array of line definitions shown below Unit/Office on each card. Each line is an array of column names and optional literal separators. Example: `[["Billet"],["Commercial"," \| ","DSN"]]` |

### Detail Panel Settings
| Setting | Description |
|---|---|
| Detail Panel Columns (JSON) | Array of line definitions shown in the click-through detail panel. Same format as Card Lines. |

### Sort Settings
| Setting | Description |
|---|---|
| Unit Sort Priority | Comma-separated unit values to pin first when "All Units" is selected |
| Office Sort Priority | Comma-separated office values to pin first within a unit |
| Leadership Offices | Comma-separated office codes that remain visible during search (e.g. `CC,CD,DS`) |

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
