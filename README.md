# Plant3D.ProjectRuntimePalettes

Runtime-generated P&ID palette for Plant 3D 2026 (.NET 8).

This revision reads the project structure from modern Plant project sources:
- `ProcessPower.dcf` (SQLite) for the full P&ID class hierarchy and class attributes
- `*PnIdPart.xml` for additional graphical-style rule candidates
- `projSymbolStyle.dwg` for exact symbol-style-to-block mapping and rendered symbol thumbnails

## Behavior

- Command: `PROJECTPIDPALETTES`
- Prompts for mode before opening the dockable palette:
  - `All`
  - `RespectSupportedStandards`
- Left side: full expandable project class tree for the requested palette groups
- Toolbar buttons at the top-left of the window: expand/collapse the whole tree and hide/show the left tree panel
- Search box next to the toolbar buttons filters the class tree while keeping parent paths visible
- Right side: only styled leaf classes that are intended to be inserted
- Clicking an item inserts it through the Plant P&ID runtime API directly, without relying on pre-existing palette definitions
- Hover tooltip shows the resolved symbol block name, so style-to-block mapping issues are easier to diagnose

## Requested palette roots

- Equipment
- Valves
- Fittings
- Speciality
- Reducers
- Instrumentation
- Lines
- Nozzles
- NonEngineeringItems

## Inclusion rules

### All
Shows every class under the requested P&ID class tree that has a graphical style.

### RespectSupportedStandards
Shows only classes with a graphical style whose supported-standard mask matches the current project standard.

### Optional `tpincluded`
If the boolean class property `tpincluded` can be resolved for all style-bearing classes in scope, the palette only includes classes where `tpincluded = true`.
If the property is missing on at least one class, it is ignored for that project.

## Preview resolution

1. Open `projSymbolStyle.dwg`
2. Read the Named Object Dictionary path `Autodesk_PNP -> PNP_STYLES`
3. Resolve each style entry to its `ACPPASSETSTYLE` object
4. Read DXF group code `4` from that asset-style object to obtain the block name
5. Render the resolved DWG block directly as the symbol thumbnail
6. Use the project line-style information for line previews
7. Only if the exact block cannot be rendered, show a generated fallback thumbnail

## Insertion path

1. Resolve the class and style from `ProcessPower.dcf`
2. Ensure the style exists in the current drawing by copying it from `projSymbolStyle.dwg` when needed
3. Resolve the style id through the Plant style API
4. Create the P&ID object directly through the Plant runtime API (`AssetAdder` / `LineSegmentAdder`)

## Deployment note

Because the project uses `Microsoft.Data.Sqlite`, load the plugin from the full build output folder, not by copying only the main DLL.

## Notes on project symbol mapping

- Symbol thumbnail lookup uses the strict project path: `Autodesk_PNP` -> `PNP_STYLES` -> style entry ObjectId/handle -> DXF export scan for `ACPPASSETSTYLE` group code `4` block names.
- The symbol-style DWG is also set as the temporary `WorkingDatabase` while the style map and block preview index are built, because side-database ObjectDBX work can depend on the active working database context.
