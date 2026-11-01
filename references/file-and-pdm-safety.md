# File and PDM scope for CAD automation

Code generation and static validation do not authorize live edits or persistence. The [assembly recipes](../recipes/assemblies/README.md) change the in-memory test document; [drawing view creation](../recipes/drawings/create_three_views.bas) creates a new unsaved document and restores the original active part. The [drawing inspection](../recipes/drawings/inspect_drawing_semantics.bas) and [sheet-metal inspection](../recipes/sheetmetal/README.md) recipes request read-only diagnostics. The folder READMEs record explicitly authorized bounded CLI/COM live tests and remaining **NOT_LIVE_TESTED** scopes. The recipes do not call save, PDM checkout/check-in, delete or close; separate reviewed test fixture/observer macros saved and closed only newly generated local fixtures.

The three assembly modules can write a separately configured fresh completion log. `RESULT_PATH` defaults to empty; when configured, create-only file creation precedes CAD work and refuses overwriting existing logs. This filesystem diagnostic write is separate from saving CAD documents. An incomplete log, logging error or successful macro-dispatch Boolean cannot establish successful CAD results.

## Before any live write

Identify the exact full path/document type, active configuration, unsaved-change state, read-only state, output path and overwrite policy. Use an isolated disposable fixture/copy; for assemblies also account for referenced parts, external references and virtual components. Capturing the active document title alone is insufficient identity. Make a reviewable change and verify it before any separately authorized save/export/check-in.

Rebuild, configuration switching, flattening, resolving components, cut-list updates and mate solving can affect in-memory state or dependent values without a Save call. Report these side effects. A failed creation can leave a feature; the examples report that possibility instead of attempting an unreviewed deletion/rollback. Closing without saving can discard unrelated existing changes, so do not use it automatically as cleanup.

## What a filesystem check cannot establish

A writable local file is not proof of PDM checkout, latest cache version, ownership, workflow permission or permission to check in. A path that looks like a vault folder is not sufficient evidence either. PDM and SOLIDWORKS document state must be inspected through their respective APIs and the actual vault policy.

[IEdmVault5.GetFileFromPath](https://help.solidworks.com/2023/english/api/epdmapi/EPDM.Interop.epdm~EPDM.Interop.epdm.IEdmVault5~GetFileFromPath.html) maps a filesystem path to `IEdmFile5` and optionally its parent folder after vault login; handle a missing result and login errors. Do not infer unmanaged status merely from an arbitrary lookup failure. This repository does not supply credentials, vault login automation or a PDM type-library validation claim.

[IEdmFile5.LockFile](https://help.solidworks.com/2024/english/api/epdmapi/epdm.interop.epdm~epdm.interop.epdm.iedmfile5~lockfile.html?format=P&value=) performs **checkout**, using parent folder ID, parent window handle and lock flags. It can be denied by permission, an existing lock or a vault plugin. Its documentation links `UnlockFile` (check-in) and `UndoLockFile`; these are distinct operations with different effects, not cleanup synonyms. Multi-file work requires a defined batch policy rather than repeatedly checking out files opportunistically.

No PDM mutations are implemented or live-tested here. To add them, explicitly define vault identity, cached/latest version policy, current lock owner, checkout scope, reference handling, write target, check-in comment/version policy and recovery behavior. Validate the installed PDM interop and official method contracts separately from `sldworks.tlb`, then execute on a designated vault test fixture. A successful SOLIDWORKS Save does not establish successful PDM check-in.
