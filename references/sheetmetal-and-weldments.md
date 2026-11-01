# Sheet metal, weldments and cut-list diagnostics

Runnable original example: [inspect_sheetmetal_cutlist.bas](../recipes/sheetmetal/inspect_sheetmetal_cutlist.bas). The ordinary-solid/no-cut-list branch passed a reviewed CLI/COM smoke on SOLIDWORKS 2026 SP4 (34.4.0); positive sheet-metal and cut-list branches remain **NOT_LIVE_TESTED**. See [scope/evidence](../recipes/sheetmetal/README.md). Static API checks passed against installed R2026x libraries. The recipe performs no flattening, cut-list update, configuration activation, rebuild or save.

## Body and feature evidence

[PartDoc.GetBodies2(bodyType,visibleOnly)](https://help.solidworks.com/2024/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ipartdoc~getbodies2.html) supports solid/surface bodies, with potentially variable return order. Use body identity/name for diagnostics rather than assuming array order is stable. `visibleOnly=False` includes hidden bodies; suppressed geometry is still configuration dependent.

[Body2.IsSheetMetal](https://help.solidworks.com/2025/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IBody2~IsSheetMetal.html) is Boolean. The documented lightweight-component exception can return False for sheet-metal bodies obtained through `Component2.GetBodies3`; resolve the context before classifying them. The part-only recipe avoids that assembly ambiguity. A body passing this test and a `SheetMetal` feature are distinct observations.

For a `SheetMetal` feature, `Feature.GetDefinition` provides `SheetMetalFeatureData`; [Thickness](https://help.solidworks.com/2011/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISheetMetalFeatureData~Thickness.html) is a scalar feature parameter, reported in model SI length units (meters). Do not equate one feature's thickness with every body in a multibody part. Report the owning feature and compare against a test fixture.

The presence of a `FlatPattern` feature does not establish its suppression state, that the active configuration is flat, or that a DXF export succeeded. Export verification additionally needs a selected body/flat-pattern mapping, configuration and output-file evidence. This recipe intentionally reports presence without changing suppression.

## Cut-list scope and cached properties

The official [body-folder traversal guide](https://help.solidworks.com/2024/english/api/sldworksapiprogguide/OVERVIEW/Accessing_Bodies_in_Body_Folders.htm) distinguishes `SolidBodyFolder`, `SurfaceBodyFolder`, `CutListFolder`, `SubWeldFolder` and `SubAtomFolder`. The top-level cut-list node is a solid-body folder; its item nodes are cut-list folders. Obtain `BodyFolder` through `Feature.GetSpecificFeature2` and traverse subfeatures explicitly.

[GetBodyCount](https://help.solidworks.com/2024/English/api/sldworksapi/SolidWorks.interop.sldworks~SolidWorks.interop.sldworks.IBodyFolder~GetBodyCount.html) and `GetBodies` refer to that folder's bodies, excluding descendant-folder bodies. A root-folder count is not a recursive cut-list quantity. An empty item can be a stale/suppressed artifact; diagnose it without silently deleting it.

Read properties from the **cut-list feature's** `CustomPropertyManager`, rather than document/configuration property managers. [Get6](https://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ICustomPropertyManager~Get6.html) returns a status, raw expression, evaluated text, resolution flag and link flag. The recipe requests cached values and labels freshness unknown. Fresh evaluation may traverse configurations; a generalized writer must specify state restoration and update policy before requesting that behavior.

Manual verification on copies: compare body names with folder membership, inspect SheetMetal feature thickness, compare raw/evaluated cut-list properties, then manually switch configuration and rerun. A weldment folder or drawing weld symbol alone does not establish structural-member dimensions, weld geometry or manufacturing validity.
