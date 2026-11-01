# Selection, marks and durable identities

Start with [plane-selection](../recipes/selection/plane-selection/README.md) and
[persistent-reference](../recipes/selection/persistent-reference/README.md). Both are complete
original VBA recipes, with inputs, negative cases and explicit evidence status.

## Object selection and language independence

Use a typed object reference when available: IFeature.Select2(Append, Mark),
ISketchSegment.Select4(Append, SelectData), or IEntity.Select4(Append, SelectData).
Check every returned Boolean with `= False`/`<> False`. An invariant feature type from
IFeature.GetTypeName2 is different from a user-visible localized Name. A recorded name
such as Front Plane, Sketch1, Default or Line1 is not a durable identity.

The executable plane recipe enumerates exactly three RefPlane features and selects the
first actual feature object. That gives a reproducible plane in the controlled blank fixture,
not a geometric classification of arbitrary reference planes. For a required global normal,
inspect IRefPlane.Transform and use MathTransform/MathPoint transformations explicitly;
reject ambiguous candidates. Do not infer geometric orientation from a localized plane name.
2D sketch coordinates are local; 3D sketch coordinates are model coordinates.

SelectionManager indices are 1-based; returned arrays use their own LBound/UBound.
Pass mark -1 to inspect all selections, or the exact positive/zero mark required by the
operation. Validate count and type before casting GetSelectedObject6. ClearSelection2
must be deliberate: these modeling recipes consume/clear selection and document that side
effect; an inspection-only tool should preserve selection or use an isolated selection list.

| Operation in this library | Selection role | Mark |
|---|---|---:|
| New 2D sketch | Reference plane | 0 |
| Extrusion/cut | Closed sketch | 0 |
| Revolve | Sketch / rotation axis | 0 / 4 |
| Part linear pattern | Direction 1 / direction 2 / seed features | 1 / 2 / 4 |
| Part circular pattern | Direction axis / seed features | 1 / 4 |

Marks are operation-specific. Component patterns and mates use different contracts;
never transfer a part-pattern mark table to an assembly operation.

## Persistent references

Capture a supported object with GetPersistReference3(DispObj). Keep the resulting byte
array associated with its owning document and intended configuration. Resolve with
GetObjectByPersistReference3(PersistId, ByRef ErrorCode); require both
swPersistReferencedObject_Ok and a non-Nothing object. Suppression, deletion, document
replacement and topology edits require explicit handling; a failed lookup must not select a
nearby edge as a fallback. Returned plane/axis/configuration references can be base features;
use IFeature.GetSpecificFeature2 when the documented specific interface is needed.

Serialize bytes losslessly; do not parse the internal token or promise the byte sequence is
immutable. The recipe validates decimal triplets, rebuilds and resolves the decoded bytes.
Compare objects through ISldWorks.IsSame rather than VBA pointer `Is`. Compare persistent IDs
using the API's identity semantics when needed, not a claim about token byte stability.
Feature names and dimensions' fully qualified names identify editable application objects
within a known context; they are not persistent topology identities.

The recipe tests only the same open document after rebuild. Cross-session acceptance must
save a **private** fixture and token, close/reopen that same fixture through the reviewed runner,
resolve and check type/geometry; then exercise suppression, deletion and a different document.
No production document should be closed for a demonstration.

## Official source contracts

Consulted 2026-11-01; links establish API semantics, not execution evidence. Recipes are
original; no external macro source was copied.

- [IFeature.GetTypeName2](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeature~GetTypeName2.html)
- [IFeature.Select2](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeature~Select2.html)
- [ISketchSegment.Select4](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchSegment~Select4.html)
- [IRefPlane.Transform](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IRefPlane~Transform.html)
- [GetObjectByPersistReference3](https://help.solidworks.com/2023/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDocExtension~GetObjectByPersistReference3.html)
- [Persistent Reference IDs guide](https://help.solidworks.com/2019/english/api/sldworksapiprogguide/Overview/Persistent_Reference_IDs.htm)
- [FeatureRevolve2 selection contract](https://help.solidworks.com/2024/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ifeaturemanager~featurerevolve2.html)
- [Linear pattern FeatureData](https://help.solidworks.com/2023/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ILinearPatternFeatureData.html)
- [Circular pattern FeatureData](https://help.solidworks.com/2020/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ICircularPatternFeatureData.html)
