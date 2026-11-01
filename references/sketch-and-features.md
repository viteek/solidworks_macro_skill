# Sketches, features, holes and patterns

All supplied modeling modules are original complete VBA files with `main`, explicit input
guards and result assertions. Import with SOLIDWORKS and SOLIDWORKS Constant library references.
Use disposable fixtures and the reviewed runner; read each folder's README before execution.
Controlled positive cases passed on SOLIDWORKS 2026 SP4 (34.4.0, Polish UI), using copied
fixtures and independent readback; see [local validation](../validation/2026-11-01/README.md).
Unexecuted branches and releases remain NOT_TESTED. L1 lint and installed 2026 TLB member,
arity and constant checks do not establish complete-project compilation or geometry by themselves.

| Intent | Executable recipe | Result checked by the macro |
|---|---|---|
| Fully constrained 2D parametric profile and boss | [sketch-extrude](../recipes/modeling/sketch-extrude/README.md) | Constraint status, one solid, volume |
| Ordinary cylindrical through hole and linear feature pattern | [hole-linear-pattern](../recipes/modeling/hole-linear-pattern/README.md) | Seed/final volume, cylindrical-face count, instance count |
| Full-turn revolve around a construction line | [revolve](../recipes/modeling/revolve/README.md) | Rebuild, one solid, annular-cylinder volume |
| Circular feature pattern of axial holes | [circular-pattern](../recipes/modeling/circular-pattern/README.md) | Seed/final volume, count, pitch |
| Model-space 3D sketch polyline | [sketch-3d](../recipes/modeling/sketch-3d/README.md) | Is3D, segment types/count, total length |
| Modify or configure existing depth | [feature modification](feature-modification.md) | Fresh definition and rebuilt per-configuration geometry |

## Coordinates, units and sketch lifecycle

Use meters for length, radians for angles and m³ for volume regardless of display units.
Use variable names that carry units. A 10 mm depth is 0.010, a full turn is 2*pi, not 360.
An active 2D sketch uses its own coordinate system. ISketch.ModelToSketchTransform maps model
to sketch; its inverse maps the other way. A 3D sketch uses model coordinates and does not
require a plane. Do not paste face/world coordinates into a planar sketch without transformation.

Before toggling InsertSketch or Insert3DSketch, require no active sketch. After opening, verify
ActiveSketch and, for a 3D operation, Is3D. Check every created segment/feature for Nothing.
Keep a typed SketchSegment; GetSketch returns the owning sketch object, which can expose its
Feature interface (as in official access examples). Use Feature.GetSpecificFeature2 in the
other direction. ISketch has no GetFeature member. Reacquire entities after leaving edit mode
before selecting them; selection validity can depend on the edit context in which they were obtained.

AddToDB=True suppresses inference and improves repeatability while creating geometry.
Restore the prior setting even on failure, and switch it off before asking the solver for
relations/dimensions. Coordinate coincidence alone does not prove a merged-point relation.
Do not claim that a 3D polyline is a constrained sweep path just because its endpoint numbers
match. These recipes deliberately leave partial geometry for inspection after an error rather
than deleting unknown user features.

## Defining the profile

The rectangle recipe fixes one corner, then uses FullyDefineSketch to add relations and
dimensions for the rest. Fixing the entire sketch would prevent a meaningful parametric
example. The method's ten parameters are: EntitiesToFullyDefine, UseRelations,
RelationsToApply, UseDimensions, HorizontalDimScheme, HorizontalDatumDisp, VerticalDimScheme,
VerticalDatumDisp, HorizontalDimPlacement, VerticalDimPlacement. The published return value
is undefined: inspect GetConstrainedStatus and require swFullyConstrained (3).
It must reject under/over-constrained, no-solution and autosolve-off states.

For manually placed dimensions, AddHorizontalDimension2/AddVerticalDimension2 operate on
the current selected entities and return a display dimension. GetDimension2(0) gives its
model dimension; SetSystemValue3 sets SI units and returns swSetValueReturnStatus_e, not
a Boolean. Getter and setter configuration enums differ. Do not infer a successful solve
from a non-Nothing dimension alone; rebuild and inspect geometry/constraint status.

## Extrude, cut and revolve contracts

FeatureExtrusion3 has 23 ordered arguments:
`Sd, Flip, Dir, T1, T2, D1, D2, Dchk1, Dchk2, Ddir1, Ddir2, Dang1, Dang2,
OffsetReverse1, OffsetReverse2, TranslateSurface1, TranslateSurface2, Merge, UseFeatScope,
UseAutoSelect, T0, StartOffset, FlipStartOffset`.
FeatureCut4 has 27: the first 17 through TranslateSurface2, then `NormalCut, UseFeatScope,
UseAutoSelect, AssemblyFeatureScope, AutoSelectComponents, PropagateFeatureToParts, T0,
StartOffset, FlipStartOffset, OptimizeGeometry`. Its normal direction is opposite to a boss;
the hole recipes reverse it from the same base plane into the positive-normal boss.

FeatureRevolve2 has 20 arguments: `SingleDir, IsSolid, IsThin, IsCut, ReverseDir,
BothDirectionUpToSameEntity, Dir1Type, Dir2Type, Dir1Angle, Dir2Angle, OffsetReverse1,
OffsetReverse2, OffsetDistance1, OffsetDistance2, ThinType, ThinThickness1, ThinThickness2,
Merge, UseFeatScope, UseAutoSelect`. Mark sketch=0 and axis=4 (16 is also documented).
Keep a non-crossing closed profile on one side of the axis for the simple annular example.
Do not silently accept an open profile by turning a requested solid into a thin feature.

## Holes and pattern semantics

The through-hole recipe is a circular cut, suitable for demonstrating cylinder geometry.
It carries no Hole Wizard database/standard/thread metadata. A normed Hole Wizard recipe
must bind installed standard, hole type and size, verify definition/readback and compare the
result to the requested engineering standard. Database availability and localization matter;
inventing a standard/size string or enum is not an acceptable shortcut.

Use CreateDefinition/CreateFeature for part linear/circular patterns; the older
FeatureLinearPattern4/FeatureCircularPattern4 calls are obsolete. Marks and seed counts are
explicit in the recipes. Total instances includes the seed. Linear pitch is meters; circular
pitch is radians. The circular recipe sets EqualSpacing=False and pitch pi/2, avoiding any
ambiguity between a total angle and adjacent spacing. Rebuild and reacquire FeatureData;
check counts and solid geometry. Verify actual instance positions as a separate result
assertion: volume alone can match for a misplaced feature.

## Official sources

Consulted 2026-11-01. API signatures/semantics informed original code; no external macro copied.

- [CreateCornerRectangle](https://help.solidworks.com/2018/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISketchManager~CreateCornerRectangle.html)
- [FullyDefineSketch](https://help.solidworks.com/2015/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISketchManager~FullyDefineSketch.html)
- [AddHorizontalDimension2](https://help.solidworks.com/2016/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDoc2~AddHorizontalDimension2.html)
- [GetDimension2](https://help.solidworks.com/2019/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.idisplaydimension~getdimension2.html)
- [FeatureExtrusion3](https://help.solidworks.com/2024/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ifeaturemanager~featureextrusion3.html)
- [FeatureCut4](https://help.solidworks.com/2024/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ifeaturemanager~featurecut4.html)
- [FeatureRevolve2](https://help.solidworks.com/2024/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ifeaturemanager~featurerevolve2.html)
- [CreateLine](https://help.solidworks.com/2019/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISketchManager~CreateLine.html)
- [Insert3DSketch](https://help.solidworks.com/2016/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISketchManager~Insert3DSketch.html)
- [IBody2.GetMassProperties](https://help.solidworks.com/2020/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.ibody2~getmassproperties.html)

Read source-specific caveats before adapting a method to assemblies, multibody parts,
thin features or sheet metal. A tested toy part would not certify those broader contexts.
