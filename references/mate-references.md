# Mate references: exact insertion contract and readback

Runnable original example: [mate_reference.bas](../recipes/assemblies/mate_reference.bas). A three-planar-face fixture passed creation and three-slot readback on SOLIDWORKS 2026 SP4 (34.4.0), using a reviewed CLI/COM wrapper and a copied block part. [Scope/evidence](../recipes/assemblies/README.md): optional one/two-slot and negative cases remain **NOT_LIVE_TESTED**. Installed R2026x API/member/argument/constant checks passed; whole-project compilation remains **NOT_TESTED**.

A mate reference is a feature describing candidate geometry for automatic mating; creating one does not itself constrain two assembly components. Its API is `IFeatureManager.InsertMateReference2`, not an `InsertMateReference` overload with an invented axis flag.

The [official InsertMateReference2 signature](https://help.solidworks.com/2025/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeatureManager~InsertMateReference2.html) has exactly twelve positional arguments:

| Position | Parameter | VBA type / meaning |
|---|---|---|
| 1 | BstrMateReferenceName | String; feature/reference name |
| 2 | PrimaryReferenceEntity | Entity |
| 3 | PrimaryReferenceType | Long; `swMateReferenceType_e` |
| 4 | PrimaryReferenceAlignment | Long; `swMateReferenceAlignment_e` |
| 5 | PrimaryReferenceAlignAxes | Boolean; origin/coordinate-system axis alignment |
| 6 | SecondaryReferenceEntity | Entity |
| 7 | SecondaryReferenceType | Long; reference-type enum |
| 8 | SecondaryReferenceAlignment | Long; alignment enum |
| 9 | SecondaryReferenceAlignAxes | Boolean; origin/coordinate-system axis alignment |
| 10 | TertiaryReferenceEntity | Entity |
| 11 | TertiaryReferenceType | Long; reference-type enum |
| 12 | TertiaryReferenceAlignment | Long; alignment enum |

There is **no tertiary AlignAxes argument**. The result is `Feature`; reject Nothing. The example supplies explicit face objects for one to three planar faces, `Nothing` for unused entity slots, coincident reference types, Any alignment and False for the two axis flags. This tests optional reference slots; it does not claim to test coordinate-system axis alignment. No name-based or coordinate-based geometry selection is hidden in the recipe.

Use the mate-reference enums, not assembly mate enum values: [swMateReferenceType_e](https://help.solidworks.com/2024/English/api/swconst/SolidWorks.Interop.swconst~SolidWorks.Interop.swconst.swMateReferenceType_e.html), [swMateReferenceAlignment_e](https://help.solidworks.com/2026/English/api/swconst/SolidWorks.Interop.swconst~SolidWorks.Interop.swconst.swMateReferenceAlignment_e.html). For example coincident reference type is 2; `swMateCOINCIDENT` is a different enum contract.

Get `MateReference` using [Feature.GetSpecificFeature2](https://help.solidworks.com/2023/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IMateReference.html?id=2.2.1.213). Read [ReferenceEntityCount](https://help.solidworks.com/2024/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IMateReference~ReferenceEntityCount.html) **before** reading slots: it describes three primary/secondary/tertiary slots, including empty slots. [Slot indexes](https://help.solidworks.com/2025/English/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swMateReferenceIndex_e.html) are 0, 1 and 2.

For each slot get [ReferenceEntity2(index)](https://help.solidworks.com/2019/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IMateReference~ReferenceEntity2.html?format=P&value=); check Nothing, then inspect [ReferenceEntityType](https://help.solidworks.com/2024/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IMateReference~ReferenceEntityType.html) before casting the returned Entity/SketchPoint. Compare [ReferenceType](https://help.solidworks.com/2024/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IMateReference~ReferenceType.html) and [ReferenceAlignment](https://help.solidworks.com/2024/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IMateReference~ReferenceAlignment.html). Empty slots may report -1. Do not equate three slots with three selected entities, or assert that this interface exposes axis-flag readback when it does not.
