# FeatureData edits and configuration-specific changes

Use [modify-extrude](../recipes/modeling/feature-modification/README.md) for an original complete
AccessSelections/ModifyDefinition example, and [configuration-variants](../recipes/modeling/configuration-variants/README.md)
for named configuration dimensions. Both validate the controlled block fixture before edits,
rebuild, read back the resulting model and never save it. Optional create-only RESULT_PATH
logs report final OK/FAIL after protected cleanup; inspect those fresh logs, not RunMacro2
True alone, to assess a runtime test. Both controlled positive cases passed on SOLIDWORKS
2026 SP4 (34.4.0, Polish UI), with independent volume/configuration readback in the
[local validation record](../validation/2026-11-01/README.md). Complete-project compilation,
other versions and the full negative-case matrix remain NOT_TESTED.

## Definition lifecycle

1. Resolve the feature in the intended document/configuration. Validate type, suppression,
   required end condition and initial values. Do not cast a face or sketch as feature data.
2. GetDefinition into the correct typed FeatureData interface, e.g. ExtrudeFeatureData2.
3. Check AccessSelections(TopDoc, Component) against False. For these part recipes TopDoc is
   the part and Component is Nothing. Assembly edits require the top assembly and the owning
   component, correct referenced configuration and resolved context; do not reuse part Nothing
   arguments as a universal assembly recipe.
4. Change documented properties/methods, such as SetDepth(Forward, DepthInMeters).
5. Check Feature.ModifyDefinition(Data, TopDoc, Component). Success releases selection access;
   failure or abandonment after successful access requires ReleaseSelectionAccess.
6. Rebuild the intended configuration; reacquire FeatureData and geometric result objects.
   Compare requested values and actual body metrics. Reject rebuild/readback mismatch.

Track whether selection access is held; cleanup is conditional, and cleanup errors are
reported independently. Capture errors in Failed, then use Resume Cleanup to leave the active
VBA handler before protected cleanup calls. Enabling On Error Resume Next inside an active
handler does not protect it from a second exception. Never continue modifying with a cached definition after a failed
commit. Neither successful COM dispatch nor a nonzero Boolean proves a healthy resulting
feature. The recipe's one-body volume check strengthens the assertion, but cannot prove
all geometric details of arbitrary features.

Do not change the definition's properties before AccessSelections merely because a numeric
getter works without it. Do not automatically roll back a production feature by issuing a
second unverified ModifyDefinition; retain the disposable failed fixture and report what
changed. These examples intentionally do not claim a transactional undo guarantee.

## Dimensions and configurations

IDimension.SetSystemValue3(NewValue, WhichConfigurations, Config_names) uses
swSetValueInConfiguration_e and returns swSetValueReturnStatus_e. Compare to
swSetValue_Successful, not True/False. GetSystemValue3 uses swInConfigurationOpts_e and returns
a value array; validate Empty/array bounds before indexing unfamiliar results. The configuration
name argument matters only with the documented specific-configurations option.

IFeature.Parameter("D1") accepts a short parameter name. IModelDoc2.Parameter requires its
full qualified name, e.g. D1@OwnedFeature. Neither is a universal semantic lookup for depth.
In the supplied generated block fixture D1 is checked against ExtrudeFeatureData2.GetDepth
before use; a real application should own a dimension name/attribute or validate its role.

Capture ConfigurationManager.ActiveConfiguration.Name. Do not hard-code Default or Domyslna.
Create variants under new known names and fail on collisions. Specify the setter scope;
activate each configuration and call ForceRebuild3(False) before reading its geometry.
ForceRebuild3 does not verify every configuration by itself. Restore the original configuration
on success and attempt restoration after failure; restoration failure must suppress overall OK.
Never delete a pre-existing configuration as cleanup for a failed example.

## Acceptance contract

- Use fresh copies of the 40x20x10 mm generated fixture; modification and variants are separate runs.
- Edit result: depth 0.015 m, one solid, volume 12e-6 m³.
- Variant result: named 0.010/0.020 m depths and 8e-6/16e-6 m³, original active configuration
  restored with unchanged 8e-6 m³; three configurations total.
- Test missing feature, wrong type, non-blind end, altered fixture, multiple configurations,
  failed access/commit and failed restoration. No failed branch may produce the overall OK line.
- Record L1, installed API, VBA compile, execution and independent result/safety assertions
  separately. Installed 2026 TLB checks passed for the shipped examples; they are not compile
  or runtime evidence and do not certify 2018–2025 installations.

## Official source contracts

Consulted 2026-11-01; original implementations, no copied external macro code.

- [AccessSelections](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IExtrudeFeatureData2~AccessSelections.html)
- [Official access/modify/release explanation](https://help.solidworks.com/2016/English/api/sldworksapi/Access_Selections_Example_VB.htm)
- [ModifyDefinition](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeature~ModifyDefinition.html)
- [ReleaseSelectionAccess](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IExtrudeFeatureData2~ReleaseSelectionAccess.html)
- [GetDepth](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IExtrudeFeatureData2~GetDepth.html)
- [SetDepth](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IExtrudeFeatureData2~SetDepth.html)
- [IFeature.Parameter](https://help.solidworks.com/2016/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeature~Parameter.html)
- [SetSystemValue3](https://help.solidworks.com/2025/English/api/sldworksapi/SolidWorks.interop.sldworks~SolidWorks.interop.sldworks.IDimension~SetSystemValue3.html)
- [GetSystemValue3](https://help.solidworks.com/2017/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IDimension~GetSystemValue3.html)
- [AddConfiguration3](https://help.solidworks.com/2016/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDoc2~AddConfiguration3.html)
