# CLI observers and fixture helpers

All are original VBA modules. Use fresh output paths inside the reviewed test root,
pin prepared constants and hashes in the runner profile, and inspect outputs explicitly.
They do not select production documents by title or save an original CAD file.

| Module | Input / settings | Observable result / side effects |
| --- | --- | --- |
| `selection_diagnostic.bas` | Active copied document; optional new `RESULT_PATH`; optional `EXPECTED_SELECTION_TYPE` | Selection API type, feature type/name, mark and owner. It does not perform a selection operation. Unspecified expected type yields OBSERVED, not PASS. |
| `model_state_snapshot.bas` | Active copied document; new `RESULT_PATH` | JSON with feature/configuration inventory, selection count, solid count/volume and sketch measurements. `collection_status: OK` means collected, not geometrically correct. |
| `boolean_probe.bas` | Active disposable part; new `RESULT_PATH`; edit-enabled profile | Checks numeric 0/1/-1 comparisons and actual COM truth representation. Deliberately marks only the copied model dirty; never saves it. |
| `create_blank_part_fixture.bas` | Existing copied input for runner; copied local `TEMPLATE_PATH`; new `OUTPUT_PATH` and `RESULT_PATH` | Creates and saves a new blank part from the copied template, closes only its newly created document. Refuses existing outputs. |
| `save_fixture_copy.bas` | Active generated test model; exact `EXPECTED_INPUT_PATH`; new `OUTPUT_PATH` and `RESULT_PATH` | Saves a distinct native copy using SaveAs Copy; never overwrites a result or saves the active input. Intended as an explicitly reviewed fixture-preparation observer. |

For geometry measurements use the existing `../templates/source_metrics_template.bas` with
`METRICS_PATH`; it supports solid parts and produces millimeter-based body metrics.

Modeling smoke preparation is reproducible with:

```powershell
python scripts/smoke_modeling.py --blank C:\SWTest\blank.sldprt --block C:\SWTest\block.sldprt --test-root C:\SWTest\new-smoke
```

The blank fixture has one configuration, no sketch/solid/surface and the three default planes.
The block is the saved 40×20×10 mm `sketch_extrude` result, with `Recipe_Profile` and `Recipe_Block`.
The command above only prepares copies/profiles. Add `--run` to intentionally run the reviewed
recipes through COM. It uses separate copied inputs per case, terminal result logs, independent
state snapshots and source hash checks. Never automatically retry UNKNOWN sessions.
