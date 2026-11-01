# Persistent-reference byte round trip

`persistent_roundtrip.bas` is a complete original VBA module. **LIVE_SMOKE_TESTED** for the controlled positive case only.
Target: SOLIDWORKS 2018+ desktop VBA; installed 2026 type-library check passed.
Import with SOLIDWORKS and SOLIDWORKS Constant library references; run `main` and inspect
the Immediate window. Requires a disposable active part, no sketch edit, exactly one
selected supported feature, face, edge or other selectable object.

The macro obtains `GetPersistReference3`, encodes each byte as three decimal digits,
validates and decodes that representation into a VBA Byte array, releases the original
object, rebuilds, then resolves with `GetObjectByPersistReference3`. It checks the state
against `swPersistReferencedObject_Ok`, rejects Nothing and compares the two resolutions
using `ISldWorks.IsSame`. Units: none. With empty RESULT_PATH, output is printed and no filesystem state is written.

Expected: a non-empty `REFERENCE decimal-triplets=...` line and
`RESULT persistent_roundtrip OK; bytes=<positive count>`. No save or close occurs.
The rebuild may mark the model dirty. Selection is retained; it is not used after capture.
Store a document identity/path and configuration with the bytes in a real application.
Do not resolve a token against a different document, assume byte strings never change,
or use a feature name as a persistent topology identifier.

Negative cases: zero/multiple selections; unsupported object; sketch edit; after capturing
a token externally, delete its face/feature or suppress its owner in a private fixture.
Resolution must reject any non-OK state or Nothing. Decoder rejects blank input, lengths
not divisible by 3, non-digits and values above 255. The complete macro exercises an
unchanged-topology in-memory round trip only: close/reopen, suppress/delete and invalid-token
cases need a separate controlled harness/manual test. It does not claim those tests ran.

Official contracts (consulted 2026-11-01; no external example code copied):

- [GetPersistReference3(DispObj)](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDocExtension~GetPersistReference3.html)
- [GetObjectByPersistReference3(PersistId, ByRef ErrorCode)](https://help.solidworks.com/2023/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDocExtension~GetObjectByPersistReference3.html)
- [Persistent Reference IDs](https://help.solidworks.com/2019/english/api/sldworksapiprogguide/Overview/Persistent_Reference_IDs.htm)
- [ISldWorks.IsSame](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISldWorks~IsSame.html)

L1 lint and installed API signatures/constants pass; project compilation, live rebuild,
cross-session persistence and negative integration tests are not verified.

## Optional result log and cleanup

`RESULT_PATH=""` disables file output. Set it to a fresh result file in the reviewed
private run directory when using runner result assertions. The macro refuses an existing
file and never creates directories. It writes its final `RESULT ... OK` only after the
model assertions and protected cleanup; caught model/assertion/cleanup failures write
`FAIL` when the log was opened. Existing/unwritable log paths must fail, without overwriting
prior output. An absent expected result file is a failed test even if RunMacro2 returns True.
No CAD save or close is added by this log option.

The Failed label captures the original error and uses `Resume Cleanup` before activating
`On Error Resume Next` for cleanup. Each cleanup call is followed by an Err check; cleanup
failures suppress overall OK. Log errors use their own handler and `Resume LogCleanup`.
This matters because VBA cannot handle a second exception inside an already active handler.
Negative acceptance should deliberately trigger an input failure with a writable fresh log
and verify that it contains FAIL, then exercise existing/unwritable result paths separately.

Bounded live evidence (2026-11-01): SOLIDWORKS 2026 SP4.0, API 34.4.0, Polish UI; reviewed CLI/COM execution on a disposable copy, no dialogs, no input save, final session IDLE. See [case report](../../../validation/2026-11-01/modeling-final/persistent/runner.json) and [scope/limits](../../../validation/2026-11-01/README.md). Whole-project compilation, other releases and unexecuted negative cases remain NOT_TESTED. Source headers predate this recorded run.
