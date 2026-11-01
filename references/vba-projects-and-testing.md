# Complete VBA project compilation and evidence

A `.swp` is a native SOLIDWORKS VBA project. An exported `.bas` is one standard module; `.cls` files hold class modules and forms require both `.frm` text and associated `.frx` resources. Collect all modules, forms, references, constants, entry points and external dependencies before review. An isolated module can pass static checks while the actual project fails because of missing references, duplicate names, incorrect event signatures or absent form resources.

## Review and compile the actual project

1. Copy the complete project and fixtures into a disposable test directory. Record SHA256 for the `.swp` and an exported-source manifest covering every `.bas`, `.cls`, `.frm`, `.frx` and dependency. Record the original project/fixture locations separately without running against them.
2. Open the copied `.swp` using SOLIDWORKS **Tools → Macro → Edit**. Inspect the Project Explorer, project name and every module/class/form. Export source/resources for review; verify module names and the intended public entry procedure. Source-only deliveries must be imported into a complete test project with the required forms/classes.
3. In VBE **Tools → References**, record every reference, library version/location and any `MISSING` entry. Resolve references to the intended installed libraries. Record SOLIDWORKS version, service pack, API revision, architecture and UI language. Inspect conditional compilation constants and Windows API declarations for the target architecture.
4. Select VBE **Debug → Compile <project>**. Fix every compile error and repeat on the entire copied project. A missing/disabled Compile command is not sufficient fresh evidence by itself: record the actual compile operation or mark this phase `NOT_TESTED`. Do not invoke the entry procedure as a substitute for explicit compilation.
5. Save only the reviewed disposable project if compilation/import altered it, close its editing session, then hash the final bytes and refresh its reviewed runner profile. Confirm the persisted module/procedure inventory with `ISldWorks.GetMacroMethods(path, swAllMethods)` (`swAllMethods = 3` in the checked 2026 library). Compare it with the exported-source manifest before execution: an in-memory edited project can differ from the saved binary. Retain compilation evidence identifying the project, time, references and source manifest. Evidence must bind to the accepted final artifact, not an earlier export or a different saved project.
6. Run only the intended entry point against declared fixture copies, capture independent observations before cleanup, and evaluate the task assertions and source/output safety evidence separately.

The runner supports native invocation with `--module ModuleName --entry ProcedureName`, hash-bound to `entry_point` in the mandatory reviewed profile. For text sources it creates a temporary `.swb` with prepared constants and calls `RunMacro2`; this path tests the submitted text invocation. It does not demonstrate full `.swp` compilation, class/form behavior or dependency closure. Native `.swp` invocation likewise leaves the runner's `compilation_status: NOT_TESTED`; attach explicit VBE evidence for that phase.

For a **CLI-only** test requirement, use the Python/PowerShell COM runner and installed-library checks.
Do not fall back to screen, keyboard or UI automation. `RunMacro2` is execution, not a documented
compile-only interface for an entire VBA project. Keep L3 `NOT_TESTED` until matching compilation
evidence is supplied; successful `.swb`/`.swp` invocation must not promote that phase automatically.

## Keep the validation claims separate

| Stage | What it establishes | What remains separate |
| --- | --- | --- |
| Static review | Checked source patterns and declared types | Actual compiler, dynamic dispatch and CAD behavior |
| Installed API check | Checked member/signature/constants in actual libraries | Complete VBA semantics, callbacks and runtime context |
| Project compile | Complete target VBE project compiles with recorded references | Successful execution and correct geometry/documents |
| Runtime | Intended entry returns with recorded errors/dialogs | Result correctness and unintended side effects |
| Result verification | Measured task assertions match expected values | Source/output safety and collector provenance |
| Safety audit | Authorized writes/state restoration and before/after evidence | Other tasks, projects, releases or service packs |

Use the [six-phase verifier contract](../benchmarks/README.md); missing evidence is `NOT_TESTED`. The runner JSON can support runtime evidence but is not itself a full result-verifier observation document. Offline unit tests check Python contracts and rejection cases, not a CAD installation. A live result for one metric macro or STEP export on an isolated vendor copy is scoped to that exact macro, fixture and installed revision; it does not certify the 24-task benchmark, native form/class projects or every supported version. See [compatibility](version-compatibility.md).
