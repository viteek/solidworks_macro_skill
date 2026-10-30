# VSTA, C# / VB.NET, Python

Choose by need:
- **VBA macro** - default. Zero setup, runs inside SW, easy for designers to edit.
- **VSTA macro (.dll, VB.NET/C#)** - inside SW, .NET libraries (LINQ, JSON, Excel via EPPlus...). Edited in
  the VSTA IDE that ships with SW; SW 2018+ uses VSTA 3.0 (.NET Framework 4.x).
- **Standalone exe (C#)** - batch jobs, scheduled tasks, integration with ERP/DB; drives SW over COM.
- **Add-in (ISwAddin)** - permanent UI (toolbar, Task Pane, PropertyManager), event handlers
  (on save, on rebuild). Needs registration (regasm /codebase) - more admin overhead.
- **Python (pywin32)** - quick scripting for people who already use Python; out-param handling is awkward.

## C# essentials

References (from `<SW install>\api\redist\` or the SDK): `SolidWorks.Interop.sldworks.dll`,
`SolidWorks.Interop.swconst.dll` (and `swpublished` for add-ins). Set *Embed Interop Types = False*
for sldworks if you hit casting issues with arrays/events. Target .NET Framework 4.7.2/4.8 for add-ins;
standalone exe can be .NET 6+ (COM interop works on Windows) - but `Marshal.GetActiveObject` is not
available on .NET 5+, use the snippet below.

```csharp
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

static ISldWorks ConnectToSw()
{
    // Attach to running instance, or start a new one
    var progId = Type.GetTypeFromProgID("SldWorks.Application");
    // May attach to a running session OR start a new one; a new one is invisible, started with /embed
    // and loads NO add-ins. For a specific session use the ROT moniker "SolidWorks_PID_<pid>"
    // (advanced-topics.md §17) and wait for OnIdleNotify before using it.
    var app = (ISldWorks)Activator.CreateInstance(progId);
    app.Visible = true;
    return app;
}

int errors = 0, warnings = 0;
var model = (ModelDoc2)swApp.OpenDoc6(path, (int)swDocumentTypes_e.swDocPART,
    (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
if (model == null) throw new Exception($"Open failed, errors={errors}");

var cpm = model.Extension.CustomPropertyManager[""];      // indexed property in C#
cpm.Add3("Material", (int)swCustomInfoType_e.swCustomInfoText, "S235",
         (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
cpm.Get6("Material", false, out string val, out string resolved, out bool wasResolved, out bool linked);
```
- VBA properties with arguments become indexers / `get_X(...)` methods in C#
  (`CustomPropertyManager[""]`, `swDim.GetSystemValue3(...)`).
- Arrays come back as `object` → cast `(object[])` then each element to the interface.
- Enums need `(int)` casts.
- Release COM objects in long batches? Usually unnecessary; closing docs is what matters. Call
  `swApp.CloseDoc` and occasionally `GC.Collect()` in very long runs.
- Run SW work on a single STA thread; don't parallelize API calls.
- For a nicer add-in framework the community uses **xCAD.NET** (open source, by CodeStack) -
  mention as an option, not a requirement.

- Events: subscribe on the non-"I" types (`(SldWorks)app`, `(PartDoc)model`) - only those expose events.
  "I" members (`IActiveDoc2`) return strongly typed interfaces and need no cast.
- `MultiSelect2` with mixed object types: wrap each element in `new DispatchWrapper(obj)`.

## VSTA specifics
- Entry point: `SolidWorksMacro.Main()` in the generated project; `swApp` is a field already set.
- Keep the project's own namespace (Project Properties > Default namespace) and replace only the body of
  `public partial class SolidWorksMacro`; a mismatch gives "SolidWorksMacro doesn't contain a definition for swApp".
- Deploy the whole `bin\Release` folder (interop dlls included); unblock files downloaded from the internet.
- Macros with events or PropertyManager pages: untick "Stop VSTA debugger on macro exit" so the dll stays loaded.
- Macro is the compiled `.dll` in `SwMacro\bin\Debug|Release`; run via Tools > Macro > Run selecting the dll.
- References to extra NuGet packages are tricky in VSTA - prefer standalone exe if heavy deps are needed.

## Python (pywin32)

```python
import pythoncom
import win32com.client as win32

sw = win32.Dispatch("SldWorks.Application")
sw.Visible = True

# Out-parameters must be passed as VARIANT ByRef objects:
errors = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
warnings = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
model = sw.OpenDoc6(r"C:\path\part.SLDPRT", 1, 1, "", errors, warnings)   # swDocPART, Silent
if model is None:
    raise RuntimeError(f"Open failed: {errors.value}")

cpm = model.Extension.CustomPropertyManager("")
cpm.Add3("Material", 30, "S235", 2)      # swCustomInfoText, swCustomPropertyReplaceValue
```
- Use numeric enum values (comment the constant names).
- `Nothing` arguments: pass `win32.VARIANT(pythoncom.VT_DISPATCH, None)`.
- Late binding: some property-with-arguments calls behave differently; if a call fails try
  `win32.gencache.EnsureDispatch` or the `GetX/SetX` method form.
- Don't run from a background thread without `pythoncom.CoInitialize()`.
