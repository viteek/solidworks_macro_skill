# VBA and SOLIDWORKS COM semantics

Use `Option Explicit`, declared SOLIDWORKS interface types and explicit API return/error checks. `check_vba.py` and `check_api.py` catch useful subsets of mistakes; neither is the VBA compiler. Late-bound `Object`/`Variant`, runtime dispatch, callbacks, default members and context-dependent calls require full project compilation and runtime observation.

## COM Boolean results

Compare a SOLIDWORKS Boolean explicitly with `False`: use `= False` for failure and `<> False`
for nonzero success. Avoid `Not result`, `result = True` and implicit `If result Then` on an
unnormalized COM result. Normalize before using ordinary VBA Boolean expressions:

```vb
Dim saved As Boolean
saved = (swModel.Save3(swSaveAsOptions_Silent, errors, warnings) <> False)
If Not saved Then
    Err.Raise vbObjectError + 100, , "Save failed"
End If
```

This rule concerns actual Boolean results. Numeric status methods (`CustomPropertyManager.Add3`,
`Dimension.SetSystemValue3`, `Component2.SetSuppression2`) require their documented success enums.
`FileSystemObject.FileExists` and normalized comparisons produce standard VBA Booleans.
The [local probe](../assets/diagnostic-macros/boolean_probe.bas) checks numeric 0/1/-1 comparisons
and a real `GetSaveFlag` result on a disposable copy. Recorded R2026x execution observed raw API
truth as `1` and the explicit nonzero comparison as VBA `-1`; see [validation evidence](../validation/2026-11-01/README.md).
The linter tracks bounded local provenance, known numeric members and object aliases; it does
not model every function, ByRef output, branch or interprocedural transfer.

## Guard objects before dereferencing

VBA `And`/`Or` are not short-circuit guards. Do not write `If model Is Nothing Or model.GetType <> swDocPART Then`; the member access can still execute. Use separate guards:

```vb
If model Is Nothing Then
    Err.Raise vbObjectError + 100, , "Open a part copy"
End If
If model.GetType <> swDocPART Then
    Err.Raise vbObjectError + 101, , "Expected a part"
End If
```

Do not use `IIf` to protect an invalid call either: Microsoft documents that both result expressions are evaluated. Use `If ... Then ... Else`. See Microsoft's [And operator](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/and-operator) and [IIf function](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/iif-function).

## Types, arrays and ownership

- `Dim a, b As Long` declares `a` as Variant. Type every variable separately. VBA `Long` is the normal 32-bit integer used for many API enum/error arguments; 64-bit Windows pointer declarations use the appropriate `PtrSafe`/`LongPtr` declarations, not arbitrary substitution of API `Long` values.
- `ByRef` is the VBA default and requires the precise expected type. Declare error/status outputs with the documented type and let the API write them. Parenthesized coercion can create a temporary and lose output updates; do not use it to silence a signature mismatch. Microsoft's [ByRef guidance](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/byref-argument-type-mismatch) explains this distinction.
- COM enumeration methods commonly return Variant arrays or Empty. Check `IsEmpty` and `IsArray` separately before indexing. Use `LBound`/`UBound`, not an assumed zero lower bound; `IsArray` alone does not establish that a dynamic array has allocated bounds. Keep any bounds-error handler narrow and restore normal error handling immediately. See [IsArray](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/isarray-function).
- Use `Set` for object references. Check `Nothing` and document/interface type before casting. A valid dispatch object does not prove it is the intended model, component, feature or face. Reacquire invalidated entities after rebuild/configuration changes, using persistent references when appropriate.
- A part-context face and an assembly-context face can refer to the same geometry while requiring different API usage. Confirm selected object type, selection mark, component/drawing owner and context; do not mate unrelated part-context objects or treat a drawing component as a model component.

## Calls, units and state

Read the installed API's owner, argument direction, optional arguments, enum names and return value. `RunMacro2` has explicit module/procedure and error output semantics; [official help](https://help.solidworks.com/2020/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISldWorks~RunMacro2.html?verRedirect=1) describes the invocation contract. A Boolean result, an error enum and an output object are separate evidence. A non-null feature or a generated file alone is insufficient success validation.

Use documented system units for each member (commonly meters/radians for geometry); explicitly convert millimeters/degrees and give numeric assertions tolerances. Do not infer API units from the document display settings. Scope custom-property managers deliberately to document/configuration/cut-list, distinguish stored expressions from resolved values, and record whether cached data was requested.

Selection-based operations need the exact entity types and marks in the active document/configuration. Clear/restore selection deliberately, inspect API error codes and rebuild, then traverse the resulting objects. Capture relevant global preferences, active configuration/sheet and visibility state before changing them; restore in one cleanup path even on controlled failure. Keep `On Error Resume Next` local to a documented probe; preserve the original error before cleanup and never turn an unobserved failure into a success report.
