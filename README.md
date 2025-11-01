# VBA Application Scope Utility (RAII Style)
# clsAppScope

This utility provides a clean and *safe* way to temporarily suspend Excel application behaviors such as:

- Events (`Application.EnableEvents`)
- Screen updating (`Application.ScreenUpdating`)
- Alerts (`Application.DisplayAlerts`)
- Calculation mode (`Application.Calculation`)
- StatusBar messaging (`Application.StatusBar`)

and then **restore them automatically**, even if your code errors or exits early.

This solves a classic problem in VBA:
> If your macro crashes while events or screen updating are off, Excel stays in a broken state.

This class prevents that from happening.

---

## ✅ Why This Exists

Typical VBA code looks like this:

```vb
Application.ScreenUpdating = False
Application.EnableEvents = False
' ... code ...
Application.EnableEvents = True
Application.ScreenUpdating = True
