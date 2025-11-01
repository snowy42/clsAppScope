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
Application.Calculation = xlCalculationManual
' ... code ...
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
```

This is:
- Verbose 
- Easy to forget 
- Unsafe if the macro errors 
- Doesn't remember previous state (if things were already off, you accidentally turn them back on later) 

This class uses RAII (Resource Acquisition Is Initialization) technique: 
- When the scope object is created → settings are suspended 
- When it goes out of scope → settings are restored automatically

---

## 🎯 Usage Example
```vb
Public Sub Demo()
    With AppScopeF(sEvents + sScreen + sCalc + sStatus, status:="Processing...")
        ' Your code here
        Range("A1").Value = "Updated without flicker"
    End With
End Sub
```
No manual cleanup. No risk of leaving Excel in a broken state.

---

## 🔧 Flags Reference
| Flag     | Meaning                                      |
|---------|-----------------------------------------------|
| `sEvents` | Disable event triggers during scope          |
| `sScreen` | Disable screen updating (prevents flicker)   |
| `sAlerts` | Suppress confirmation alerts                 |
| `sCalc`   | Set calculation mode to Manual for performance |
| `sStatus` | Display custom status bar text               |
| `sAll`    | Apply all options                            |

Combine flags using `+` or `Or`:

```vb
With AppScopeF(sEvents + sScreen + sCalc) ' can also use: sEvents Or sScreen Or sCalc
    ' ...
End With
```

---

## 🧱 Installation
1. Import clsAppScope.cls into your VBA project
2. Import modAppScope.bas
3. Use AppScopeF(...) in your macros

---

## 🆘 Safety Reset
If you ever hit the VBA Reset button (the "stop" square), run:
```vb
AppRestoreDefaults
```
This restores Excel to normal behavior.

---

## 📄 License
This project is licensed under the MIT License.
See the header block in the source files for full license text.

---

## 👤 Author
Matthew Snow / Your VB Tutor
Excel/VBA + AI Automation Developer + Youtube VBA Tutor (youtube.com/@YourVBTutor)
