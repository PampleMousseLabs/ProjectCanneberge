Attribute VB_Name = "modDashBranching"
'==============================================================================
' modDashBranching
' Project Canneberge — Business Enterprise Valuation Model
' PampleMousse Labs
'
' Purpose: Handles Public/Private branching logic for Dash_Prjctn projection
'          rows that require VBA (cannot be resolved by formula alone).
'
'          Pattern D — Revenue Growth %   (J11, K11, L11)
'          Pattern E — EBITDA Improvement % (J19, K19, L19)
'
' Trigger: Called from Sheet1 (Control) Worksheet_Change event when:
'            - CompanyStatus changes
'            - Control!AB9  changes (Revenue trigger for Pattern D)
'            - Control!AB20 changes (EBITDA trigger for Pattern E)
'
' Design:
'   CompanyStatus = "Publicly Traded" ? Calc mode (formula, slate text, no fill)
'   CompanyStatus = "Private"
'       + trigger cell empty/zero   ? Input mode (lavender, blank, user types %)
'       + trigger cell has value    ? Calc mode
'
' Calc formulas:
'   Pattern D  J11 = J10/I10-1    K11 = K10/J10-1    L11 = L10/K10-1
'   Pattern E  J19 = (J17/J10)-(I17/I10)
'              K19 = (K17/K10)-(J17/J10)
'              L19 = (L17/L10)-(K17/K10)
'
' Formatting constants (both patterns, both modes):
'   Number format : 0.0%
'   Font face     : Segoe UI
'   Font size     : 10
'   Input  font   : RGB(53,  9, 185)   fill RGB(237, 231, 254)
'   Calc   font   : RGB(51, 65,  85)   fill xlNone
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Public entry point — called by Sheet1 Worksheet_Change
' Evaluates which pattern(s) need refreshing based on the changed cell address,
' then delegates to ApplyDashBranchPattern for the actual work.
'------------------------------------------------------------------------------
Public Sub RefreshDashBranching(ByVal changedCell As Range)

    Dim wsControl As Worksheet
    Dim wsDash As Worksheet
    Dim statusCell As Range
    Dim isPublic As Boolean

    Dim refreshD As Boolean
    Dim refreshE As Boolean

    Set wsControl = ThisWorkbook.Worksheets("Control")
    Set wsDash = ThisWorkbook.Worksheets("Dash_Prjctn")

    On Error GoTo ErrHandler
    Set statusCell = ThisWorkbook.Names("CompanyStatus").RefersToRange
    isPublic = (Trim(CStr(statusCell.Value)) = "Publicly Traded")

        If Not Intersect(changedCell, statusCell) Is Nothing Then
        refreshD = True
        refreshE = True

    ElseIf Not Intersect(changedCell, wsControl.Range("AB9:AD9")) Is Nothing Then
        refreshD = True

    ElseIf Not Intersect(changedCell, wsControl.Range("AB16:AD18")) Is Nothing Then
        refreshE = True

    Else
        Exit Sub
    End If

    ' Pattern D — Revenue Growth %
        If refreshD Then
        Dim trigD As Variant
        Dim targD As Variant
        Dim calcD As Variant

        trigD = Array("AB9", "AC9", "AD9")
        targD = Array("J11", "K11", "L11")
        calcD = Array( _
            "=J10/I10-1", _
            "=K10/J10-1", _
            "=L10/K10-1" _
        )

        ApplyDashBranchPattern wsControl, wsDash, isPublic, trigD, targD, calcD
    End If

    ' Pattern E — EBITDA Improvement %
    If refreshE Then
        Dim trigE As Variant
        Dim targE As Variant
        Dim calcE As Variant

        trigE = Array("AB16:AB18", "AC16:AC18", "AD16:AD18")
        targE = Array("J19", "K19", "L19")
        calcE = Array( _
            "=(J17/J10)-(I17/I10)", _
            "=(K17/K10)-(J17/J10)", _
            "=(L17/L10)-(K17/K10)" _
        )

        ApplyDashBranchPattern wsControl, wsDash, isPublic, trigE, targE, calcE
    End If

    Exit Sub

ErrHandler:
    Debug.Print "modDashBranching.RefreshDashBranching ERROR " & Err.Number & ": " & Err.Description
End Sub


'------------------------------------------------------------------------------
' ApplyDashBranchPattern
'
' Shared parameterized routine (DRY) — identical logic for D and E.
'
' Parameters:
'   wsDash        Dash_Prjctn worksheet object
'   isPublic      True = "Publicly Traded"; False = "Private"
'   triggerCell   Control sheet cell whose value gates Input vs Calc for Private
'   targetAddrs   1-based array(1..3) of cell address strings on wsDash
'   calcFormulas  1-based array(1..3) of formula strings (A1 notation, wsDash-relative)
'------------------------------------------------------------------------------
Private Sub ApplyDashBranchPattern( _
        ByVal wsControl As Worksheet, _
        ByVal wsDash As Worksheet, _
        ByVal isPublic As Boolean, _
        ByVal triggerRefs As Variant, _
        ByVal targetAddrs As Variant, _
        ByVal calcFormulas As Variant)

    Dim i As Long
    Dim useCalcMode As Boolean
    Dim triggerRange As Range
    Dim cel As Range

    For i = LBound(targetAddrs) To UBound(targetAddrs)

        Set triggerRange = wsControl.Range(CStr(triggerRefs(i)))
        Set cel = wsDash.Range(CStr(targetAddrs(i)))

        If isPublic Then
            useCalcMode = True
        Else
            useCalcMode = Not IsTriggerEffectivelyEmpty(triggerRange)
        End If

        If useCalcMode Then
            ApplyCalcMode cel, CStr(calcFormulas(i))
        Else
            ApplyInputMode cel
        End If
    Next i

End Sub

Private Function IsTriggerEffectivelyEmpty(ByVal rng As Range) As Boolean

    If rng.Cells.CountLarge = 1 Then
        IsTriggerEffectivelyEmpty = IsEmpty(rng.Value) _
            Or rng.Value = "" _
            Or rng.Value = 0
    Else
        IsTriggerEffectivelyEmpty = (Application.Sum(rng) = 0)
    End If

End Function

'------------------------------------------------------------------------------
' ApplyCalcMode
' Writes formula and applies slate-on-clear formatting.
'------------------------------------------------------------------------------
Private Sub ApplyCalcMode(ByVal cel As Range, ByVal formula As String)

    With cel
        .formula = formula
        .NumberFormat = "0.0%"
        With .Font
            .Name = "Segoe UI"
            .Size = 10
            .Color = RGB(51, 65, 85)
        End With
        .Interior.ColorIndex = xlNone          ' clear any fill
    End With

End Sub


'------------------------------------------------------------------------------
' ApplyInputMode
' Clears formula, blanks value, applies lavender input formatting.
'------------------------------------------------------------------------------
Private Sub ApplyInputMode(ByVal cel As Range)

    With cel
        .ClearContents                         ' remove formula or value
        .NumberFormat = "0.0%"
        With .Font
            .Name = "Segoe UI"
            .Size = 10
            .Color = RGB(53, 9, 185)
        End With
        .Interior.Color = RGB(237, 231, 254)   ' lavender
    End With

End Sub


'------------------------------------------------------------------------------
' IsEffectivelyEmpty
' Returns True when a cell is blank, empty string, or numeric zero.
' Mirrors the spec definition used across the project.
'------------------------------------------------------------------------------
Private Function IsEffectivelyEmpty(ByVal cel As Range) As Boolean
    IsEffectivelyEmpty = IsEmpty(cel) Or cel.Value = "" Or cel.Value = 0
End Function

