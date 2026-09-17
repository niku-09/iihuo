'==============================================================
' make_load_combos_v1.bas
'--------------------------------------------------------------
' Creates the 24 Nastran LOAD Combination sets for the SCR
' upper-structure model.
'
' LOAD ID MAPPING (this model's own serial numbers)
'   1  Ec   E-dead load
'   2  Es   Structural Load (Self weight)
'   3  Mr   material load - reclaiming  (also used for "Ms")
'   4  Te
'   5  Ts
'   6  Tr
'   7  Dg
'   8  Sc
'   9  Aa
'   10 Bbs  (also used for "Bbr")
'   11 Cs   (also used for "Cc")
'   12 Egx  (also used for "Eqx")
'   13 Egy  (also used for "Eqy")
'   14 Inx
'   15 Iny
'   16 Wx
'   17 Wy
'
' NAMING
'   Titles are kept EXACTLY as the reference document, including
'   Ms / Cc / Bbr / Eqx / Eqy. Those are stacking-vs-reclaiming
'   names for the same underlying load, so the title differs but
'   the referenced load set is the same one.
'
' FACTORS
'   "+" in the title = +1.0, "-" = -1.0. Overall scale 1.0.
'
' RUN_MODE
'   0 = report only, lists what would be created
'   1 = create the FIRST combination only, then stop
'   2 = create all 24
'
' >>> SAVE A BACKUP BEFORE RUN_MODE 1 OR 2 <<<
'
' API STATUS
'   The Nastran LOAD Combination call is UNVERIFIED on this
'   build. CreateCombo() below holds the attempts, isolated so
'   it can be replaced with a recorded call without touching the
'   combination table. Every creation is read back; three
'   consecutive failures aborts.
'==============================================================

Const RUN_MODE    As Integer = 0
Const BACKUP_DONE As Integer = 0

Const OUT_DIR     As String = "C:\Users\nikhil.mohan\Desktop\Analysis testing\audit\"

' First load set ID for the new combinations. Must be free.
Const FIRST_ID    As Long = 101

Const MAX_FAILS   As Integer = 3

Dim femap As Object
Dim logF As Integer
Dim failCount As Integer

' combination table
Dim cTitle(0 To 23) As String
Dim cIDs(0 To 23, 0 To 11) As Long
Dim cFac(0 To 23, 0 To 11) As Double
Dim cN(0 To 23) As Integer
Dim nC As Integer

'==============================================================
Sub Main

    Set femap = GetObject(, "femap.model")

    If EnsureDir(OUT_DIR) = 0 Then
        femap.feAppMessage 2, "Cannot create " & OUT_DIR
        End
    End If

    If RUN_MODE > 0 Then
        If BACKUP_DONE <> 1 Then
            femap.feAppMessage 2, "REFUSING TO WRITE. Save a backup, then set BACKUP_DONE = 1."
            End
        End If
    End If

    failCount = 0

    logF = FreeFile
    Open OUT_DIR & "load_combos.log" For Output As #logF

    WL "=== make_load_combos_v1 ==="
    WL "RUN_MODE = " & RUN_MODE
    WL ""

    Call BuildTable
    Call CheckBaseLoads
    Call CheckIdFree
    Call Report
    Call CreateAll

    Close #logF
    femap.feAppMessage 0, "Done. See " & OUT_DIR & "load_combos.log"

End Sub

'--------------------------------------------------------------
' THE COMBINATION TABLE
'--------------------------------------------------------------
Sub BuildTable

    nC = 0

    Call AddC("H-case 1 = Es + Ec + Te - Inx")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(4, 1)
    Call Term(14, -1)

    Call AddC("H-case 2 = Es+Ec+Te+Inx")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(4, 1)
    Call Term(14, 1)

    Call AddC("H-case 3 = Es+Ec+Mr+Ts+Inx")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(3, 1)
    Call Term(5, 1)
    Call Term(14, 1)

    Call AddC("H-case 4 = Es+Ec+Mr+Tr+Dg-Sc+Inx")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(3, 1)
    Call Term(6, 1)
    Call Term(7, 1)
    Call Term(8, -1)
    Call Term(14, 1)

    Call AddC("H-case 5 = Es+Ec+Ms+Ts+Iny")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(3, 1)
    Call Term(5, 1)
    Call Term(15, 1)

    Call AddC("H-case 6 = Es+Ec+Mr+Tr+Dg-Sc+Iny")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(3, 1)
    Call Term(6, 1)
    Call Term(7, 1)
    Call Term(8, -1)
    Call Term(15, 1)

    Call AddC("Hz-case 7 = Es+Ec+Te-Inx-Wx")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(4, 1)
    Call Term(14, -1)
    Call Term(16, -1)

    Call AddC("Hz-case 8 = Es+Ec+Te+Iny+Wy")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(4, 1)
    Call Term(15, 1)
    Call Term(17, 1)

    Call AddC("Hz-case 9 = Es+Ec+Ms+Ts+Inx+Wx")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(3, 1)
    Call Term(5, 1)
    Call Term(14, 1)
    Call Term(16, 1)

    Call AddC("Hz-case 10 = Es+Ec+Mr+Tr+Dg-Sc+Inx+Wx")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(3, 1)
    Call Term(6, 1)
    Call Term(7, 1)
    Call Term(8, -1)
    Call Term(14, 1)
    Call Term(16, 1)

    Call AddC("Hz-case 11 = Es+Ec+Ms+Ts+Iny+Wy")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(3, 1)
    Call Term(5, 1)
    Call Term(15, 1)
    Call Term(17, 1)

    Call AddC("Hz-case 12 = Es+Ec+Mr+Tr+Dg-Sc+Iny+Wy")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(3, 1)
    Call Term(6, 1)
    Call Term(7, 1)
    Call Term(8, -1)
    Call Term(15, 1)
    Call Term(17, 1)

    Call AddC("Hzs-case 13 = Es+Ec+Ms+Ts+Inx+Wx+Cs")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(3, 1)
    Call Term(5, 1)
    Call Term(14, 1)
    Call Term(16, 1)
    Call Term(11, 1)

    Call AddC("Hzs-case 14 = Es+Ec+Mr+Tr+Dg-Sc+Inx+Wx+Cc")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(3, 1)
    Call Term(6, 1)
    Call Term(7, 1)
    Call Term(8, -1)
    Call Term(14, 1)
    Call Term(16, 1)
    Call Term(11, 1)

    Call AddC("Hzs-case 15 = Es+Ec+Te-Inx-Wx-Eqx")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(4, 1)
    Call Term(14, -1)
    Call Term(16, -1)
    Call Term(12, -1)

    Call AddC("Hzs-case 16 = Es+Ec+Te+Iny+Wy+Eqy")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(4, 1)
    Call Term(15, 1)
    Call Term(17, 1)
    Call Term(13, 1)

    Call AddC("Hzs-case 17 = Es+Ec+Ms+Ts+Inx+Wx+Eqx")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(3, 1)
    Call Term(5, 1)
    Call Term(14, 1)
    Call Term(16, 1)
    Call Term(12, 1)

    Call AddC("Hzs-case 18 = Es+Ec+Mr+Tr+Dg-Sc+Inx+Wx+Eqx")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(3, 1)
    Call Term(6, 1)
    Call Term(7, 1)
    Call Term(8, -1)
    Call Term(14, 1)
    Call Term(16, 1)
    Call Term(12, 1)

    Call AddC("Hzs-case 19 = Es+Ec+Ms+Ts+Iny+Wy+Eqy")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(3, 1)
    Call Term(5, 1)
    Call Term(15, 1)
    Call Term(17, 1)
    Call Term(13, 1)

    Call AddC("Hzs-case 20 = Es+Ec+Mr+Tr+Dg-Sc+Iny+Wy+Eqy")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(3, 1)
    Call Term(6, 1)
    Call Term(7, 1)
    Call Term(8, -1)
    Call Term(15, 1)
    Call Term(17, 1)
    Call Term(13, 1)

    Call AddC("Hzs-case 21 = Es+Ec+Te-Inx-Wx+Aa")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(4, 1)
    Call Term(14, -1)
    Call Term(16, -1)
    Call Term(9, 1)

    Call AddC("Hzs-case 22 = Es+Ec+Ms+Ts-Inx+Wx+Bbs")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(3, 1)
    Call Term(5, 1)
    Call Term(14, -1)
    Call Term(16, 1)
    Call Term(10, 1)

    Call AddC("Hzs-case 23 = Es+Ec+Mr+Tr+Dg-Sc+Inx+Wx+Bbr")
    Call Term(2, 1)
    Call Term(1, 1)
    Call Term(3, 1)
    Call Term(6, 1)
    Call Term(7, 1)
    Call Term(8, -1)
    Call Term(14, 1)
    Call Term(16, 1)
    Call Term(10, 1)

    Call AddC("h-1Es+Ec")
    Call Term(2, 1)
    Call Term(1, 1)

End Sub

Sub AddC(t As String)
    cTitle(nC) = t
    cN(nC) = 0
    nC = nC + 1
End Sub

Sub Term(lsID As Long, fac As Double)
    Dim i As Integer
    i = nC - 1
    cIDs(i, cN(i)) = lsID
    cFac(i, cN(i)) = fac
    cN(i) = cN(i) + 1
End Sub

'--------------------------------------------------------------
' Confirm every referenced base load set actually exists.
' A missing one would silently produce a combination short of a
' load, which raises no error and is very hard to spot later.
'--------------------------------------------------------------
Sub CheckBaseLoads

    Dim oLS As Object
    Dim i As Integer
    Dim j As Integer
    Dim missing As Integer
    Dim seen(0 To 200) As Integer
    Dim k As Integer

    WL "--- base load set check ---"
    missing = 0

    For k = 0 To 200
        seen(k) = 0
    Next k

    For i = 0 To nC - 1
        For j = 0 To cN(i) - 1
            If cIDs(i, j) <= 200 Then
                seen(cIDs(i, j)) = 1
            End If
        Next j
    Next i

    For k = 1 To 200
        If seen(k) = 1 Then
            Set oLS = femap.feLoadSet                     'VERIFIED (no args)
            If oLS.Get(k) <> -1 Then
                WL "  MISSING base load set " & k
                missing = missing + 1
            End If
        End If
    Next k

    If missing = 0 Then
        WL "  all referenced base load sets exist"
    Else
        WL "  *** " & missing & " referenced load sets do not exist."
        WL "  *** Fix the mapping before creating combinations."
    End If
    WL ""

End Sub

'--------------------------------------------------------------
Sub CheckIdFree

    Dim oLS As Object
    Dim i As Integer
    Dim clash As Integer

    clash = 0

    For i = 0 To nC - 1
        Set oLS = femap.feLoadSet
        If oLS.Get(FIRST_ID + i) = -1 Then
            clash = clash + 1
        End If
    Next i

    WL "--- id range check ---"
    WL "  " & FIRST_ID & " to " & (FIRST_ID + nC - 1)

    If clash > 0 Then
        WL "  *** " & clash & " of those load set IDs already exist."
        WL "  *** Raise FIRST_ID and re-run. Nothing was created."
        WL ""
        Close #logF
        femap.feAppMessage 2, "FIRST_ID range collides with existing load sets."
        End
    End If

    WL "  range is free"
    WL ""

End Sub

'--------------------------------------------------------------
Sub Report

    Dim i As Integer
    Dim j As Integer
    Dim s As String

    WL "--- combinations to create: " & nC & " ---"

    For i = 0 To nC - 1
        s = ""
        For j = 0 To cN(i) - 1
            If cFac(i, j) > 0 Then
                s = s & " +" & cIDs(i, j)
            Else
                s = s & " -" & cIDs(i, j)
            End If
        Next j
        WL "  " & (FIRST_ID + i) & "  " & cTitle(i)
        WL "        sets:" & s
    Next i

    WL ""

End Sub

'--------------------------------------------------------------
Sub CreateAll

    Dim i As Integer
    Dim made As Integer
    Dim bad As Integer
    Dim stopNow As Integer

    WL "--- creating ---"

    If RUN_MODE = 0 Then
        WL "  [report only] nothing created"
        WL ""
        WL "  Next: back up, set BACKUP_DONE = 1 and RUN_MODE = 1."
        WL "  That creates ONE combination so you can open it in"
        WL "  Model > Load > Combine and confirm the set type is"
        WL "  Nastran LOAD Combination and the factors are right."
        Exit Sub
    End If

    made = 0
    bad = 0
    stopNow = 0

    For i = 0 To nC - 1

        If stopNow = 0 Then

            If CreateCombo(FIRST_ID + i, i) = 1 Then
                made = made + 1
                If RUN_MODE = 1 Then
                    WL "  [RUN_MODE 1] created set " & (FIRST_ID + i) & " only."
                    WL "  Open it and verify before RUN_MODE 2."
                    stopNow = 1
                End If
            Else
                bad = bad + 1
                If failCount >= MAX_FAILS Then
                    stopNow = 1
                End If
            End If

        End If

    Next i

    WL "  created : " & made
    WL "  failed  : " & bad
    WL ""

End Sub

'--------------------------------------------------------------
' THE ONLY UNVERIFIED PART.
'
' Creates one Nastran LOAD Combination load set. If this fails,
' record a manual combination (Tools > Programming > Record,
' Model > Load > Combine, Set Type = Nastran LOAD Combination)
' and replace the body of this function with the recorded call.
' The combination table above does not need to change.
'--------------------------------------------------------------
Function CreateCombo(newID As Long, idx As Integer) As Integer

    Dim oLS As Object
    Dim vID As Variant
    Dim vSc As Variant
    Dim j As Integer
    Dim wrote As Integer

    CreateCombo = 0
    wrote = 0

    On Error GoTo BailCombo

    ReDim vID(0 To cN(idx) - 1)
    ReDim vSc(0 To cN(idx) - 1)
    For j = 0 To cN(idx) - 1
        vID(j) = cIDs(idx, j)
        vSc(j) = cFac(idx, j)
    Next j

    Set oLS = femap.feLoadSet                             'VERIFIED
    oLS.title = cTitle(idx)                               'UNVERIFIED as a write

    On Error Resume Next

    ' attempt 1 - set the combination through the LoadSet object
    oLS.SetCombination 1, 1#, cN(idx), vID, vSc           'UNVERIFIED
    If Err = 0 Then
        wrote = 1
    End If
    Err = 0

    ' attempt 2 - separate property style
    If wrote = 0 Then
        oLS.CombinationType = 1
        oLS.OverallScale = 1#
        oLS.vCombinationID = vID
        oLS.vCombinationScale = vSc
        If Err = 0 Then
            wrote = 1
        End If
        Err = 0
    End If

    oLS.Put newID
    Err = 0
    On Error GoTo BailCombo

    ' --- read back ---
    Set oLS = femap.feLoadSet
    If oLS.Get(newID) <> -1 Then
        Call NoteFail(newID, "load set was not created")
        Exit Function
    End If

    If wrote = 0 Then
        Call NoteFail(newID, "set created but combination data was NOT written - record a macro")
        Exit Function
    End If

    CreateCombo = 1
    failCount = 0
    Exit Function

BailCombo:
    Call NoteFail(newID, "runtime error")

End Function

Sub NoteFail(lsID As Long, reason As String)

    failCount = failCount + 1
    WL "  FAILED set " & lsID & " : " & reason

    If failCount >= MAX_FAILS Then
        WL ""
        WL "  *** ABORTING after " & MAX_FAILS & " consecutive failures."
        WL "  *** The Nastran LOAD Combination API call in"
        WL "  *** CreateCombo() is wrong for this build."
        WL "  ***"
        WL "  *** Tools > Programming > Record, then create ONE"
        WL "  *** combination by hand via Model > Load > Combine"
        WL "  *** with Set Type = Nastran LOAD Combination."
        WL "  *** Send the recorded macro - only CreateCombo()"
        WL "  *** needs replacing, the table above stays as is."
        WL ""
    End If

End Sub

'--------------------------------------------------------------
Function EnsureDir(p As String) As Integer

    Dim i As Integer
    Dim c As String
    Dim partial As String

    EnsureDir = 0
    partial = ""

    For i = 1 To Len(p)
        c = Mid(p, i, 1)
        partial = partial & c
        If c = "\" Then
            If Len(partial) > 3 Then
                On Error Resume Next
                MkDir partial
                On Error GoTo 0
            End If
        End If
    Next i

    On Error Resume Next
    If Dir(p, 16) <> "" Then
        EnsureDir = 1
    End If
    On Error GoTo 0

End Function

Sub WL(s As String)
    Print #logF, s
End Sub
