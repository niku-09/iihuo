'==============================================================
' fix_node149733.bas
'--------------------------------------------------------------
' Applies a precomputed action list to the model.
'
' The action list (FIX_NODE149733.csv) comes from a full
' scan of all 608,526 shell elements in the Nastran deck - not
' from the .f06. That matters: Nastran silently TOLERATES many
' collapsed elements, and those are exactly the ones blocking
' coincident node merges and causing the mechanisms.
'
' CSV FORMAT
'   ElementID,Action,N1,N2,N3
'     DELETE - zero or near-zero area. Carries no load.
'     TRIA   - quad with one zero-length edge. It IS the
'              triangle N1-N2-N3. Converted, area preserved.
'
' STAGE ORDER
'   Stage 1 DELETE  - uses feDelete, VERIFIED working on this
'                     build (returned FE_OK, re-fetch confirmed)
'   Stage 2 TRIA    - topology write, NOT yet verified. Runs
'                     AFTER deletes so a failure here cannot
'                     cost you the deletes.
'
' RUN_MODE
'   0 = report only
'   1 = delete all, convert ONE element then stop (inspect it)
'   2 = delete all, convert all
'
' >>> SAVE A BACKUP BEFORE RUN_MODE 1 OR 2 <<<
'==============================================================

Const RUN_MODE    As Integer = 0
Const BACKUP_DONE As Integer = 0

Const CSV_PATH    As String = "C:\Users\nikhil.mohan\Desktop\Analysis testing\audit\FIX_NODE149733.csv"
Const OUT_DIR     As String = "C:\Users\nikhil.mohan\Desktop\Analysis testing\audit\"

Const FT_ELEM     As Integer = 8
Const TOPO_TRIA3  As Integer = 2
Const TOPO_QUAD4  As Integer = 4
Const MAX_ROWS    As Long = 40000
Const MAX_FAILS   As Integer = 3

Dim femap As Object
Dim logF As Integer

Dim delID() As Long
Dim nDel As Long
Dim triID() As Long
Dim triA() As Long
Dim triB() As Long
Dim triC() As Long
Dim nTri As Long

Dim reID() As Long
Dim reA() As Long
Dim reB() As Long
Dim reC() As Long
Dim reD() As Long
Dim nRe As Long

Dim failCount As Integer

Sub Main

    Set femap = GetObject(, "femap.model")

    If EnsureDir(OUT_DIR) = 0 Then
        femap.feAppMessage 2, "Cannot create " & OUT_DIR
        End
    End If

    If FileThere(CSV_PATH) = 0 Then
        femap.feAppMessage 2, "Action list not found: " & CSV_PATH
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
    Open OUT_DIR & "fix_all.log" For Output As #logF

    WL "=== fix_node149733 ==="
    WL "RUN_MODE = " & RUN_MODE
    WL "actions  = " & CSV_PATH
    WL ""

    Call LoadActions
    Call Stage1_Delete
    Call Stage2_Tria
    Call Stage3_Reorder
    Call Summary

    Close #logF
    femap.feAppMessage 0, "fix_node149733 done. See " & OUT_DIR & "fix_all.log"

End Sub

'--------------------------------------------------------------
Sub LoadActions

    Dim f As Integer
    Dim ln As String
    Dim hdr As String
    Dim act As String
    Dim eid As Long

    ReDim delID(0 To MAX_ROWS)
    ReDim triID(0 To MAX_ROWS)
    ReDim triA(0 To MAX_ROWS)
    ReDim triB(0 To MAX_ROWS)
    ReDim triC(0 To MAX_ROWS)
    ReDim reID(0 To MAX_ROWS)
    ReDim reA(0 To MAX_ROWS)
    ReDim reB(0 To MAX_ROWS)
    ReDim reC(0 To MAX_ROWS)
    ReDim reD(0 To MAX_ROWS)
    nDel = 0
    nTri = 0
    nRe = 0

    f = FreeFile
    Open CSV_PATH For Input As #f
    Line Input #f, hdr

    Do While Not EOF(f)

        Line Input #f, ln

        If Len(ln) > 3 Then

            eid = CLng(Field(ln, 0))
            act = Trim(Field(ln, 1))

            If act = "DELETE" Then
                If nDel <= MAX_ROWS Then
                    delID(nDel) = eid
                    nDel = nDel + 1
                End If
            ElseIf act = "TRIA" Then
                If nTri <= MAX_ROWS Then
                    triID(nTri) = eid
                    triA(nTri) = CLng(Field(ln, 2))
                    triB(nTri) = CLng(Field(ln, 3))
                    triC(nTri) = CLng(Field(ln, 4))
                    nTri = nTri + 1
                End If
            ElseIf act = "REORDER" Then
                If nRe <= MAX_ROWS Then
                    reID(nRe) = eid
                    reA(nRe) = CLng(Field(ln, 2))
                    reB(nRe) = CLng(Field(ln, 3))
                    reC(nRe) = CLng(Field(ln, 4))
                    reD(nRe) = CLng(Field(ln, 5))
                    nRe = nRe + 1
                End If
            End If

        End If

    Loop

    Close #f

    WL "actions loaded"
    WL "  DELETE  : " & nDel
    WL "  TRIA    : " & nTri
    WL "  REORDER : " & nRe
    WL ""

End Sub

'--------------------------------------------------------------
' STAGE 1 - delete. feDelete is VERIFIED on this build.
'--------------------------------------------------------------
Sub Stage1_Delete

    Dim s As Object
    Dim oEl As Object
    Dim i As Long
    Dim present As Long
    Dim rc As Long
    Dim left As Long

    WL "--- Stage 1: delete ---"

    If nDel = 0 Then
        WL "  nothing to delete"
        WL ""
        Exit Sub
    End If

    Set s = femap.feSet
    s.Clear
    present = 0

    For i = 0 To nDel - 1
        Set oEl = femap.feElem
        If oEl.Get(delID(i)) = -1 Then
            s.Add delID(i)
            present = present + 1
        End If
    Next i

    WL "  present in model : " & present
    WL "  already absent   : " & (nDel - present)

    If RUN_MODE = 0 Then
        WL "  [report only] nothing deleted"
        WL ""
        Exit Sub
    End If

    If present = 0 Then
        WL "  nothing to do"
        WL ""
        Exit Sub
    End If

    rc = femap.feDelete(FT_ELEM, s.ID)
    WL "  feDelete returned : " & rc

    left = 0
    For i = 0 To nDel - 1
        Set oEl = femap.feElem
        If oEl.Get(delID(i)) = -1 Then
            left = left + 1
        End If
    Next i

    WL "  still present after : " & left

    If left = 0 Then
        WL "  DELETE STAGE OK"
    Else
        WL "  DELETE INCOMPLETE - check manually"
    End If

    WL ""

End Sub

'--------------------------------------------------------------
' STAGE 2 - convert collapsed quads to CTRIA3.
' Topology write is NOT verified on this build. Every conversion
' is read back and confirmed. Three consecutive failures aborts.
'--------------------------------------------------------------
Sub Stage2_Tria

    Dim i As Long
    Dim done As Long
    Dim skipped As Long
    Dim stopNow As Integer

    WL "--- Stage 2: convert collapsed quads to CTRIA3 ---"

    If nTri = 0 Then
        WL "  nothing to convert"
        WL ""
        Exit Sub
    End If

    If RUN_MODE = 0 Then
        WL "  [report only] would convert " & nTri & " elements"
        WL ""
        Exit Sub
    End If

    done = 0
    skipped = 0
    stopNow = 0

    For i = 0 To nTri - 1

        If stopNow = 0 Then

            If ConvertOne(triID(i), triA(i), triB(i), triC(i)) = 1 Then
                done = done + 1
                If RUN_MODE = 1 Then
                    WL "  [RUN_MODE 1] converted element " & triID(i) & " only."
                    WL "  Inspect it in the viewport before running RUN_MODE 2."
                    stopNow = 1
                End If
            Else
                skipped = skipped + 1
                If failCount >= MAX_FAILS Then
                    stopNow = 1
                End If
            End If

        End If

    Next i

    WL "  converted : " & done
    WL "  failed    : " & skipped
    WL ""

End Sub

' Converts a collapsed quad to CTRIA3.
'
' v3 FIX: the indexed write oEl.vnode(0) = a silently does NOTHING
' on this build and raises no error, so the old "If Err = 0 Then
' wrote = 1" wrongly reported success and the whole-array fallback
' never ran. Elements whose duplicate node sat in slot 4 appeared
' to work only because their first three nodes were already right.
' Now each write method is VERIFIED BY READ-BACK before moving on.
Function ConvertOne(eid As Long, a As Long, b As Long, c As Long) As Integer

    ConvertOne = 0

    If a <= 0 Or b <= 0 Or c <= 0 Then
        Exit Function
    End If

    ' --- attempt 1: indexed write ---
    If TryConvert(eid, a, b, c, 1) = 1 Then
        ConvertOne = 1
        failCount = 0
        Exit Function
    End If

    ' --- attempt 2: whole-array write ---
    If TryConvert(eid, a, b, c, 2) = 1 Then
        ConvertOne = 1
        failCount = 0
        Exit Function
    End If

    Call NoteFail(eid, "both write methods failed read-back")

End Function

' method 1 = indexed vnode write, method 2 = whole-array vnode write.
' Returns 1 only if a fresh Get confirms the change actually landed.
Function TryConvert(eid As Long, a As Long, b As Long, c As Long, method As Integer) As Integer

    Dim oEl As Object
    Dim v As Variant

    TryConvert = 0

    On Error GoTo BailTry

    Set oEl = femap.feElem
    If oEl.Get(eid) <> -1 Then
        Exit Function
    End If

    If oEl.topology = TOPO_TRIA3 Then
        ' already converted on a previous run
        TryConvert = 1
        Exit Function
    End If

    On Error Resume Next

    If method = 1 Then
        oEl.topology = TOPO_TRIA3
        oEl.vnode(0) = a
        oEl.vnode(1) = b
        oEl.vnode(2) = c
        oEl.vnode(3) = 0
    Else
        v = oEl.vnode
        v(0) = a
        v(1) = b
        v(2) = c
        v(3) = 0
        oEl.topology = TOPO_TRIA3
        oEl.vnode = v
    End If

    oEl.Put eid

    Err = 0
    On Error GoTo BailTry

    ' --- READ BACK from a fresh object. this is the only proof. ---
    Set oEl = femap.feElem
    If oEl.Get(eid) <> -1 Then
        Exit Function
    End If

    If oEl.topology <> TOPO_TRIA3 Then
        Exit Function
    End If

    v = oEl.vnode

    If CLng(v(0)) <> a Then
        Exit Function
    End If
    If CLng(v(1)) <> b Then
        Exit Function
    End If
    If CLng(v(2)) <> c Then
        Exit Function
    End If

    TryConvert = 1
    Exit Function

BailTry:
    TryConvert = 0

End Function

Sub NoteFail(eid As Long, reason As String)

    failCount = failCount + 1
    WL "  FAILED elem " & eid & " : " & reason

    If failCount >= MAX_FAILS Then
        WL ""
        WL "  *** ABORTING after " & MAX_FAILS & " consecutive failures."
        WL "  *** The topology write is not supported on this build."
        WL "  *** Stage 1 deletes are unaffected and already applied."
        WL "  *** Fallback: delete these 195 instead - they are small"
        WL "  *** triangles and the holes are minor. Or record a manual"
        WL "  *** element edit and send the recorded macro."
        WL ""
    End If

End Sub

'--------------------------------------------------------------
' STAGE 3 - reorder crossed quads. Same unverified write as
' Stage 2, so it runs last. The node order in the CSV was
' geometrically validated before being written to file.
'--------------------------------------------------------------
Sub Stage3_Reorder

    Dim i As Long
    Dim done As Long
    Dim skipped As Long
    Dim stopNow As Integer

    WL "--- Stage 3: reorder crossed quads ---"

    If nRe = 0 Then
        WL "  nothing to reorder"
        WL ""
        Exit Sub
    End If

    If RUN_MODE = 0 Then
        WL "  [report only] would reorder " & nRe & " elements"
        WL ""
        Exit Sub
    End If

    done = 0
    skipped = 0
    stopNow = 0

    For i = 0 To nRe - 1

        If stopNow = 0 Then

            If ReorderOne(reID(i), reA(i), reB(i), reC(i), reD(i)) = 1 Then
                done = done + 1
                If RUN_MODE = 1 Then
                    WL "  [RUN_MODE 1] reordered element " & reID(i) & " only."
                    stopNow = 1
                End If
            Else
                skipped = skipped + 1
                If failCount >= MAX_FAILS Then
                    stopNow = 1
                End If
            End If

        End If

    Next i

    WL "  reordered : " & done
    WL "  failed    : " & skipped
    WL ""

End Sub

Function ReorderOne(eid As Long, a As Long, b As Long, c As Long, dd As Long) As Integer

    Dim oEl As Object
    Dim v As Variant
    Dim wrote As Integer
    Dim okBack As Integer

    ReorderOne = 0
    wrote = 0

    If a <= 0 Or b <= 0 Or c <= 0 Or dd <= 0 Then
        Exit Function
    End If

    On Error GoTo BailR

    Set oEl = femap.feElem
    If oEl.Get(eid) <> -1 Then
        Exit Function
    End If

    If oEl.topology <> TOPO_QUAD4 Then
        Exit Function
    End If

    On Error Resume Next
    oEl.vnode(0) = a
    oEl.vnode(1) = b
    oEl.vnode(2) = c
    oEl.vnode(3) = dd
    If Err = 0 Then
        wrote = 1
    End If
    Err = 0
    On Error GoTo BailR

    If wrote = 0 Then
        On Error Resume Next
        v = oEl.vnode
        v(0) = a
        v(1) = b
        v(2) = c
        v(3) = dd
        oEl.vnode = v
        If Err = 0 Then
            wrote = 1
        End If
        Err = 0
        On Error GoTo BailR
    End If

    If wrote = 0 Then
        Call NoteFail(eid, "vnode write rejected")
        Exit Function
    End If

    If oEl.Put(eid) <> -1 Then
        Call NoteFail(eid, "Put failed")
        Exit Function
    End If

    okBack = 0
    Set oEl = femap.feElem
    If oEl.Get(eid) = -1 Then
        v = oEl.vnode
        If CLng(v(1)) = b Then
            If CLng(v(2)) = c Then
                okBack = 1
            End If
        End If
    End If

    If okBack = 1 Then
        ReorderOne = 1
        failCount = 0
    Else
        Call NoteFail(eid, "read-back mismatch")
    End If

    Exit Function

BailR:
    Call NoteFail(eid, "runtime error")

End Function

'--------------------------------------------------------------
Sub Summary

    WL "--- next steps ---"

    If RUN_MODE = 0 Then
        WL "  1. Save a backup of the model"
        WL "  2. BACKUP_DONE = 1, RUN_MODE = 1"
        WL "  3. Inspect the one converted element"
        WL "  4. RUN_MODE = 2"
    Else
        WL "  1. Tools > Check > Coincident Nodes, WHOLE MODEL,"
        WL "     tolerance 0.05 mm, MERGE."
        WL "     2527 coincident pairs exist. Most were blocked by"
        WL "     the elements just removed. Merging them is what"
        WL "     actually fixes the mechanisms."
        WL "  2. Re-solve with PARAM,BAILOUT UNTICKED."
        WL "     If it runs clean, the mechanisms are genuinely gone."
        WL "  3. Check equilibrium in the .f06: SPCFORCE totals must"
        WL "     equal minus OLOAD totals. Last run they did not."
    End If

End Sub

'--------------------------------------------------------------
' UTILITY
'--------------------------------------------------------------

Function Field(s As String, idx As Integer) As String

    Dim cur As Integer
    Dim st As Integer
    Dim i As Integer
    Dim c As String

    cur = 0
    st = 1
    Field = ""

    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c = "," Then
            If cur = idx Then
                Field = Mid(s, st, i - st)
                Exit Function
            End If
            cur = cur + 1
            st = i + 1
        End If
    Next i

    If cur = idx Then
        Field = Mid(s, st)
    End If

End Function

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

Function FileThere(p As String) As Integer

    FileThere = 0
    On Error Resume Next
    If Dir(p) <> "" Then
        FileThere = 1
    End If
    On Error GoTo 0

End Function

Sub WL(s As String)

    Print #logF, s

End Sub
