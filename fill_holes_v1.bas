'==============================================================
' fill_holes_v1.bas
'--------------------------------------------------------------
' Creates patch elements to close the holes left by deleting
' degenerate elements.
'
' HOW THE HOLES WERE FOUND
'   A hole shows up as a small CLOSED LOOP of free edges - edges
'   belonging to only one element. Real plate boundaries produce
'   long open runs of free edges; a 3- or 4-edge closed loop is
'   almost always a missing element.
'   305 such loops were found: 176 triangular, 129 quad.
'
' PROPERTY ASSIGNMENT
'   Each patch takes the property most common among the elements
'   already touching that hole's edges. Never defaulted.
'
' NEW ELEMENT IDs
'   Allocated from NEXT_ID upward. Set it above the highest
'   existing element ID in the model.
'
' RUN_MODE
'   0 = report only
'   1 = create ONE patch, then stop. Inspect it.
'   2 = create all
'
' >>> SAVE A BACKUP BEFORE RUN_MODE 1 OR 2 <<<
'==============================================================

Const RUN_MODE    As Integer = 0
Const BACKUP_DONE As Integer = 0

Const CSV_PATH    As String = "C:\Users\nikhil.mohan\Desktop\Analysis testing\audit\FILL_HOLES.csv"
Const OUT_DIR     As String = "C:\Users\nikhil.mohan\Desktop\Analysis testing\audit\"

' First element ID to use for the new patches. Must be above the
' highest element ID currently in the model.
Const NEXT_ID     As Long = 700000

Const ELTYPE_PLATE As Integer = 17
Const TOPO_TRIA3   As Integer = 2
Const TOPO_QUAD4   As Integer = 4
Const MAX_ROWS     As Long = 5000
Const MAX_FAILS    As Integer = 3

Dim femap As Object
Dim logF As Integer

Dim pAct() As String
Dim pProp() As Long
Dim pN1() As Long
Dim pN2() As Long
Dim pN3() As Long
Dim pN4() As Long
Dim nP As Long

Dim failCount As Integer

Sub Main

    Set femap = GetObject(, "femap.model")

    If EnsureDir(OUT_DIR) = 0 Then
        femap.feAppMessage 2, "Cannot create " & OUT_DIR
        End
    End If

    If FileThere(CSV_PATH) = 0 Then
        femap.feAppMessage 2, "Patch list not found: " & CSV_PATH
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
    Open OUT_DIR & "fill_holes.log" For Output As #logF

    WL "=== fill_holes_v1 ==="
    WL "RUN_MODE = " & RUN_MODE
    WL ""

    Call LoadPatches
    Call CheckIdFree
    Call CreatePatches

    Close #logF
    femap.feAppMessage 0, "fill_holes_v1 done. See " & OUT_DIR & "fill_holes.log"

End Sub

'--------------------------------------------------------------
Sub LoadPatches

    Dim f As Integer
    Dim ln As String
    Dim hdr As String
    Dim nTri As Long
    Dim nQuad As Long

    ReDim pAct(0 To MAX_ROWS)
    ReDim pProp(0 To MAX_ROWS)
    ReDim pN1(0 To MAX_ROWS)
    ReDim pN2(0 To MAX_ROWS)
    ReDim pN3(0 To MAX_ROWS)
    ReDim pN4(0 To MAX_ROWS)
    nP = 0
    nTri = 0
    nQuad = 0

    f = FreeFile
    Open CSV_PATH For Input As #f
    Line Input #f, hdr

    Do While Not EOF(f)

        Line Input #f, ln

        If Len(ln) > 5 Then
            If nP <= MAX_ROWS Then
                pAct(nP) = Trim(Field(ln, 0))
                pProp(nP) = CLng(Field(ln, 1))
                pN1(nP) = CLng(Field(ln, 2))
                pN2(nP) = CLng(Field(ln, 3))
                pN3(nP) = CLng(Field(ln, 4))
                pN4(nP) = CLng(Field(ln, 5))
                If pAct(nP) = "TRIA" Then
                    nTri = nTri + 1
                Else
                    nQuad = nQuad + 1
                End If
                nP = nP + 1
            End If
        End If

    Loop

    Close #f

    WL "patches loaded : " & nP
    WL "  triangular   : " & nTri
    WL "  quad         : " & nQuad
    WL ""

End Sub

'--------------------------------------------------------------
' Make sure the ID range we are about to use is actually empty.
' Creating on top of an existing element would destroy it.
'--------------------------------------------------------------
Sub CheckIdFree

    Dim oEl As Object
    Dim i As Long
    Dim clash As Long

    clash = 0

    For i = 0 To nP - 1
        Set oEl = femap.feElem
        If oEl.Get(NEXT_ID + i) = -1 Then
            clash = clash + 1
        End If
    Next i

    WL "id range check: " & NEXT_ID & " to " & (NEXT_ID + nP - 1)

    If clash > 0 Then
        WL "  *** " & clash & " of those IDs ALREADY EXIST."
        WL "  *** Raise NEXT_ID and re-run. Nothing was created."
        WL ""
        Close #logF
        femap.feAppMessage 2, "NEXT_ID range collides with existing elements. Raise NEXT_ID."
        End
    End If

    WL "  range is free"
    WL ""

End Sub

'--------------------------------------------------------------
Sub CreatePatches

    Dim i As Long
    Dim made As Long
    Dim bad As Long
    Dim stopNow As Integer

    WL "--- creating patches ---"

    If RUN_MODE = 0 Then
        WL "  [report only] would create " & nP & " elements"
        WL "  starting at element id " & NEXT_ID
        WL ""
        Exit Sub
    End If

    made = 0
    bad = 0
    stopNow = 0

    For i = 0 To nP - 1

        If stopNow = 0 Then

            If MakeOne(NEXT_ID + i, i) = 1 Then
                made = made + 1
                If RUN_MODE = 1 Then
                    WL "  [RUN_MODE 1] created element " & (NEXT_ID + i) & " only."
                    WL "  Inspect it in the viewport before RUN_MODE 2."
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

    If made > 0 Then
        WL "  Next: re-solve. The patches carry the same property as"
        WL "  the elements around each hole, so local stress there"
        WL "  should now be meaningful rather than a free-edge artefact."
    End If

End Sub

' Creates one patch element. Verified by read-back.
Function MakeOne(newID As Long, idx As Long) As Integer

    Dim oEl As Object
    Dim v As Variant
    Dim topo As Integer
    Dim wrote As Integer

    MakeOne = 0
    wrote = 0

    If pN1(idx) <= 0 Or pN2(idx) <= 0 Or pN3(idx) <= 0 Then
        Exit Function
    End If

    If pAct(idx) = "TRIA" Then
        topo = TOPO_TRIA3
    Else
        topo = TOPO_QUAD4
        If pN4(idx) <= 0 Then
            Exit Function
        End If
    End If

    On Error GoTo BailMake

    Set oEl = femap.feElem

    On Error Resume Next

    oEl.type = ELTYPE_PLATE
    oEl.topology = topo
    oEl.propID = pProp(idx)

    v = oEl.vnode
    v(0) = pN1(idx)
    v(1) = pN2(idx)
    v(2) = pN3(idx)
    If topo = TOPO_QUAD4 Then
        v(3) = pN4(idx)
    Else
        v(3) = 0
    End If
    oEl.vnode = v

    oEl.Put newID

    Err = 0
    On Error GoTo BailMake

    ' --- READ BACK ---
    Set oEl = femap.feElem
    If oEl.Get(newID) <> -1 Then
        Call NoteFail(newID, "element was not created")
        Exit Function
    End If

    If oEl.topology <> topo Then
        Call NoteFail(newID, "wrong topology after create")
        Exit Function
    End If

    If oEl.propID <> pProp(idx) Then
        Call NoteFail(newID, "wrong property after create")
        Exit Function
    End If

    v = oEl.vnode
    If CLng(v(0)) <> pN1(idx) Then
        Call NoteFail(newID, "node 1 mismatch")
        Exit Function
    End If
    If CLng(v(1)) <> pN2(idx) Then
        Call NoteFail(newID, "node 2 mismatch")
        Exit Function
    End If
    If CLng(v(2)) <> pN3(idx) Then
        Call NoteFail(newID, "node 3 mismatch")
        Exit Function
    End If

    MakeOne = 1
    failCount = 0
    Exit Function

BailMake:
    Call NoteFail(newID, "runtime error during create")

End Function

Sub NoteFail(eid As Long, reason As String)

    failCount = failCount + 1
    WL "  FAILED id " & eid & " : " & reason

    If failCount >= MAX_FAILS Then
        WL ""
        WL "  *** ABORTING after " & MAX_FAILS & " consecutive failures."
        WL "  *** Element creation is not working as written on this"
        WL "  *** build. Report the messages above."
        WL ""
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
