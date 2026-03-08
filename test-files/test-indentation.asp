<%
' test-indentation.asp — Smart indentation test
'
' HOW TO USE THIS FILE:
'   Each test section has a comment telling you exactly WHERE to place your
'   cursor and WHAT key to press.  The "EXPECTED:" comment shows what the
'   next line should look like after you press that key.
'   An arrow  ←  marks the cursor position.  A pipe  |  shows col 0.
'
' KEY:
'   [Enter]  = press Enter / Return
'   [Tab]    = press Tab on a blank line
'   & _      = line-continuation character (string alignment test)


' ══════════════════════════════════════════════════════════════════════════════
' PART A — ENTER KEY: VBScript block openers
' ══════════════════════════════════════════════════════════════════════════════

' ── A1. If / Then ─────────────────────────────────────────────────────────────
' ACTION:  Place cursor at end of "If condition Then"  and press [Enter]
' EXPECTED: next line indented +1 level (2 spaces)

If condition Then←
    Response.Write "inside If"
End If

' ── A2. ElseIf — auto-snaps to If indent, body +1 ────────────────────────────
' ACTION:  Type "ElseIf" on the line below "Response.Write", then press [Enter]
' EXPECTED: ElseIf snaps to column 0 (same as If), next line indents to +1

If condition Then
    Response.Write "A"
ElseIf otherCondition Then←
    Response.Write "B"
Else←
    Response.Write "C"
End If←

' ── A3. For / Next ────────────────────────────────────────────────────────────
' ACTION:  Place cursor at end of "For i = 1 To 10"  and press [Enter]
' EXPECTED: next line at +1 indent

For i = 1 To 10←
    total = total + i
Next←

' ── A4. For Each / Next ───────────────────────────────────────────────────────
' ACTION:  Place cursor at end of "For Each item In itemList"  and press [Enter]
' EXPECTED: next line at +1 indent

For Each item In itemList←
    Response.Write item
Next←

' ── A5. While / Wend ─────────────────────────────────────────────────────────
' ACTION:  Place cursor at end of "While x < 10"  and press [Enter]
' EXPECTED: next line at +1 indent

While x < 10←
    x = x + 1
Wend←

' ── A6. Do While / Loop ──────────────────────────────────────────────────────
' ACTION:  Place cursor at end of "Do While Not rs.EOF"  and press [Enter]
' EXPECTED: next line at +1 indent

Do While Not rs.EOF←
    Response.Write rs("Name")
    rs.MoveNext
Loop←

' ── A7. Do / Loop Until ──────────────────────────────────────────────────────
' ACTION:  Place cursor at end of "Do"  and press [Enter]
' EXPECTED: +1 indent.  Then place cursor at end of "Loop Until" and press [Enter]
' EXPECTED: back to column 0

Do←
    cursor = cursor + 1
Loop Until cursor >= 10←

' ── A8. Select Case / Case / End Select ──────────────────────────────────────
' ACTION:  Place cursor at end of "Select Case status"  → [Enter] → expect +1
'          Type "Case" → auto-snaps to +1 (inside Select), body goes to +2
'          Type "End Select" → auto-snaps to column 0

Select Case status←
    Case "active"←
        Response.Write "Active"
    Case "pending"←
        Response.Write "Pending"
    Case Else←
        Response.Write "Unknown"
End Select←

' ── A9. Function / End Function ───────────────────────────────────────────────
' ACTION:  Place cursor at end of "Function CalculateTotal(...)"  and press [Enter]
' EXPECTED: next line at +1 indent inside the function

Function CalculateTotal(qty, price)←
    Dim total
    total = qty * price
    CalculateTotal = total
End Function←

' ── A10. Sub / End Sub ────────────────────────────────────────────────────────
' ACTION:  Place cursor at end of "Sub LogMessage(msg)"  and press [Enter]
' EXPECTED: next line at +1

Sub LogMessage(msg)←
    Response.Write "[LOG] " & msg
End Sub←

' ── A11. With / End With ──────────────────────────────────────────────────────
' ACTION:  Place cursor at end of "With rs"  and press [Enter]
' EXPECTED: next line at +1

With rs←
    .Open stmt, conn
    .Close
End With←

' ── A12. Class / End Class ────────────────────────────────────────────────────
' ACTION:  Place cursor at end of "Class UserModel"  and press [Enter]
' EXPECTED: next line at +1

Class UserModel←
    Private mName
    Public Property Get Name
        Name = mName
    End Property←
End Class←


' ══════════════════════════════════════════════════════════════════════════════
' PART B — ENTER KEY: Auto-snap closers to correct indent
' ══════════════════════════════════════════════════════════════════════════════

' ── B1. End If snap ───────────────────────────────────────────────────────────
' ACTION:  Deliberately type "End If" at column 0 (wrong indent).
'          As soon as you finish typing the last letter, it should snap
'          to align with its "If".

If condition Then
    Response.Write "test"
End If    ' ← try typing this at column 0 — it should snap to column 0 (correct here)

' Now nested — End If should snap to the inner If's indent (2 spaces):
If outer Then
    If inner Then
        Response.Write "nested"
    End If    ' ← try typing at col 0 — should snap to col 4
End If

' ── B2. Else snap ────────────────────────────────────────────────────────────
' ACTION:  Type "Else" indented too far (e.g. 8 spaces).
'          It should snap back to align with its "If".

If condition Then
    Response.Write "yes"
Else    ' ← try typing with too much indent — should snap to col 0
    Response.Write "no"
End If

' ── B3. Next snap ────────────────────────────────────────────────────────────
' ACTION:  Type "Next" indented incorrectly (e.g. 8 spaces).
'          Should snap to align with "For".

For i = 1 To 5
    total = total + i
Next    ' ← try typing at col 8 — should snap to col 0

' ── B4. Loop snap ────────────────────────────────────────────────────────────
' ACTION:  Type "Loop" at wrong indent.  Should snap to "Do While" column.

Do While cursor < 10
    cursor = cursor + 1
Loop    ' ← try at col 8 — should snap to col 0

' ── B5. Case snap ────────────────────────────────────────────────────────────
' ACTION:  Type "Case" at column 0. Should snap to +1 inside Select Case.

Select Case grade
    Case "A"    ' ← try typing at col 0 — should snap to col 2
        label = "Excellent"
    Case "B"    ' ← try at col 0 — should snap to col 2
        label = "Good"
End Select

' ── B6. %> snap to matching <% ───────────────────────────────────────────────
' ACTION:  Type "%>" at wrong indent inside an HTML block.
'          Should snap to align with its opening "<%".

%>
<div>
<%
    Response.Write "hello"
%>    <%' ← try typing at col 4 — should snap to col 0 %>
</div>
<%

' ── B7. Deeply nested — multiple levels of snap ───────────────────────────────
' ACTION:  Each closer below was typed at the wrong indent.
'          Verify each snaps to the correct level.
'          Expected final indentation shown in the comment on each line.

Function ProcessBatch(items)                 ' col 0
    For Each item In items                   ' col 2
        If IsNull(item) Then                 ' col 4
            item = ""                        ' col 6
        ElseIf item = "skip" Then            ' col 4  ← snap
            ' do nothing
        Else                                 ' col 4  ← snap
            Select Case item                 ' col 6
                Case "A"                     ' col 8  ← snap
                    count = count + 1        ' col 10
                Case Else                    ' col 8  ← snap
                    other = other + 1        ' col 10
            End Select                       ' col 6  ← snap
        End If                               ' col 4  ← snap
    Next                                     ' col 2  ← snap
End Function                                 ' col 0  ← snap


' ══════════════════════════════════════════════════════════════════════════════
' PART C — ENTER KEY: String continuation alignment  (& _)
' ══════════════════════════════════════════════════════════════════════════════

' ── C1. Basic string concat alignment ────────────────────────────────────────
' ACTION:  Place cursor at end of the first string line and press [Enter].
' EXPECTED: next line cursor aligns to column of the opening " (col 7)

stmt = "SELECT * FROM users " & _←
       "WHERE status = 'active' " & _←
       "ORDER BY name"←
' ← after the last line (no _), press [Enter]
' EXPECTED: cursor snaps BACK to column 0 (stmt's indent level)

' ── C2. Long variable name shifts alignment column ───────────────────────────
' ACTION:  Press [Enter] after each & _ line.
' EXPECTED: each new line aligns to column 18 (the opening " of the first line)

anotherlongstmt = "SELECT productCode " & _←
                  "FROM [SampleDb].[dbo].[Products] " & _←
                  "WHERE isActive = 1"←
' ← [Enter] here: cursor snaps back to column 0

' ── C3. Assignment-head-only first line  (= _ pattern) ───────────────────────
' ACTION:  Press [Enter] after "stmt = _"
' EXPECTED: cursor at col 2 (one indent level, no string column to align to)
' Then press [Enter] after each & _ line
' EXPECTED: cursor aligns to the " column established by the first string line

stmt = _←
    "SELECT productCode " & _←
    "FROM [SampleDb].[dbo].[Products] " & _←
    "WHERE isActive = 1"←
' ← [Enter] here: cursor snaps back to column 0

' ── C4. Blank line inside continuation chain — should still snap back correctly
stmt = "SELECT * FROM users " & _

       "ORDER BY name"←
' ← [Enter] here: cursor snaps back to column 0 (stmt's indent)

' ── C5. String concat inside a function (indent level > 0) ───────────────────
' ACTION:  Press [Enter] after each & _ line inside the function body.
' EXPECTED: alignment column is relative to document (NOT relative to function indent)
' i.e. cursor lands at the column of the " on the first string line.

Function BuildQuery(tableName)
    Dim q
    q = "SELECT * FROM " & tableName & _←
        " WHERE isActive = 1 " & _←
        " ORDER BY name"←
    ' ← [Enter]: snap back to col 4 (the indent level of the 'q =' line)
    BuildQuery = q
End Function


' ══════════════════════════════════════════════════════════════════════════════
' PART D — ENTER KEY: ASP tag boundaries
' ══════════════════════════════════════════════════════════════════════════════

' ── D1. Expand empty <%|%> block ──────────────────────────────────────────────
' ACTION:  Type <%  then %>  immediately (no space). Place cursor between them.
' EXPECTED: [Enter] expands to:
'   <%
'   |         ← cursor here
'   %>

%>
<%  %>    <%' ← place cursor between <% and %>, press [Enter] %>
<%

' ── D2. Enter after <% alone ─────────────────────────────────────────────────
' ACTION:  Place cursor at end of the standalone <% line and press [Enter].
' EXPECTED: next line at SAME indent as <% (VBScript code level = <% level)

%>
<%←
Response.Write "hello"
%>
<%

' ── D3. Enter after standalone %> ────────────────────────────────────────────
' ACTION:  Place cursor at end of the standalone %> line and press [Enter].
' EXPECTED: next line uses the enclosing HTML element's child indent.
' (If %> is inside a <ul>, next line should be at <ul>'s child level.)

%>
<ul>
  <%
  For Each item In items
  %> ←  <%' ← [Enter] here: should land at <ul> child indent (col 2) %>
  <li><%= item %></li>
  <%
  Next
  %>
</ul>
<%


' ══════════════════════════════════════════════════════════════════════════════
' PART E — TAB KEY: Smart Tab on blank lines
' ══════════════════════════════════════════════════════════════════════════════

' ── E1. Tab after a block opener snaps to +1 ─────────────────────────────────
' ACTION:  Place cursor on the blank line below "If condition Then" and press [Tab].
' EXPECTED: cursor moves to col 2 (one indent level)

If condition Then
←                ' ← blank line, press [Tab] → should jump to col 2
    Response.Write "test"
End If

' ── E2. Tab after a regular line stays at same level ─────────────────────────
' ACTION:  Place cursor on the blank line below "Response.Write" and press [Tab].
' EXPECTED: cursor moves to same indent as the line above

If condition Then
    Response.Write "test"
←                ' ← blank line, press [Tab] → should jump to col 2 (same as above)
End If

' ── E3. Tab below a closer snaps to that closer's indent ─────────────────────
' ACTION:  Place cursor on the blank line below "End If" and press [Tab].
' EXPECTED: cursor moves to col 0

If condition Then
    Response.Write "test"
End If
←                ' ← blank line below End If, press [Tab] → should go to col 0

' ── E4. Tab below <% ─────────────────────────────────────────────────────────
' ACTION:  Place cursor on the blank line after <% and press [Tab].
' EXPECTED: cursor at same col as <% (no extra indent added for the <% line itself)

%>
<%
←                <%' ← blank line, press [Tab] → col 0 (same as <%) %>
Response.Write "test"
%>
<%

' ── E5. Tab below %> returns to HTML child indent ────────────────────────────
' ACTION:  Place cursor on the blank line after the %> inside the <div> and press [Tab].
' EXPECTED: cursor at col 2 (the <div>'s child indent)

%>
<div>
  <%
  Response.Write "test"
  %>
←                <%' ← blank line, press [Tab] → col 2 (inside <div>) %>
</div>
<%


' ══════════════════════════════════════════════════════════════════════════════
' PART F — HTML ENTER: Block expansion
' ══════════════════════════════════════════════════════════════════════════════
%>

<!-- F1. Press [Enter] between opening and closing tag on same line -->
<!-- ACTION: <div>|</div>  → [Enter] should expand to:  -->
<!--   <div>                                             -->
<!--     |    ← cursor                                   -->
<!--   </div>                                            -->

<div>←</div>

<!-- F2. Press [Enter] after an opening tag (closing tag on next line) -->
<!-- EXPECTED: indent +1, cursor inside the block -->

<ul>←
  <li>item</li>
</ul>

<!-- F3. Press [Enter] after an opening tag when no closing tag exists yet -->
<!-- ACTION: type <section>  then press [Enter]              -->
<!-- EXPECTED:                                               -->
<!--   <section>                                             -->
<!--     |    ← cursor, AND </section> is auto-created below -->

<!-- F4. Self-closing tags — [Enter] after them should NOT expand -->
<!-- EXPECTED: plain newline at same indent, no </br> etc. -->

<br>←
<hr>←
<input type="text">←

<!-- F5. HTML comment auto-close -->
<!-- ACTION: type <!--  (the fourth dash) -->
<!-- EXPECTED: cursor lands inside  <!--  |  --> -->
<!--  -->
