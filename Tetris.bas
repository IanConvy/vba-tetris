Attribute VB_Name = "Tetris"
Private Pieces As Object 'Dictionary for piece coordinates'
Private PieceColors As Object 'Dictionary for piece colors'
Private pieceLocation As Long 'Column # of the leftmost piece block'
Private pieceOrient As Long 'Orientation of the piece'
Private pieceName As String 'Name of the current piece'
Private startHeight As Long 'Height to position new piece'

'Resets the game by clearing the board/line counts and getting a new piece'
Sub Reset()
Attribute Reset.VB_Description = "Starts new game of Tetris."
Attribute Reset.VB_ProcData.VB_Invoke_Func = "R\n14"
    Application.ScreenUpdating = False
    startHeight = 3
    CreatePieceDict
    Randomize
    ClearRange Range("board")
    Range("lines") = 0
    NextPiece
End Sub

'Moves piece to the left or right at the top of the board'
Sub ShiftPiece(direct As Long)
    PlacePiece pieceName, startHeight, pieceLocation, pieceOrient, True
    If Not PieceCollision(pieceName, startHeight, pieceLocation + direct, pieceOrient) Then
        pieceLocation = pieceLocation + direct
    End If
    PlacePiece pieceName, startHeight, pieceLocation, pieceOrient
End Sub

'Rotates piece in place (roughly) at the top of the board'
Sub RotatePiece()
    Dim newOrient As Long
    Application.ScreenUpdating = False
    PlacePiece pieceName, startHeight, pieceLocation, pieceOrient, True
    If pieceOrient = 3 Then
        newOrient = 0
    Else
        newOrient = pieceOrient + 1
    End If
    If Not PieceCollision(pieceName, startHeight, pieceLocation, newOrient) Then
        pieceOrient = newOrient
    End If
    PlacePiece pieceName, startHeight, pieceLocation, pieceOrient
End Sub

'Places piece by letting it fall vertically until it collides'
Sub DropPiece()
    Dim height As Long
    Application.ScreenUpdating = False
    PlacePiece pieceName, startHeight, pieceLocation, pieceOrient, True
    For height = startHeight + 1 To 24
    If PieceCollision(pieceName, height, pieceLocation, pieceOrient) Then
        height = height - 1
        Exit For
    End If
    Next height
    PlacePiece pieceName, height, pieceLocation, pieceOrient
    ClearLines
    NextPiece
End Sub

'Retrieves a random piece, assigns it to pieceName, and places it on the board'
Private Sub NextPiece()
    pieceLocation = 4
    pieceOrient = 0
    Select Case Rnd() * 7
        Case 0 To 1
            pieceName = "sq"
        Case 1 To 2
            pieceName = "i"
        Case 2 To 3
            pieceName = "t"
        Case 3 To 4
            pieceName = "zl"
        Case 4 To 5
            pieceName = "zr"
        Case 5 To 6
            pieceName = "ll"
        Case Else
            pieceName = "lr"
    End Select
    PlacePiece pieceName, startHeight, pieceLocation, pieceOrient
End Sub

'Colors in (or clears) board cells based on the current piece and the given coordinates'
Private Sub PlacePiece(piece As String, x0 As Long, y0 As Long, orient As Long, Optional blank As Boolean = False)
    Dim coords As Variant
    Dim block As Long, x As Long, y As Long
    Dim cell As Range
    If Pieces Is Nothing Then Reset
    coords = Pieces.Item(piece)
    For block = 0 To 3
        x = x0 + coords(orient, block, 0)
        y = y0 + coords(orient, block, 1)
        If (x > 0) And (y > 0) Then
            Set cell = Range("board")(x, y)
            If blank Then
                ClearRange cell
            Else
                AddBlocks cell, piece
            End If
        End If
    Next block
End Sub

'Removes piece blocks from the board'
Private Sub ClearRange(cells As Range)
    cells.Clear
    cells.Interior.ColorIndex = xlNone
End Sub

'Adds piece blocks to the board based on given Range and piece'
Private Sub AddBlocks(cells As Range, piece As String)
    Dim color As Long
    color = PieceColors.Item(piece)
    cells.Interior.ColorIndex = color
    cells.Font.ColorIndex = color
    cells.Value = 1
    cells.BorderAround xlContinuous, xlThin, xlAutomatic
End Sub

'Removes any board rows that have been completely filled'
Private Sub ClearLines()
    Dim height As Long
    Dim board As Range: Set board = Range("board")
    Dim subrange As Range
    Dim lineCount As Long: lineCount = 0
    For height = 1 To 22
        If WorksheetFunction.Sum(Range(board.cells(height, 1), board.cells(height, 10))) = 10 Then
            lineCount = lineCount + 1
            Range(board.cells(1, 1), board.cells(height - 1, 10)).Copy
            Range(board.cells(2, 1), board.cells(height, 10)).PasteSpecial xlPasteAll
            Range("A1").Select
        End If
    Next height
    UpdateScore lineCount
End Sub

'Updates the line count table based on the number of clears'
Private Sub UpdateScore(lineClears As Long)
    If lineClears > 0 Then
        Range("lines")(1, lineClears) = Range("lines")(1, lineClears) + 1
    End If
End Sub

'Checks if a piece will collide with walls or other pieces at the given position'
Private Function PieceCollision(piece As String, x0 As Long, y0 As Long, orient As Long) As Boolean
    Dim coords As Variant
    Dim block As Long, x As Long, y As Long
    coords = Pieces.Item(piece)
    PieceCollision = False
    For block = 0 To 3
        x = x0 + coords(orient, block, 0)
        y = y0 + coords(orient, block, 1)
        PieceCollision = PieceCollision Or CheckCollision(x, y)
    Next block
End Function

'Checks if a specified cell is already occupied or is out of bounds.
Private Function CheckCollision(x0 As Long, y0 As Long) As Boolean
    'Check left boundary collision'
    CheckCollision = Not (Intersect(Range("board")(x0, y0), Range("left")) Is Nothing)
    'Check right boundary collision'
    CheckCollision = CheckCollision Or Not (Intersect(Range("board")(x0, y0), Range("right")) Is Nothing)
    'Check bottom boundary condition'
    CheckCollision = CheckCollision Or Not (Intersect(Range("board")(x0, y0), Range("bottom")) Is Nothing)
    'Check piece collision'
    CheckCollision = CheckCollision Or (Range("board")(x0, y0) = 1)
End Function

'Creates the dictionaries which hold all immutable piece information'
Private Sub CreatePieceDict()
    'Assign piece colors'
    
    Set PieceColors = CreateObject("Scripting.Dictionary")
    PieceColors.Add "sq", 15
    PieceColors.Add "i", 15
    PieceColors.Add "t", 15
    PieceColors.Add "zl", 33
    PieceColors.Add "zr", 22
    PieceColors.Add "ll", 33
    PieceColors.Add "lr", 22

    'Declare coordinate arrays'

    Dim sqCoords(3, 3, 1) As Long
    Dim iCoords(3, 3, 1) As Long
    Dim tCoords(3, 3, 1) As Long
    Dim zlCoords(3, 3, 1) As Long
    Dim zrCoords(3, 3, 1) As Long
    Dim llCoords(3, 3, 1) As Long
    Dim lrCoords(3, 3, 1) As Long
    
    'Define square piece coordinates'
    
    sqCoords(0, 0, 0) = 0
    sqCoords(0, 0, 1) = 0
    sqCoords(0, 1, 0) = 0
    sqCoords(0, 1, 1) = 1
    sqCoords(0, 2, 0) = 1
    sqCoords(0, 2, 1) = 0
    sqCoords(0, 3, 0) = 1
    sqCoords(0, 3, 1) = 1
    
    sqCoords(1, 0, 0) = 0
    sqCoords(1, 0, 1) = 0
    sqCoords(1, 1, 0) = 0
    sqCoords(1, 1, 1) = 1
    sqCoords(1, 2, 0) = 1
    sqCoords(1, 2, 1) = 0
    sqCoords(1, 3, 0) = 1
    sqCoords(1, 3, 1) = 1
    
    sqCoords(2, 0, 0) = 0
    sqCoords(2, 0, 1) = 0
    sqCoords(2, 1, 0) = 0
    sqCoords(2, 1, 1) = 1
    sqCoords(2, 2, 0) = 1
    sqCoords(2, 2, 1) = 0
    sqCoords(2, 3, 0) = 1
    sqCoords(2, 3, 1) = 1
    
    sqCoords(3, 0, 0) = 0
    sqCoords(3, 0, 1) = 0
    sqCoords(3, 1, 0) = 0
    sqCoords(3, 1, 1) = 1
    sqCoords(3, 2, 0) = 1
    sqCoords(3, 2, 1) = 0
    sqCoords(3, 3, 0) = 1
    sqCoords(3, 3, 1) = 1
    
    'Define I piece coordinates'
    
    iCoords(0, 0, 0) = 0
    iCoords(0, 0, 1) = 0
    iCoords(0, 1, 0) = 0
    iCoords(0, 1, 1) = 1
    iCoords(0, 2, 0) = 0
    iCoords(0, 2, 1) = 2
    iCoords(0, 3, 0) = 0
    iCoords(0, 3, 1) = 3
    
    iCoords(1, 0, 0) = -1
    iCoords(1, 0, 1) = 0
    iCoords(1, 1, 0) = 0
    iCoords(1, 1, 1) = 0
    iCoords(1, 2, 0) = 1
    iCoords(1, 2, 1) = 0
    iCoords(1, 3, 0) = 2
    iCoords(1, 3, 1) = 0
    
    iCoords(2, 0, 0) = 0
    iCoords(2, 0, 1) = 0
    iCoords(2, 1, 0) = 0
    iCoords(2, 1, 1) = 1
    iCoords(2, 2, 0) = 0
    iCoords(2, 2, 1) = 2
    iCoords(2, 3, 0) = 0
    iCoords(2, 3, 1) = 3
    
    iCoords(3, 0, 0) = -1
    iCoords(3, 0, 1) = 0
    iCoords(3, 1, 0) = 0
    iCoords(3, 1, 1) = 0
    iCoords(3, 2, 0) = 1
    iCoords(3, 2, 1) = 0
    iCoords(3, 3, 0) = 2
    iCoords(3, 3, 1) = 0
    
    'Define T piece coordinates'
    
    tCoords(0, 0, 0) = 0
    tCoords(0, 0, 1) = 0
    tCoords(0, 1, 0) = 0
    tCoords(0, 1, 1) = 1
    tCoords(0, 2, 0) = 1
    tCoords(0, 2, 1) = 1
    tCoords(0, 3, 0) = 0
    tCoords(0, 3, 1) = 2
    
    tCoords(1, 0, 0) = 0
    tCoords(1, 0, 1) = 0
    tCoords(1, 1, 0) = 0
    tCoords(1, 1, 1) = 1
    tCoords(1, 2, 0) = 1
    tCoords(1, 2, 1) = 1
    tCoords(1, 3, 0) = -1
    tCoords(1, 3, 1) = 1
    
    tCoords(2, 0, 0) = 0
    tCoords(2, 0, 1) = 0
    tCoords(2, 1, 0) = 0
    tCoords(2, 1, 1) = 1
    tCoords(2, 2, 0) = -1
    tCoords(2, 2, 1) = 1
    tCoords(2, 3, 0) = 0
    tCoords(2, 3, 1) = 2
    
    tCoords(3, 0, 0) = 0
    tCoords(3, 0, 1) = 0
    tCoords(3, 1, 0) = 0
    tCoords(3, 1, 1) = 1
    tCoords(3, 2, 0) = -1
    tCoords(3, 2, 1) = 0
    tCoords(3, 3, 0) = 1
    tCoords(3, 3, 1) = 0
    
    'Define Z-right piece coordinates'
    
    zrCoords(0, 0, 0) = 1
    zrCoords(0, 0, 1) = 0
    zrCoords(0, 1, 0) = 1
    zrCoords(0, 1, 1) = 1
    zrCoords(0, 2, 0) = 0
    zrCoords(0, 2, 1) = 1
    zrCoords(0, 3, 0) = 0
    zrCoords(0, 3, 1) = 2
    
    zrCoords(1, 0, 0) = 0
    zrCoords(1, 0, 1) = 0
    zrCoords(1, 1, 0) = -1
    zrCoords(1, 1, 1) = 0
    zrCoords(1, 2, 0) = 0
    zrCoords(1, 2, 1) = 1
    zrCoords(1, 3, 0) = 1
    zrCoords(1, 3, 1) = 1
    
    zrCoords(2, 0, 0) = 1
    zrCoords(2, 0, 1) = 0
    zrCoords(2, 1, 0) = 1
    zrCoords(2, 1, 1) = 1
    zrCoords(2, 2, 0) = 0
    zrCoords(2, 2, 1) = 1
    zrCoords(2, 3, 0) = 0
    zrCoords(2, 3, 1) = 2
    
    zrCoords(3, 0, 0) = 0
    zrCoords(3, 0, 1) = 0
    zrCoords(3, 1, 0) = -1
    zrCoords(3, 1, 1) = 0
    zrCoords(3, 2, 0) = 0
    zrCoords(3, 2, 1) = 1
    zrCoords(3, 3, 0) = 1
    zrCoords(3, 3, 1) = 1
    
    'Define Z-left piece coordinates'
    
    zlCoords(0, 0, 0) = 0
    zlCoords(0, 0, 1) = 0
    zlCoords(0, 1, 0) = 0
    zlCoords(0, 1, 1) = 1
    zlCoords(0, 2, 0) = 1
    zlCoords(0, 2, 1) = 1
    zlCoords(0, 3, 0) = 1
    zlCoords(0, 3, 1) = 2
    
    zlCoords(1, 0, 0) = 0
    zlCoords(1, 0, 1) = 0
    zlCoords(1, 1, 0) = 1
    zlCoords(1, 1, 1) = 0
    zlCoords(1, 2, 0) = 0
    zlCoords(1, 2, 1) = 1
    zlCoords(1, 3, 0) = -1
    zlCoords(1, 3, 1) = 1
    
    zlCoords(2, 0, 0) = 0
    zlCoords(2, 0, 1) = 0
    zlCoords(2, 1, 0) = 0
    zlCoords(2, 1, 1) = 1
    zlCoords(2, 2, 0) = 1
    zlCoords(2, 2, 1) = 1
    zlCoords(2, 3, 0) = 1
    zlCoords(2, 3, 1) = 2
    
    zlCoords(3, 0, 0) = 0
    zlCoords(3, 0, 1) = 0
    zlCoords(3, 1, 0) = 1
    zlCoords(3, 1, 1) = 0
    zlCoords(3, 2, 0) = 0
    zlCoords(3, 2, 1) = 1
    zlCoords(3, 3, 0) = -1
    zlCoords(3, 3, 1) = 1
    
    'Define L-right piece coordinates'
    
    lrCoords(0, 0, 0) = 0
    lrCoords(0, 0, 1) = 0
    lrCoords(0, 1, 0) = 1
    lrCoords(0, 1, 1) = 0
    lrCoords(0, 2, 0) = 0
    lrCoords(0, 2, 1) = 1
    lrCoords(0, 3, 0) = 0
    lrCoords(0, 3, 1) = 2
    
    lrCoords(1, 0, 0) = 0
    lrCoords(1, 0, 1) = 1
    lrCoords(1, 1, 0) = -1
    lrCoords(1, 1, 1) = 1
    lrCoords(1, 2, 0) = -1
    lrCoords(1, 2, 1) = 0
    lrCoords(1, 3, 0) = 1
    lrCoords(1, 3, 1) = 1
    
    lrCoords(2, 0, 0) = 0
    lrCoords(2, 0, 1) = 0
    lrCoords(2, 1, 0) = 0
    lrCoords(2, 1, 1) = 1
    lrCoords(2, 2, 0) = 0
    lrCoords(2, 2, 1) = 2
    lrCoords(2, 3, 0) = -1
    lrCoords(2, 3, 1) = 2
    
    lrCoords(3, 0, 0) = 0
    lrCoords(3, 0, 1) = 0
    lrCoords(3, 1, 0) = -1
    lrCoords(3, 1, 1) = 0
    lrCoords(3, 2, 0) = 1
    lrCoords(3, 2, 1) = 0
    lrCoords(3, 3, 0) = 1
    lrCoords(3, 3, 1) = 1
    
    'Define L-left piece coordinates'
    
    llCoords(0, 0, 0) = 0
    llCoords(0, 0, 1) = 0
    llCoords(0, 1, 0) = 0
    llCoords(0, 1, 1) = 1
    llCoords(0, 2, 0) = 0
    llCoords(0, 2, 1) = 2
    llCoords(0, 3, 0) = 1
    llCoords(0, 3, 1) = 2
    
    llCoords(1, 0, 0) = 1
    llCoords(1, 0, 1) = 0
    llCoords(1, 1, 0) = 1
    llCoords(1, 1, 1) = 1
    llCoords(1, 2, 0) = 0
    llCoords(1, 2, 1) = 1
    llCoords(1, 3, 0) = -1
    llCoords(1, 3, 1) = 1
    
    llCoords(2, 0, 0) = 0
    llCoords(2, 0, 1) = 0
    llCoords(2, 1, 0) = -1
    llCoords(2, 1, 1) = 0
    llCoords(2, 2, 0) = 0
    llCoords(2, 2, 1) = 1
    llCoords(2, 3, 0) = 0
    llCoords(2, 3, 1) = 2
    
    llCoords(3, 0, 0) = 0
    llCoords(3, 0, 1) = 0
    llCoords(3, 1, 0) = -1
    llCoords(3, 1, 1) = 0
    llCoords(3, 2, 0) = -1
    llCoords(3, 2, 1) = 1
    llCoords(3, 3, 0) = 1
    llCoords(3, 3, 1) = 0
    
    'Create public dictionary for easy access'
    
    Set Pieces = CreateObject("Scripting.Dictionary")
    Pieces.Add "sq", sqCoords
    Pieces.Add "i", iCoords
    Pieces.Add "t", tCoords
    Pieces.Add "zl", zlCoords
    Pieces.Add "zr", zrCoords
    Pieces.Add "ll", llCoords
    Pieces.Add "lr", lrCoords

End Sub
