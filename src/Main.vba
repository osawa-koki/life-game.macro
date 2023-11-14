Option Explicit

Function LiveNeighborCount(row As Integer, Column As Integer, Width As Integer, Height As Integer, Cells() As Boolean) As Integer
    Dim count As Integer
    count = 0
    Dim delta_row As Variant
    Dim tmp_array() As Variant
    tmp_array = Array(-1, 0, 1)

    For Each delta_row In tmp_array
        Dim delta_col As Variant
        For Each delta_col In tmp_array
            If delta_row = 0 And delta_col = 0 Then
                GoTo NextLoop
            End If

            Dim neighbor_row As Integer
            neighbor_row = (row + delta_row) Mod Height
            Dim neighbor_col As Integer
            neighbor_col = (Column + delta_col) Mod Width
            Dim idx As Integer
            idx = (neighbor_row - 1) * Width + (neighbor_col - 1)
            If idx >= 0 And idx < UBound(Cells) Then
                If Cells(idx) Then
                    count = count + 1
                End If
            End If
NextLoop:
        Next delta_col
    Next delta_row
    LiveNeighborCount = count
End Function

Sub Show(Cells() As Boolean, Width As Integer, Height As Integer, sheet As Worksheet)
    Dim row As Integer
    Dim col As Integer
    For row = 1 To Height
        For col = 1 To Width
            Dim idx As Integer
            idx = (row - 1) * Width + (col - 1)
            If Cells(idx) Then
                sheet.Cells(row, col).Interior.Color = RGB(0, 0, 0)
            Else
                sheet.Cells(row, col).Interior.Color = RGB(255, 255, 255)
            End If
        Next col
    Next row
End Sub

Sub LifeGame()
    Dim Width As Integer
    Dim Height As Integer
    Dim Cells() As Boolean
    Dim SheetName As String

    Width = 64
    Height = 64
    ReDim Cells(0 To Width * Height - 1)
    SheetName = "LifeGame"

    Dim idx As Integer
    For idx = 0 To UBound(Cells)
        Dim cell As Boolean
        cell = False
        If idx Mod 2 = 0 Or idx Mod 7 = 0 Then
            cell = True
        End If
        Cells(idx) = cell
    Next idx

    ' イミディエイトウィンドウの出力をクリア。
    Debug.Print String(100, vbCrLf)

    ' デバグ用に値を出力。
    Debug.Print "Width: " & Width
    Debug.Print "Height: " & Height
    Debug.Print "SheetName: " & SheetName

    ' シートの削除
    Application.DisplayAlerts = False ' メッセージを非表示
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = SheetName Then
        ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True  ' メッセージを表示

    ' シートの追加
    Dim sheet As Worksheet
    Set sheet = Worksheets.Add
    sheet.Name = SheetName
    sheet.Activate

    ' 行と列のサイズを設定
    sheet.Range(Rows(1), Rows(Height)).RowHeight = 7.5
    sheet.Range(Columns(1), Columns(Width)).ColumnWidth = 0.77

    Application.Wait Now() + TimeValue("00:00:01")

    Dim neighbor_count As Integer
    Dim generation As Integer
    generation = 0
    Do
        sheet.Name = SheetName & "(" & generation & ")"
        Call Show(Cells, Width, Height, sheet)
        generation = generation + 1
        Dim new_cells() As Boolean
        ReDim new_cells(0 To Width * Height - 1)
        Dim row As Integer
        For row = 1 To Height
            Dim col As Integer
            For col = 1 To Width
                idx = (row - 1) * Width + (col - 1)
                cell = Cells(idx)
                neighbor_count = LiveNeighborCount(row, col, Width, Height, Cells)
                If cell Then
                    If neighbor_count = 2 Or neighbor_count = 3 Then
                        new_cells(idx) = True
                    Else
                        new_cells(idx) = False
                    End If
                Else
                    If neighbor_count = 3 Then
                        new_cells(idx) = True
                    Else
                        new_cells(idx) = False
                    End If
                End If
            Next col
        Next row
        Cells = new_cells
        Application.Wait Now() + TimeValue("00:00:01")
    Loop
End Sub
