Option Explicit

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

    ' セル一覧をループしてセルの背景色を設定
    Dim row As Integer
    Dim col As Integer
    For row = 1 To Height
        For col = 1 To Width
            idx = (row - 1) * Width + (col - 1)
            If Cells(idx) Then
                sheet.Cells(row, col).Interior.Color = RGB(0, 0, 0)
            Else
                sheet.Cells(row, col).Interior.Color = RGB(255, 255, 255)
            End If
        Next col
    Next row
End Sub
