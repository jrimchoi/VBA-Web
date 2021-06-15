Attribute VB_Name = "PictureInsert"
Sub InsertPictureAsComment()
    Dim PicturePath As String
    Dim CommentBox As Comment
    With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = True
    .Title = "Select Comment Image"
    .ButtonName = "Insert Image"
    .Filters.Clear
    .Filters.Add "Images", "*.png; *.jpg"
    .Show
    'Store Selected File Path
    On Error GoTo UserCancelled
    PicturePath = .SelectedItems(1)
    On Error GoTo 0
    End With
    'Clear Any Existing Comment
    Application.ActiveCell.ClearComments
    'Create a New Cell Comment
    Set CommentBox = Application.ActiveCell.AddComment
    'Remove Any Default Comment Text
    CommentBox.Text Text:=""
    'Insert The Image and Resize
    CommentBox.Shape.Fill.UserPicture (PicturePath)
    CommentBox.Shape.ScaleHeight 6, msoFalse, msoScaleFormTopLeft
    CommentBox.Shape.ScaleWidth 4.8, msoFalse, msoScaleFromTopLeft
    'Ensure Comment is Hidden (Swith to TRUE if you want visible)
    CommentBox.Visible = False
    Exit Sub
    'ERROR HANDLERS
UserCancelled:
    MsgBox "Done"
End Sub

Sub URLToCellPictureInsert()
'Updateby Extendoffice 20180608
    Dim Pshp As Shape
    Dim xRg As Range
    Dim xCol As Long
    On Error Resume Next
    Set Rng = Application.InputBox("Please select the url cells:", "KuTools for excel", Selection.Address, , , , , 8)
    If Rng Is Nothing Then Exit Sub
    Set xRg = Application.InputBox("Please select a cell to put the image as comment:", "KuTools for excel", , , , , , 8)
    If xRg Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    For i = 1 To Rng.Count
        filenam = Rng(i)
        ActiveSheet.Pictures.Insert(filenam).Select
        Set Pshp = Selection.ShapeRange.Item(1)
        If Pshp Is Nothing Then GoTo lab
        xCol = cell.Column + 1
        Set xRg = xRg.Offset(i - 1, 0)
        With Pshp
            LockAspectRatio = msoFalse
            .Width = xRg.Width
            .Height = xRg.Height
            '.Top = xRg.Top + (xRg.Height - .Height) / 2
            '.Left = xRg.Left + (xRg.Width - .Width) / 2
            .Left = xRg.Left
            .Top = xRg.Top
        End With
lab:
        Set Pshp = Nothing
        Range("A2").Select
    Next
    Application.ScreenUpdating = True
End Sub

