Attribute VB_Name = "Styles"
' Description:
'    Borrowed largely from http://www.jkp-ads.com/Articles/styles06.asp

Option Explicit

' Description:
'    This is the "driver" for the entire module.
Public Sub DropUnusedStyles()

    Dim styleObj As Style
    Dim rngCell As Range
    Dim wb As Workbook
    Dim wsh As Worksheet
    Dim str As String
    Dim iStyleCount As Long
    Dim dict As New Scripting.Dictionary    ' <- from Tools / References... / "Microsoft Scripting Runtime"

    ' wb := workbook of interest.  Choose one of the following
    ' Set wb = ThisWorkbook ' choose this module's workbook
    Set wb = ActiveWorkbook ' the active workbook in excel


    Debug.Print "BEGINNING # of styles in workbook: " & wb.Styles.Count
    MsgBox "BEGINNING # of styles in workbook: " & wb.Styles.Count

    ' dict := list of styles
    For Each styleObj In wb.Styles
        str = styleObj.NameLocal
        iStyleCount = iStyleCount + 1
        Call dict.Add(str, 0)    ' First time:  adds keys
    Next styleObj
    Debug.Print "  dictionary now has " & dict.Count & " entries."
    ' Status, dictionary has styles (key) which are known to workbook


    ' Traverse each visible worksheet and increment count each style occurrence
    For Each wsh In wb.Worksheets
        If wsh.Visible Then
            For Each rngCell In wsh.UsedRange.Cells
                str = rngCell.Style
                dict.Item(str) = dict.Item(str) + 1     ' This time:  counts occurrences
            Next rngCell
        End If
    Next wsh
    ' Status, dictionary styles (key) has cell occurrence count (item)


    ' Try to delete unused styles
    Dim aKey As Variant
    On Error Resume Next    ' wb.Styles(aKey).Delete may throw error

    For Each aKey In dict.Keys

        ' display count & stylename
        '    e.g. "24   Normal"
        Debug.Print dict.Item(aKey) & vbTab & aKey

        If dict.Item(aKey) = 0 Then
            ' Occurrence count (Item) indicates this style is not used
            Call wb.Styles(aKey).Delete
            If Err.Number <> 0 Then
                Debug.Print vbTab & "^-- failed to delete"
                Err.Clear
            End If
            Call dict.Remove(aKey)
        End If

    Next aKey

    Debug.Print "ENDING # of style in workbook: " & wb.Styles.Count
    MsgBox "ENDING # of style in workbook: " & wb.Styles.Count

End Sub

