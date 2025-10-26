### 1) Purpose & Benefits

<img width="1596" height="616" alt="螢幕擷取畫面 2025-10-26 145636" src="https://github.com/user-attachments/assets/c8d3d34c-ebb6-490e-b600-0f1ef839a1d0" />




### 2) As-Is

<img width="740" height="797" alt="image" src="https://github.com/user-attachments/assets/4c965013-6141-41ce-96e5-caef255afef5" />


### 3) Required 7 Manual Steps before
<img width="1223" height="1008" alt="螢幕擷取畫面 2025-10-26 145903" src="https://github.com/user-attachments/assets/14ba2440-735b-4b62-a32d-b4e5f784e41d" />
<img width="1224" height="861" alt="螢幕擷取畫面 2025-10-26 150003" src="https://github.com/user-attachments/assets/bdcc2ddd-4d4d-41cb-9fc9-58a2c973b971" />

### 4) Now, Just One Click Away
[Pls click this link to see the effect: https://youtu.be/EBg2MD4snbQ](https://youtu.be/EBg2MD4snbQ)

<details>
  <summary>VBA Code</summary>
  Sub Selectall()

Cells.Select

Selection.ClearFormats
Columns("A:C").EntireColumn.Hidden = True
Columns("F:J").EntireColumn.Hidden = True

Dim i As Long
Dim myRow As Long

myRow = Range("E1").End(xlDown).Row
    
    For i = 1 To 1000
    
If IsEmpty(Cells(i, 4).Value) = True Then
Range("A1:J10000").Rows(i).ClearContents
   
End If

If InStr(1, Cells(i, 4), "TOTAL") Then

Rows(i).Font.Bold = True


End If

If InStr(1, Cells(i, 4), "N.OPEX") Then

Rows(i).Font.Bold = True
End If
If InStr(1, Cells(i, 4), "NET (PROFIT)/LOSS") Then

Rows(i).Font.Bold = True
End If

If InStr(1, Cells(i, 4), "TTL NON-BUDGET ITEMS") Then

Rows(i).Font.Bold = True
End If


Next i

Range("A1:J10000").Rows(1).ClearContents
Rows(2).Insert
Rows(2).Insert
Rows(2).Insert

Range("D7:J7").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("D7:J7").Borders(xlEdgeBottom).Weight = xlMedium
Cells(7, 4).Value = "Account Name"
Columns(5).Insert
Cells(7, 5).Value = "Voucher No."
Cells(7, 6).Value = "Current Period Cost"
Rows(7).Font.Bold = True
Columns("D").ColumnWidth = 50
Columns("E").ColumnWidth = 16
Columns("F").ColumnWidth = 20
Cells(1, 4).Value = "EFFICIENT MANAGEMENT LIMITED"
Cells(2, 4).Value = "Profit and Loss for the month end "
Cells(3, 4).Value = "Company Name : XYZ"
Cells(4, 4).Value = "Vessel Name: ABC"

Rows(1).Font.Bold = True
Rows(2).Font.Bold = True
Rows(1).Font.Size = 14
Rows(2).Font.Size = 13
Rows(3).Font.Size = 13
Rows(4).Font.Size = 13
Columns("F").NumberFormat = "#,##0.00_);(#,##0.00)"
Columns("E").HorizontalAlignment = xlCenter
ActiveSheet.UsedRange.Select

 
End Sub
</details>
