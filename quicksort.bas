Attribute VB_Name = "quicksort"
Public Sub subQuickSort(lonLower, lonUpper)
    On Error Resume Next
    Dim lonRandomPivot As Long
    Dim lonTempLower As Long
    Dim lonTempUpper As Long
    Dim tmpitem     As my_file
    Dim lastitem    As my_file
    
    Randomize Timer
    
    If lonLower < lonUpper Then
        If lonUpper - lonLower = 1 Then

            If files_found(lonLower).folder > files_found(lonUpper).folder Then
                tmpitem = files_found(lonUpper)
                files_found(lonUpper) = files_found(lonLower)
                files_found(lonLower) = tmpitem
            End If
        Else

            lonRandomPivot = Int(Rnd _
                * (lonUpper - lonLower + 1)) + lonLower

            tmpitem = files_found(lonUpper)
            files_found(lonUpper) = files_found(lonRandomPivot)
            files_found(lonRandomPivot) = tmpitem

            lastitem = files_found(lonUpper)
            Do

                lonTempUpper = lonUpper
                lonTempLower = lonLower

                Do While (lonTempLower < lonTempUpper) And _
                    (files_found(lonTempLower).folder <= lastitem.folder)
                        lonTempLower = lonTempLower + 1
                Loop

                Do While (lonTempUpper > lonTempLower) And _
                    (files_found(lonTempUpper).folder >= lastitem.folder)
                        lonTempUpper = lonTempUpper - 1
                Loop

                If lonTempLower < lonTempUpper Then
                    tmpitem = files_found(lonTempUpper)
                    files_found(lonTempUpper) = files_found(lonTempLower)
                    files_found(lonTempLower) = tmpitem
                End If
            Loop While (lonTempLower < lonTempUpper)

            tmpitem = files_found(lonTempLower)
            files_found(lonTempLower) = files_found(lonUpper)
            files_found(lonUpper) = tmpitem

            If (lonTempLower - lonLower) < (lonUpper - lonTempLower) Then
                subQuickSort lonLower, lonTempLower - 1
                subQuickSort lonTempLower + 1, lonUpper
            Else
                subQuickSort lonTempLower + 1, lonUpper
                subQuickSort lonLower, lonTempLower - 1
            End If
        End If
    End If

End Sub


