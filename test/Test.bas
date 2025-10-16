Attribute VB_Name = "Test"



' Run all tests.
Public Sub Test()
	Debug.Print "#################"
	Debug.Print "## Arr_Index() ##"
	Debug.Print "#################"
	Debug.Print
	Test__Arr_Index
	
	Debug.Print
	Debug.Print
	Debug.Print
	
	Debug.Print "#############"
	Debug.Print "## Index() ##"
	Debug.Print "#############"
	Debug.Print
	Test__Index
End Sub



' #################
' ## Arr_Index() ##
' #################

' Run all tests on Arr_Index().
Public Sub Test__Arr_Index(ByRef multiArr As Variant)
	Debug.Print "##########################################"
	Debug.Print "## Arr_Index() | Multidimensional Array ##"
	Debug.Print "##########################################"
	Debug.Print
	Test__Arr_Index__MultiArray arr := multiArr
End Sub


' Test Arr_Index() on a multidimensional array.
Public Sub Test__Arr_Index__MultiArray(ByRef arr As Variant)
	' ' Populate 3D array.
	' Dim arr(1 To 2, 3 To 4, 5 To 6) As String
	' arr(1, 3, 5) = "1.3.5"
	' arr(1, 3, 6) = "1.3.6"
	' arr(1, 4, 5) = "1.4.5"
	' arr(1, 4, 6) = "1.4.6"
	' arr(2, 3, 5) = "2.3.5"
	' arr(2, 3, 6) = "2.3.6"
	' arr(2, 4, 5) = "2.4.5"
	' arr(2, 4, 6) = "2.4.6"
	
	
	' Display array metadata.
	Debug.Print "arr.Rank = " & Idx.Arr_Rank(arr)
	Debug.Print
	Debug.Print "arr.Length(1) = " & Idx.Arr_Length(arr, dimension := 1)
	Debug.Print "arr.Length(2) = " & Idx.Arr_Length(arr, dimension := 2)
	Debug.Print "arr.Length(3) = " & Idx.Arr_Length(arr, dimension := 3)
	
	
	' Extract data.
	Debug.Print "arr()           = """ & Idx.Arr_Index(arr, Array()) & """"
	Debug.Print "arr(1)          = """ & Idx.Arr_Index(arr, Array(1)) & """"
	Debug.Print "arr(1, 4)       = """ & Idx.Arr_Index(arr, Array(1, 4)) & """"
	Debug.Print "arr(1, 4, 5)    = """ & Idx.Arr_Index(arr, Array(1, 4, 5)) & """"
	Debug.Print "arr(1, 4, 5, 7) = """ & Idx.Arr_Index(arr, Array(1, 4, 5, 7)) & """"
End Sub



' #############
' ## Index() ##
' #############

' Run all tests on Index().
Public Sub Test__Index(ByRef nestArr As Variant)
	Debug.Print "############################"
	Debug.Print "## Index() | Nested Array ##"
	Debug.Print "############################"
	Debug.Print
	Test__Index__NestedArray arr := nestArr
End Sub


' Test Index() on a nested array.
Public Sub Test__Index__NestedArray(ByRef arr As Variant)
	' ' Nest arrays 3 deep.
	' Dim arr As Variant: arr = Array( _
	' 	Array( _
	' 		Array( _
	' 			"0.0.0", _
	' 			"0.0.1" _
	' 		), _
	' 		Array( _
	' 			"0.1.0", _
	' 			"0.1.1" _
	' 		) _
	' 	), _
	' 	Array( _
	' 		Array( _
	' 			"1.0.0", _
	' 			"1.0.1" _
	' 		), _
	' 		Array( _
	' 			"1.1.0", _
	' 			"1.1.1" _
	' 		) _
	' 	) _
	' )
	
	
	' Display array metadata.
	Debug.Print "arr.Rank       = " & Idx.Arr_Rank(a)
	Debug.Print "arr(0).Rank    = " & Idx.Arr_Rank(a(0))
	Debug.Print "arr(0)(0).Rank = " & Idx.Arr_Rank(a(0)(0))
	Debug.Print
	Debug.Print "arr.Length       = " & Idx.Arr_Length(a, dimension := 1)
	Debug.Print "arr(0).Length    = " & Idx.Arr_Length(a(0), dimension := 1)
	Debug.Print "arr(0)(0).Length = " & Idx.Arr_Length(a(0)(0), dimension := 1)
	
	
	' Extract data.
	Debug.Print "arr()           = """ & Idx.Index(arr, Array()) & """
	Debug.Print "arr(0)          = """ & Idx.Index(arr, Array(0)) & """
	Debug.Print "arr(0)(1)       = """ & Idx.Index(arr, Array(0, 1)) & """
	Debug.Print "arr(0)(1)(0)    = """ & Idx.Index(arr, Array(0, 1, 0)) & """
	Debug.Print "arr(0)(1)(0)(1) = """ & Idx.Index(arr, Array(0, 1, 0, 1)) & """
End Sub
