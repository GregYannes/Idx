Attribute VB_Name = "Test"



Public Sub Test__Arr_Index__3dArray()
	' Populate 3D array.
	Dim a(1 To 2, 3 To 4, 5 To 6) As String
	a(1, 3, 5) = "1.3.5"
	a(1, 3, 6) = "1.3.6"
	a(1, 4, 5) = "1.4.5"
	a(1, 4, 6) = "1.4.6"
	a(2, 3, 5) = "2.3.5"
	a(2, 3, 6) = "2.3.6"
	a(2, 4, 5) = "2.4.5"
	a(2, 4, 6) = "2.4.6"
	
	
	' Display array metadata.
	Debug.Print "a.Rank = " & Idx.Arr_Rank(a)
	Debug.Print
	Debug.Print "a.Length(1) = " & Idx.Arr_Length(a, dimension := 1)
	Debug.Print "a.Length(2) = " & Idx.Arr_Length(a, dimension := 2)
	Debug.Print "a.Length(3) = " & Idx.Arr_Length(a, dimension := 3)
	
	
	' Extract data.
	Debug.Print "a()           = """ & Idx.Arr_Index(a, Array()) & """"
	Debug.Print "a(1)          = """ & Idx.Arr_Index(a, Array(1)) & """"
	Debug.Print "a(1, 4)       = """ & Idx.Arr_Index(a, Array(1, 4)) & """"
	Debug.Print "a(1, 4, 5)    = """ & Idx.Arr_Index(a, Array(1, 4, 5)) & """"
	Debug.Print "a(1, 4, 5, 7) = """ & Idx.Arr_Index(a, Array(1, 4, 5, 7)) & """"
End Sub


Public Sub Test__Index__NestedArray()
	' Nest arrays 3 deep.
	Dim a As Variant: a = Array( _
		Array( _
			Array( _
				"0.0.0", _
				"0.0.1" _
			), _
			Array( _
				"0.1.0", _
				"0.1.1" _
			) _
		), _
		Array( _
			Array( _
				"1.0.0", _
				"1.0.1" _
			), _
			Array( _
				"1.1.0", _
				"1.1.1" _
			) _
		) _
	)
	
	
	' Display array metadata.
	Debug.Print "a.Rank       = " & Idx.Arr_Rank(a)
	Debug.Print "a(0).Rank    = " & Idx.Arr_Rank(a(0))
	Debug.Print "a(0)(0).Rank = " & Idx.Arr_Rank(a(0)(0))
	Debug.Print
	Debug.Print "a.Length       = " & Idx.Arr_Length(a, dimension := 1)
	Debug.Print "a(0).Length    = " & Idx.Arr_Length(a(0), dimension := 1)
	Debug.Print "a(0)(0).Length = " & Idx.Arr_Length(a(0)(0), dimension := 1)
	
	
	' Extract data.
	Debug.Print "a()           = """ & Idx.Index(a, Array()) & """
	Debug.Print "a(0)          = """ & Idx.Index(a, Array(0)) & """
	Debug.Print "a(0)(1)       = """ & Idx.Index(a, Array(0, 1)) & """
	Debug.Print "a(0)(1)(0)    = """ & Idx.Index(a, Array(0, 1, 0)) & """
	Debug.Print "a(0)(1)(0)(1) = """ & Idx.Index(a, Array(0, 1, 0, 1)) & """
End Sub
