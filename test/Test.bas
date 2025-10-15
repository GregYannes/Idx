Attribute VB_Name = "Test"



Public Sub Test__Arr_Index__3dArray()
	Dim a(1 To 2, 3 To 4, 5 To 6) As String
	a(1, 3, 5) = "1.3.5"
	a(1, 3, 6) = "1.3.6"
	a(1, 4, 5) = "1.4.5"
	a(1, 4, 6) = "1.4.6"
	a(2, 3, 5) = "2.3.5"
	a(2, 3, 6) = "2.3.6"
	a(2, 4, 5) = "2.4.5"
	a(2, 4, 6) = "2.4.6"
	
	Debug.Print "a.Rank = " & Idx.Arr_Rank(Array())
	Debug.Print "a.Length = " & Idx.Arr_Length(Array())
	
	Debug.Print "a() = """ & Idx.Arr_Index(a, Array()) & """"
	Debug.Print "a(1, 4) = """ & Idx.Arr_Index(a, Array(1, 4)) & """"
	Debug.Print "a(1, 4, 5) = """ & Idx.Arr_Index(a, Array(1, 4, 5)) & """"
	Debug.Print "a(1, 4, 5, 7) = """ & Idx.Arr_Index(a, Array(1, 4, 5, 7)) & """"
End Sub


Public Sub Test__Index__NestedArray()
	' ...
End Sub
