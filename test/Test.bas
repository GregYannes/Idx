Attribute VB_Name = "Test"



' Run all tests.
Public Sub Test()
	' ##########
	' ## Data ##
	' ##########
	
	' Declare an uninitialized array.
	Dim uninitArr() As String
	
	
	' Initialize an empty array.
	Dim emptyArr As Variant: emptyArr = Array()
	
	
	' Populate a 3D array.
	Dim multiArr(1 To 2, 3 To 4, 5 To 6) As String
	multiArr(1, 3, 5) = "multiArr(1, 3, 5)"
	multiArr(1, 3, 6) = "multiArr(1, 3, 6)"
	multiArr(1, 4, 5) = "multiArr(1, 4, 5)"
	multiArr(1, 4, 6) = "multiArr(1, 4, 6)"
	multiArr(2, 3, 5) = "multiArr(2, 3, 5)"
	multiArr(2, 3, 6) = "multiArr(2, 3, 6)"
	multiArr(2, 4, 5) = "multiArr(2, 4, 5)"
	multiArr(2, 4, 6) = "multiArr(2, 4, 6)"
	
	
	' Populate a collection.
	Dim clx As Collection: Set clx = New Collection
	clx.Add "clx(1)"
	clx.Add "clx!key_2", key := "key_2"
	
	
	' Nest arrays 3 deep...with the collection too.
	Dim nestArr As Variant: nestArr = Array( _
		Array( _
			Array( _
				"nestArr(0)(0)(0)", _
				"nestArr(0)(0)(1)" _
			), _
			Array( _
				"nestArr(0)(1)(0)", _
				"nestArr(0)(1)(1)" _
			) _
		), _
		Array( _
			Array( _
				"nestArr(1)(0)(0)", _
				"nestArr(1)(0)(1)" _
			), _
			Array( _
				"nestArr(1)(1)(0)", _
				"nestArr(1)(1)(1)", _
				clx _
			) _
		) _
	)
	
	
	' Assemble complex data even deeper.
	Dim complexData As Variant: complexData = nestArr
	clx.Add multiArr, key := "key_3"
	
	
	
	' #############
	' ## Testing ##
	' #############
	
	Debug.Print "#################"
	Debug.Print "## Arr_Index() ##"
	Debug.Print "#################"
	Debug.Print
	Test__Arr_Index uninitArr := uninitArr, emptyArr := emptyArr, multiArr := multiArr
	
	Debug.Print
	Debug.Print
	Debug.Print
	
	Debug.Print "#############"
	Debug.Print "## Index() ##"
	Debug.Print "#############"
	Debug.Print
	Test__Index uninitArr := uninitArr, emptyArr := emptyArr, multiArr := multiArr, nestArr := nestArr, clx := clx, complexData := complexData
End Sub



' #################
' ## Arr_Index() ##
' #################

' Run all tests on Arr_Index().
Public Sub Test__Arr_Index( _
	ByRef uninitArr As Variant, _
	ByRef emptyArr As Variant, _
	ByRef multiArr As Variant _
)
	Debug.Print "#######################################"
	Debug.Print "## Arr_Index() | Uninitialized Array ##"
	Debug.Print "#######################################"
	Debug.Print
	Test__Arr_Index__UninitArray arr := uninitArr
	
	Debug.Print
	Debug.Print
	
	Debug.Print "###############################"
	Debug.Print "## Arr_Index() | Empty Array ##"
	Debug.Print "###############################"
	Debug.Print
	Test__Arr_Index__EmptyArray arr := emptyArr
	
	Debug.Print
	Debug.Print
	
	Debug.Print "##########################################"
	Debug.Print "## Arr_Index() | Multidimensional Array ##"
	Debug.Print "##########################################"
	Debug.Print
	Test__Arr_Index__MultiArray arr := multiArr
End Sub


' Test Arr_Index() on an uninitialized array.
Public Sub Test__Arr_Index__UninitArray(ByRef arr As Variant)
	' ' Declare uninitialized array.
	' Dim arr() As String
	
	
	' Display array metadata.
	Debug.Print "arr.Rank = " & Idx.Arr_Rank(arr)
	Debug.Print
	Debug.Print "arr.Length(1) = " & Idx.Arr_Length(arr, dimension := 1)
	' Debug.Print "arr.Length(2) = " & Idx.Arr_Length(arr, dimension := 2)
	' Debug.Print "arr.Length(3) = " & Idx.Arr_Length(arr, dimension := 3)
	
	
	' Extract data.
	Debug.Print "arr()           = " & Idx.Arr_Index(arr, Array())
	Debug.Print "arr(1)          = " & Idx.Arr_Index(arr, Array(1))
	' Debug.Print "arr(1, 4)       = " & Idx.Arr_Index(arr, Array(1, 4))
	' Debug.Print "arr(1, 4, 5)    = " & Idx.Arr_Index(arr, Array(1, 4, 5))
	' Debug.Print "arr(1, 4, 5, 7) = " & Idx.Arr_Index(arr, Array(1, 4, 5, 7))
End Sub


' Test Arr_Index() on an empty array.
Public Sub Test__Arr_Index__EmptyArray(ByRef arr As Variant)
	' ' Create empty array.
	' Dim arr As Variant: arr = Array()
	
	
	' Display array metadata.
	Debug.Print "arr.Rank = " & Idx.Arr_Rank(arr)
	Debug.Print
	Debug.Print "arr.Length(1) = " & Idx.Arr_Length(arr, dimension := 1)
	Debug.Print "arr.Length(2) = " & Idx.Arr_Length(arr, dimension := 2)
	' Debug.Print "arr.Length(3) = " & Idx.Arr_Length(arr, dimension := 3)
	
	
	' Extract data.
	Debug.Print "arr()           = " & Idx.Arr_Index(arr, Array())
	Debug.Print "arr(1)          = " & Idx.Arr_Index(arr, Array(1))
	Debug.Print "arr(1, 4)       = " & Idx.Arr_Index(arr, Array(1, 4))
	' Debug.Print "arr(1, 4, 5)    = " & Idx.Arr_Index(arr, Array(1, 4, 5))
	' Debug.Print "arr(1, 4, 5, 7) = " & Idx.Arr_Index(arr, Array(1, 4, 5, 7))
End Sub


' Test Arr_Index() on a multidimensional array.
Public Sub Test__Arr_Index__MultiArray(ByRef arr As Variant)
	' Display array metadata.
	Debug.Print "arr.Rank = " & Idx.Arr_Rank(arr)
	Debug.Print
	Debug.Print "arr.Length(1) = " & Idx.Arr_Length(arr, dimension := 1)
	Debug.Print "arr.Length(2) = " & Idx.Arr_Length(arr, dimension := 2)
	Debug.Print "arr.Length(3) = " & Idx.Arr_Length(arr, dimension := 3)
	
	
	' Extract data.
	Debug.Print "arr()           = " & Idx.Arr_Index(arr, Array())
	Debug.Print "arr(1)          = " & Idx.Arr_Index(arr, Array(1))
	Debug.Print "arr(1, 4)       = " & Idx.Arr_Index(arr, Array(1, 4))
	Debug.Print "arr(1, 4, 5)    = " & Idx.Arr_Index(arr, Array(1, 4, 5))
	Debug.Print "arr(1, 4, 5, 7) = " & Idx.Arr_Index(arr, Array(1, 4, 5, 7))
End Sub



' #############
' ## Index() ##
' #############

' Run all tests on Index().
Public Sub Test__Index( _
	ByRef uninitArr As Variant, _
	ByRef emptyArr As Variant, _
	ByRef multiArr As Variant, _
	ByRef nestArr As Variant, _
	ByRef clx As Collection, _
	ByRef complexData As Variant _
)
	Debug.Print "###################################"
	Debug.Print "## Index() | Uninitialized Array ##"
	Debug.Print "###################################"
	Debug.Print
	Test__Index__UninitArray arr := uninitArr
	
	Debug.Print
	Debug.Print
	
	Debug.Print "###########################"
	Debug.Print "## Index() | Empty Array ##"
	Debug.Print "###########################"
	Debug.Print
	Test__Index__EmptyArray arr := emptyArr
	
	Debug.Print
	Debug.Print
	
	Debug.Print "######################################"
	Debug.Print "## Index() | Multidimensional Array ##"
	Debug.Print "######################################"
	Debug.Print
	Test__Index__MultiArray arr := multiArr
	
	Debug.Print
	Debug.Print
	
	Debug.Print "############################"
	Debug.Print "## Index() | Nested Array ##"
	Debug.Print "############################"
	Debug.Print
	Test__Index__NestedArray arr := nestArr
	
	Debug.Print
	Debug.Print
	
	Debug.Print "##########################"
	Debug.Print "## Index() | Collection ##"
	Debug.Print "##########################"
	Debug.Print
	Test__Index__Collection clx := clx
	
	Debug.Print
	Debug.Print
	
	Debug.Print "############################"
	Debug.Print "## Index() | Complex Data ##"
	Debug.Print "############################"
	Debug.Print
	Test__Index__Complex data := complexData
End Sub


' Test Index() on an uninitialized array.
Public Sub Test__Index__UninitArray(ByRef arr As Variant)
	' ' Declare uninitialized array.
	' Dim arr() As String
	
	
	' Display array metadata.
	Debug.Print "arr.Rank = " & Idx.Arr_Rank(arr)
	Debug.Print
	Debug.Print "arr.Length(1) = " & Idx.Arr_Length(arr, dimension := 1)
	' Debug.Print "arr.Length(2) = " & Idx.Arr_Length(arr, dimension := 2)
	' Debug.Print "arr.Length(3) = " & Idx.Arr_Length(arr, dimension := 3)
	
	
	' Extract data.
	Debug.Print "arr()           = " & Idx.Index(arr, Array())
	Debug.Print "arr(1)          = " & Idx.Index(arr, Array(1))
	' Debug.Print "arr(1, 4)       = " & Idx.Index(arr, Array(1, 4))
	' Debug.Print "arr(1, 4, 5)    = " & Idx.Index(arr, Array(1, 4, 5))
	' Debug.Print "arr(1, 4, 5, 7) = " & Idx.Index(arr, Array(1, 4, 5, 7))
End Sub


' Test Index() on an empty array.
Public Sub Test__Index__EmptyArray(ByRef arr As Variant)
	' ' Create empty array.
	' Dim arr As Variant: arr = Array()
	
	
	' Display array metadata.
	Debug.Print "arr.Rank = " & Idx.Arr_Rank(arr)
	Debug.Print
	Debug.Print "arr.Length(1) = " & Idx.Arr_Length(arr, dimension := 1)
	Debug.Print "arr.Length(2) = " & Idx.Arr_Length(arr, dimension := 2)
	' Debug.Print "arr.Length(3) = " & Idx.Arr_Length(arr, dimension := 3)
	
	
	' Extract data.
	Debug.Print "arr()           = " & Idx.Index(arr, Array())
	Debug.Print "arr(1)          = " & Idx.Index(arr, Array(1))
	Debug.Print "arr(1, 4)       = " & Idx.Index(arr, Array(1, 4))
	' Debug.Print "arr(1, 4, 5)    = " & Idx.Index(arr, Array(1, 4, 5))
	' Debug.Print "arr(1, 4, 5, 7) = " & Idx.Index(arr, Array(1, 4, 5, 7))
End Sub


' Test Index() on a multidimensional array.
Public Sub Test__Index__MultiArray(ByRef arr As Variant)
	' Display array metadata.
	Debug.Print "arr.Rank = " & Idx.Arr_Rank(arr)
	Debug.Print
	Debug.Print "arr.Length(1) = " & Idx.Arr_Length(arr, dimension := 1)
	Debug.Print "arr.Length(2) = " & Idx.Arr_Length(arr, dimension := 2)
	Debug.Print "arr.Length(3) = " & Idx.Arr_Length(arr, dimension := 3)
	
	
	' Extract data.
	Debug.Print "arr()           = " & Idx.Index(arr, Array())
	Debug.Print "arr(1)          = " & Idx.Index(arr, Array(1))
	Debug.Print "arr(1, 4)       = " & Idx.Index(arr, Array(1, 4))
	Debug.Print "arr(1, 4, 5)    = " & Idx.Index(arr, Array(1, 4, 5))
	Debug.Print "arr(1, 4, 5, 7) = " & Idx.Index(arr, Array(1, 4, 5, 7))
End Sub


' Test Index() on a nested array.
Public Sub Test__Index__NestedArray(ByRef arr As Variant)
	' Display array metadata.
	Debug.Print "arr.Rank       = " & Idx.Arr_Rank(arr)
	Debug.Print "arr(0).Rank    = " & Idx.Arr_Rank(arr(0))
	Debug.Print "arr(0)(0).Rank = " & Idx.Arr_Rank(arr(0)(0))
	Debug.Print
	Debug.Print "arr.Length       = " & Idx.Arr_Length(arr, dimension := 1)
	Debug.Print "arr(0).Length    = " & Idx.Arr_Length(arr(0), dimension := 1)
	Debug.Print "arr(0)(0).Length = " & Idx.Arr_Length(arr(0)(0), dimension := 1)
	
	
	' Extract data.
	Debug.Print "arr()           = " & Idx.Index(arr, Array())
	Debug.Print "arr(0)          = " & Idx.Index(arr, Array(0))
	Debug.Print "arr(0)(1)       = " & Idx.Index(arr, Array(0, 1))
	Debug.Print "arr(0)(1)(0)    = " & Idx.Index(arr, Array(0, 1, 0))
	Debug.Print "arr(0)(1)(0)(1) = " & Idx.Index(arr, Array(0, 1, 0, 1))
End Sub


' Test Index() on a Collection.
Public Sub Test__Index__Collection(ByRef clx As Collection)
	' Extract data.
	Debug.Print "clx()     = " & Idx.Index(clx, Array())
	Debug.Print "clx(0)    = " & Idx.Index(clx, Array(0))
	Debug.Print "clx(1)    = " & Idx.Index(clx, Array(1))
	Debug.Print "clx(2)    = " & Idx.Index(clx, Array(2))
	Debug.Print "clx!key_2 = " & Idx.Index(clx, Array("key_2"))
	Debug.Print "clx(1, 1) = " & Idx.Index(clx, Array(1, 1))
End Sub


' Test Index() on arbitrarily complex data.
Public Sub Test__Index__Complex(ByRef data As Variant)
	' Extract data.
	Debug.Print "data.Type                    = " & VBA.TypeName(data)
	Debug.Print "data()                       = " & Idx.Index(data, Array())
	Debug.Print "data(1)(1)(1)                = " & Idx.Index(data, Array(1, 1, 1))
	Debug.Print "data(1)(1)(2).Type           = " & VBA.TypeName(Idx.Index(data, Array(1, 1, 2)))
	Debug.Print "data(1)(1)(2)(1)             = " & Idx.Index(data, Array(1, 1, 2, 1))
	Debug.Print "data(1)(1)(2)(2)             = " & Idx.Index(data, Array(1, 1, 2, 2))
	Debug.Print "data(1)(1)(2)!key_2          = " & Idx.Index(data, Array(1, 1, 2, "key_2"))
	Debug.Print "data(1)(1)(2)!key_3.Type     = " & VBA.TypeName(Idx.Index(data, Array(1, 1, 2, "key_3"))
	Debug.Print "data(1)(1)(2)!key_3(1, 4, 5) = " & Idx.Index(data, Array(1, 1, 2, "key_3", 1, 4, 5))
End Sub
