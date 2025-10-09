Attribute VB_Name = "Idx"



' #########
' ## API ##
' #########

' Manually extract a value (by index) from an arbitrary data structure.
Public Function Index0(ByRef x As Variant, _
	ParamArray indices() As Variant _
) As Variant
	Dim i As Variant: i = indices
	Assign Index0, Index(x, indices := i)
End Function


' Programmatically extract a value (by index) from an arbitrary data structure.
Public Function Index(ByRef x As Variant, _
	ByRef indices As Variant _
) As Variant
	' ################
	' ## Validation ##
	' ################
	
	' Short-circuit for no array.
	If Not VBA.IsArray(indices) Then Debug.Print "ERROR: The indices must be an array."
	
	' Short-circuit for uninitialized...
	Dim iRnk As Long: iRnk = Arr_Rank(indices)
	If iRnk = 0 Then
		Debug.Print "ERROR: The indices must be initialized."
		
	' ...or multidimensional indices.
	ElseIf iRnk > 1 Then
		Debug.Print "ERROR: The indices must have one dimension."
	End If
	
	' Short-circuit for no indices...
	' Dim iLen As Long: iLen = Arr_Length(indices, dimension := 1)
	' If iLen = 0 Then
	' 	Debug.Print "ERROR: At least one index is required."
	' End If
	
	
	' ################
	' ## Extraction ##
	' ################
	
	' Record the bounds of the indices.
	Dim low As Long: low = LBound(indices, 1)
	Dim up As Long: up = UBound(indices, 1)
	
	' Index dynamically into the structure and extract the value there.
	IndexRaw x, v := Index, i := indices, l := low, u := up
End Function


' ' Manually extract a value (by index) from a multidimensional array.
' ' 
' ' NOTE: This is identical to regular indexing and thus unnecessary:
' ' 	Arr_Index0(arr, 1, 2, 3, ...)
' ' 	           arr( 1, 2, 3, ...)
' Public Function Arr_Index0(ByRef arr As Variant, _
' 	ParamArray indices() As Variant _
) As Variant
' 	' Dim exists As Boolean
' 	Dim i As Variant: i = indices
' 	Assign Arr_Index0, Arr_Index(arr, indices := i)  ' exists := exists
' End Function


' Programmatically extract a value (by index) from a multidimensional array.
Public Function Arr_Index(ByRef arr As Variant, _
	ByRef indices As Variant _
) As Variant
' 	Optional ByRef exists As Boolean
	
	' See MS docs here for limit of 60: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/too-many-dimensions
	Const MAX_RANK As Long = 60
	' Const MAX_LENGTH As Long
	
	
	' ################
	' ## Validation ##
	' ################
	
	' Short-circuit for no arrays.
	If Not VBA.IsArray(arr) Then Debug.Print "ERROR: Input must be an array."
	If Not VBA.IsArray(indices) Then Debug.Print "ERROR: The indices must be an array."
	
	' Short-circuit for uninitialized...
	Dim iRnk As Long: iRnk = Arr_Rank(indices)
	If iRnk = 0 Then
		Debug.Print "ERROR: The indices must be initialized."
		
	' ...or multidimensional indices.
	ElseIf iRnk > 1 Then
		Debug.Print "ERROR: The indices must have exactly one dimension."
	End If
	
	' Short-circuit for impossibly many indices...
	Dim iLen As Long: iLen = Arr_Length(indices, dimension := 1)
	If iLen > MAX_RANK Then Debug.Print "ERROR: No array may accept more than " & VBA.CStr(MAX_RANK) & " indices for " & VBA.CStr(MAX_RANK) & " dimensions."
	
	' ' ...or for none.
	' If iLen = 0 Then
	' 	Debug.Print "ERROR: At least one index is required."
	' End If
	
	' Short-circuit for uninitialized array.
	Dim aRnk As Long: aRnk = Arr_Rank(arr)
	If aRnk = 0 Then Debug.Print "ERROR: The array must be initialized."
	
	' Short-circuit for wrong number of indices.
	If iLen <> aRnk Then Debug.Print "ERROR: There must be exactly as many indices as dimensions (" & VBA.CStr(aRnk) ") in the array.
	
	
	' ################
	' ## Extraction ##
	' ################
	
	' Record the bounds of the indices.
	Dim low As Long: low = LBound(indices, 1)
	Dim up As Long: up = UBound(indices, 1)
	
' 	On Error GoTo INDEX_ERROR
	' Index dynamically into the array and extract the value there.
	Arr_IndexRaw arr, v := Arr_Index, i := indices, l := low, u := up
	
' 	exists = True
' 	Exit Function
' 	
' INDEX_ERROR:
' 	exists = False
' 	Arr_Index = Empty  ' VBA.CVErr(VBA.Err.Number)
End Function



' #############
' ## Support ##
' #############

' Lean workhorse for recursive extraction from arbitrary data structures.
Private Sub IndexRaw(ByRef x As Variant, _
	ByRef v As Variant, _
	ByRef i As Variant, _
	ByVal l As Long, _
	ByVal u As Long _
) As Variant
	' Recursive case: several indices left.
	If l < u Then
		' Index into an array on all its dimensions.
		If VBA.IsArray(x) Then
			' Short circuit for uninitialized array.
			Dim r As Long: r = Arr_Rank(x)
			If r = 0 Then Debug.Print "ERROR: Uninitialized array."
			
			' Demarcate the indices for this array.
			Dim k As Long: k = l + r - 1
			
			' Short-circuit for too few indices.
			If k > u Then Debug.Print "ERROR: Too few indices for array."
			
			' Extract the value from the array at those indices.
			Arr_IndexRaw x, v := v, i := i, l := l, u := k
			
			' Index (recursively) into that value using any further indices.
			If k < u Then
				IndexRaw v, v := v, i := i, l := k + 1, u := u
			End If
			
		' Index singly into any other structure.
		Else
			IndexRaw x(i(l)), v := v, i := i, l := l + 1, u := u
		End If
		
	' Base case: only one index left.
	Else
		Assign v, x(i(l))
	End If
End Sub


' Lean workhorse for extraction from multidimensional arrays.
Private Sub Arr_IndexRaw(ByRef a As Variant, _
	ByRef v As Variant, _
	ByRef i As Variant, _
	ByVal l As Long, _
	ByVal u As Long _
)
	Dim n As Long: n = u - l + 1
	
	Select Case n
	Case  1: Assign v, a(i(l))
	Case  2: Assign v, a(i(l),i(l+1))
	Case  3: Assign v, a(i(l),i(l+1),i(l+2))
	' ...
	Case 60: Assign v, a(i(l),i(l+1),i(l+2),…,i(l+59))
	End Select
End Sub



' ###############
' ## Utilities ##
' ###############

' Assign a value (scalar or objective) to a variable.
Public Sub Assign( _
	ByRef var As Variant, _
	ByVal val As Variant _
)
	If VBA.IsObject(val) Then
		Set var = val
	Else
		var = val
	End If
End Sub


' Get the length (along a dimension) of an array.
Public Function Arr_Length(ByRef arr As Variant, _
	Optional ByVal dimension As Long = 1 _
) As Long
	On Error GoTo BOUND_ERROR
	Arr_Length = UBound(arr, dimension) - LBound(arr, dimension) + 1
	Exit Function
	
BOUND_ERROR:
	Arr_Length = 0
End Function


' Get the "rank" of an array: the count of its dimensions.
Public Function Arr_Rank(ByRef arr As Variant) As Long
	Dim tst As Long
	Arr_Rank = 0
	
	On Error GoTo BOUND_ERROR
	Do While True
		Arr_Rank = Arr_Rank + 1
		tst = UBound(arr, Arr_Rank)
	Loop
	
BOUND_ERROR:
	Arr_Rank = Arr_Rank - 1
End Function
