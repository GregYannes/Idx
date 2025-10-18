Attribute VB_Name = "Idx"



' #############
' ## Options ##
' #############

' Explicitly declare all variables.
Option Explicit

' ' Hide these developer functions from end users in Excel.
' Option Private Module



' ##############
' ## Metadata ##
' ##############

Public Const MOD_NAME As String = "Idx"

' Public Const MOD_VERSION As String = "0.1.0"

Public Const MOD_REPO As String = "https://github.com/GregYannes/Idx"



' ###############
' ## Constants ##
' ###############

' The most dimensions an array may have.  See MS docs: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/too-many-dimensions
Public Const MAX_ARR_RANK As Long = 60

' ' ...
' Public Const MAX_ARR_LENGTH As Long = ...



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
		Debug.Print "ERROR: The indices must have exactly one dimension."
	End If
	
	' Short-circuit for no indices.
	Dim iLen As Long: iLen = Arr_Length(indices, dimension := 1)
	If iLen = 0 Then
		Debug.Print "ERROR: At least one index is required."
	End If
	
	
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
' ' NOTE: This is identical to regular indexing and thus superfluous without tracking existence:
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
	
	' Short-circuit for no indices...
	Dim iLen As Long: iLen = Arr_Length(indices, dimension := 1)
	If iLen = 0 Then
		Debug.Print "ERROR: At least one index is required."
		
	' ...or for impossibly many.
	ElseIf iLen > MAX_ARR_RANK Then
		Debug.Print "ERROR: No array may accept more than " & VBA.CStr(MAX_ARR_RANK) & " indices for " & VBA.CStr(MAX_ARR_RANK) & " dimensions."
	End If
	
	' Short-circuit for uninitialized array.
	Dim aRnk As Long: aRnk = Arr_Rank(arr)
	If aRnk = 0 Then Debug.Print "ERROR: The array must be initialized."
	
	' Short-circuit for wrong number of indices.
	If iLen <> aRnk Then Debug.Print "ERROR: There must be exactly as many indices (" & VBA.CStr(iLen) & ") as dimensions (" & VBA.CStr(aRnk) ") in the array."
	
	
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
)
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
	Case  0: Assign v, a()
	Case  1: Assign v, a(i(l))
	Case  2: Assign v, a(i(l),i(l+1))
	Case  3: Assign v, a(i(l),i(l+1),i(l+2))
	Case  4: Assign v, a(i(l),i(l+1),i(l+2),i(l+3))
	Case  5: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4))
	Case  6: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5))
	Case  7: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6))
	Case  8: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7))
	Case  9: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8))
	Case 10: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9))
	Case 11: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10))
	Case 12: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11))
	Case 13: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12))
	Case 14: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13))
	Case 15: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14))
	Case 16: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15))
	Case 17: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16))
	Case 18: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17))
	Case 19: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18))
	Case 20: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19))
	Case 21: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20))
	Case 22: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21))
	Case 23: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22))
	Case 24: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23))
	Case 25: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24))
	Case 26: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25))
	Case 27: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26))
	Case 28: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27))
	Case 29: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28))
	Case 30: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29))
	Case 31: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30))
	Case 32: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31))
	Case 33: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32))
	Case 34: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33))
	Case 35: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34))
	Case 36: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35))
	Case 37: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36))
	Case 38: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37))
	Case 39: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38))
	Case 40: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39))
	Case 41: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40))
	Case 42: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41))
	Case 43: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41),i(l+42))
	Case 44: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41),i(l+42),i(l+43))
	Case 45: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41),i(l+42),i(l+43),i(l+44))
	Case 46: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41),i(l+42),i(l+43),i(l+44),i(l+45))
	Case 47: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41),i(l+42),i(l+43),i(l+44),i(l+45),i(l+46))
	Case 48: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41),i(l+42),i(l+43),i(l+44),i(l+45),i(l+46),i(l+47))
	Case 49: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41),i(l+42),i(l+43),i(l+44),i(l+45),i(l+46),i(l+47),i(l+48))
	Case 50: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41),i(l+42),i(l+43),i(l+44),i(l+45),i(l+46),i(l+47),i(l+48),i(l+49))
	Case 51: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41),i(l+42),i(l+43),i(l+44),i(l+45),i(l+46),i(l+47),i(l+48),i(l+49),i(l+50))
	Case 52: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41),i(l+42),i(l+43),i(l+44),i(l+45),i(l+46),i(l+47),i(l+48),i(l+49),i(l+50),i(l+51))
	Case 53: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41),i(l+42),i(l+43),i(l+44),i(l+45),i(l+46),i(l+47),i(l+48),i(l+49),i(l+50),i(l+51),i(l+52))
	Case 54: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41),i(l+42),i(l+43),i(l+44),i(l+45),i(l+46),i(l+47),i(l+48),i(l+49),i(l+50),i(l+51),i(l+52),i(l+53))
	Case 55: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41),i(l+42),i(l+43),i(l+44),i(l+45),i(l+46),i(l+47),i(l+48),i(l+49),i(l+50),i(l+51),i(l+52),i(l+53),i(l+54))
	Case 56: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41),i(l+42),i(l+43),i(l+44),i(l+45),i(l+46),i(l+47),i(l+48),i(l+49),i(l+50),i(l+51),i(l+52),i(l+53),i(l+54),i(l+55))
	Case 57: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41),i(l+42),i(l+43),i(l+44),i(l+45),i(l+46),i(l+47),i(l+48),i(l+49),i(l+50),i(l+51),i(l+52),i(l+53),i(l+54),i(l+55),i(l+56))
	Case 58: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41),i(l+42),i(l+43),i(l+44),i(l+45),i(l+46),i(l+47),i(l+48),i(l+49),i(l+50),i(l+51),i(l+52),i(l+53),i(l+54),i(l+55),i(l+56),i(l+57))
	Case 59: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41),i(l+42),i(l+43),i(l+44),i(l+45),i(l+46),i(l+47),i(l+48),i(l+49),i(l+50),i(l+51),i(l+52),i(l+53),i(l+54),i(l+55),i(l+56),i(l+57),i(l+58))
	Case 60: Assign v, a(i(l),i(l+1),i(l+2),i(l+3),i(l+4),i(l+5),i(l+6),i(l+7),i(l+8),i(l+9),i(l+10),i(l+11),i(l+12),i(l+13),i(l+14),i(l+15),i(l+16),i(l+17),i(l+18),i(l+19),i(l+20),i(l+21),i(l+22),i(l+23),i(l+24),i(l+25),i(l+26),i(l+27),i(l+28),i(l+29),i(l+30),i(l+31),i(l+32),i(l+33),i(l+34),i(l+35),i(l+36),i(l+37),i(l+38),i(l+39),i(l+40),i(l+41),i(l+42),i(l+43),i(l+44),i(l+45),i(l+46),i(l+47),i(l+48),i(l+49),i(l+50),i(l+51),i(l+52),i(l+53),i(l+54),i(l+55),i(l+56),i(l+57),i(l+58),i(l+59))
	Case Else: Assign v, a()
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
