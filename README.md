# Idx #

The [**`Idx`**][idx_mod] module lets you easily extract data from nested data structures in VBA.  It is inspired by the [`pluck()`][r_pluck] family in [R][r_lang].


> [!TIP]
> 
> See [here][idx_rls] for the latest (pre)release.


# Usage #

Supply an array of `indices` to [`Arr_Index()`][arr_idx], and it will index into (say) a multidimensional [array][vba_arr].

```vba
' Declare 3D array...
Dim multi(1 To 2, 3 To 4, 5 To 6) As String

' ...and populate it.
multi(1, 3, 5) = "multi(1, 3, 5)"
multi(1, 3, 6) = "multi(1, 3, 6)"
multi(1, 4, 5) = "multi(1, 4, 5)"
multi(1, 4, 6) = "multi(1, 4, 6)"
multi(2, 3, 5) = "multi(2, 3, 5)"
multi(2, 3, 6) = "multi(2, 3, 6)"
multi(2, 4, 5) = "multi(2, 4, 5)"
multi(2, 4, 6) = "multi(2, 4, 6)"


' Index into the array programmatically.
Dim ind() As Variant: ind = Array(1, 4, 5)
Debug.Print Arr_Index(multi, ind)
```

> ```
> multi(1, 4, 5)
> ```

<br>

Do the same for [`Index()`][idx_fun] itself, and it will index into _any_ data structure: the multidimensional array…

```vba
' Index programmatically...
Debug.Print Index(multi, ind)

' ...or manually.
Debug.Print Index0(multi, 1, 4, 5)
```

> ```
> multi(1, 4, 5)
> multi(1, 4, 5)
> ```

<br>

…or a nested array…

```vba
' Create a nested array.
Dim nested() As Variant: nested = Array( _
	Array( _
		Array("nested(0)(0)(0)", "nested(0)(0)(1)"), _
		Array("nested(0)(1)(0)", "nested(0)(1)(1)") _
	), _
	Array( _
		Array("nested(1)(0)(0)", "nested(1)(0)(1)"), _
		Array("nested(1)(1)(0)", "nested(1)(1)(1)") _
	) _
)


' Index manually.
Debug.Print Index0(nested, 1, 0, 1)
```

> ```
> nested(1)(0)(1)
> ```

<br>

…or a [`Collection`][vba_clx].

```vba
' Create a collection...
Dim clx As Collection: Set clx = New Collection

' ...and populate it.
clx.Add "clx(1)"
clx.Add "clx!key_2", key := "key_2"


' Index by position...
Debug.Print Index0(clx, 1)

' ...or by key.
Debug.Print Index0(clx, "key_2")
```

> ```
> clx(1)
> clx!key_2
> ```

<br>

But the _true_ power of `Index()` is seen with nested data structures of diverse types.  Here we "drill down" to the very innermost data and extract it elegantly!

```vba
' Wrap the 3D array within the original collection...
clx.Add multi, key := "key_3"

' ...and nest it at the very end.
nested = Array( _
	Array( _
		Array("nested(0)(0)(0)", "nested(0)(0)(1)"), _
		Array("nested(0)(1)(0)", "nested(0)(1)(1)") _
	), _
	Array( _
		Array("nested(1)(0)(0)", "nested(1)(0)(1)"), _
		Array("nested(1)(1)(0)", "nested(1)(1)(1)", clx) _
	) _
)
'		                                            ^^^
'		                                         Insertion


' Index manually...
Debug.Print Index0(nested, 1, 1, 2, "key_3", 2, 3, 6)

' ...or programmatically.
ind = Array(1, 1, 2, "key_3", 2, 3, 6)
Debug.Print Index(nested, ind)
```

> ```
> multi(2, 3, 6)
> multi(2, 3, 6)
> ```


# API #

Here are all the features provided by **`Idx`**, which are useful in both VBA and Excel[^1].  If you are a developer, and wish to hide these functions from _your_ users in Excel, then look [here][idx_prv] to activate [`Option Private`][vba_prv] for the [`Idx.bas`][idx_mod] module.


## Metadata ##

Describe the module _itself_.

  - [`MOD_NAME`][idx_met]: The name (`String`) of the module.
  - [`MOD_VERSION`][idx_met]: Its current [version][sem_ver] (`String`).
  - [`MOD_REPO`][idx_met]: The URL (`String`) to its repository.


## Array Extraction ##

Index into [arrays][vba_arr], especially multidimensional arrays.

  - [`Arr_Index()`][arr_idx]: Extract an element programmatically with an array of `indices`.


## Arbitrary Extraction ##

Index into arbitrary data.  This may be either an [array][vba_arr] or an object with a [default member][vba_dfl], like a [`Collection`][vba_clx] or [`Dictionary`][vba_dix].  Or it may be some arbitrary nesting of such data structures.

  - [`Index()`][idx_fun]: Extract an element programmatically with an array of `indices`…
  - [`Index0()`][idx_fun]: …and manually with literal indices.


## Utilities ##

Perform broadly useful tasks.

  - [`Assign()`][idx_utl]: Assign any value (scalar or objective) to a variable (by [reference][vba_byr]).
  - [`Arr_Rank()`][idx_utl]: Get the ["rank"][net_rnk] (`Long`) of an [array][vba_arr], which is the count of its [dimensions][net_dim][^2].
  - [`Arr_Length()`][idx_utl]: Get the length (`Long`) of an array.



  [^1]: You may qualify [`Idx.Index()`][idx_fun] to avoid clashes with the [`INDEX()`][xl_idx] function native to Excel.
  [^2]: In VBA an [array][vba_arr] may have at most [60 dimensions][vba_dim].



  [idx_mod]: src/Idx.bas
  [r_pluck]: https://purrr.tidyverse.org/reference/pluck
  [r_lang]:  https://www.r-project.org/about.html
  [idx_rls]: ../../releases/tag/latest
  [arr_idx]: docs/Arr_Index.md
  [vba_arr]: https://learn.microsoft.com/office/vba/language/concepts/getting-started/using-arrays
  [idx_fun]: docs/Index.md
  [vba_clx]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object
  [idx_prv]: src/Idx.bas#L12-L13
  [vba_prv]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/option-private-statement
  [idx_met]: docs/Metadata.md
  [sem_ver]: https://semver.org
  [vba_dfl]: http://www.cpearson.com/excel/DefaultMember.aspx
  [vba_dix]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/dictionary-object
  [idx_utl]: docs/Utilities.md
  [vba_byr]: https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/procedures/passing-arguments-by-value-and-by-reference
  [net_rnk]: https://learn.microsoft.com/dotnet/api/system.array.rank
  [net_dim]: https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/arrays/array-dimensions
  [xl_idx]:  https://support.microsoft.com/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd
  [vba_dim]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/too-many-dimensions
