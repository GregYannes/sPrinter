# sPrinter #

The [`sprintf()`][c_spf] family is popular across [many programming languages][spf_lang] for conveniently displaying information in an attractive way.  Despite [vocal demand][so_q], neither VBA nor Excel support such features—until now.

<br>

Introducing the [**`sPrinter`**][proj_mod] module for Excel VBA!  Simply write a template for your message, and use curly braces `{…}` to embed data inside.

```vba
vPrint "You have a meeting with {1} {2} at {3} on {4}.", "John", "Doe", Time(), Date()
'       ˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄˄   ˄˄˄˄˄˄  ˄˄˄˄˄  ˄˄˄˄˄˄  ˄˄˄˄˄˄
'                      Message Template                               Data
```

> ```
> You have a meeting with John Doe at 1:30:00_PM on 1/1/26.
> ```

<br>

You can make this look even nicer, by applying [format codes][docs_fmt] to your data!

```vba
vPrint "You have a meeting with {1} {2} at {3:h:MM AM/PM} on {4:dddd, mmmm d}.", "John", "Doe", Time(), Date()
'                                             ˄˄˄˄˄˄˄˄˄˄        ˄˄˄˄˄˄˄˄˄˄˄˄
'                                             Time Format        Date Format
```

> ```
> You have a meeting with John Doe at 1:30 PM on Thursday, January 1.
> ```


# Syntax #

See [here][docs_stx] for a detailed guide to syntax for message templates.  You may construct a template from these components:

  - **Field**: Use curly braces `{…}` to embed a data field in your message.
  - **Plaintext**: Everything outside a field is displayed verbatim as regular text.


> [!TIP]
> 
> To display a regular curly brace as plaintext, simply neutralize it with a backslash: `\{`

<br>

You may fine-tune a field by specifying these things:

  - **Index**: Identify the value in the data, either by position or by name.
  - **Format**: Adjust how the value is displayed, using a [format code][docs_fmt] like `m/d/yyyy` that is native[^fmt_code] to Excel/VBA.

```vba
'  Index              Name
'    ˅             ˅˅˅˅˅˅˅˅˅˅
    {2:m/d/yyyy}  {"birthday":m/d/yyyy}
'      ˄˄˄˄˄˄˄˄               ˄˄˄˄˄˄˄˄
'       Format                 Format
```

<br>

You may specify the index in several ways:

  - **Position**: The (numeric) _location_ of the value within the data.  So `1` is the first value, and `2` is the second.
    
    ```
    {1} and {2}
    ```
    
    Use a _negative_ number to count from the _end_.  So `-1` is the last value, and `-2` is the second-to-last.
    
    ```
    {-2} and {-1}
    ```
    
  - **Name**: The (textual) _name_ of the value[^fmt_name] within the data.  You must wrap this in quotes like `"birthday"` or in further braces like `{birthday}`.
    
    ```
    {"birthday"} or {{birthday}}
    ```
  
  - **Auto**: Simply omit the index altogether, and it uses the _next available_ value.
    
    ```
    {} and {}
    ```


> [!IMPORTANT]
> 
> If you use braces inside your index or format, without `\` to neutralize them, then you must keep them _balanced_.  Every active `{` must (eventually) be followed by `}`.


# API #

Here are the features provided by **`sPrinter`**, which are useful in VBA and (mostly) in Excel.  If you are a developer, and wish to hide these functions from _your_ users in Excel, then look [here][mod_prv] to activate [`Option Private`][vba_prv] for the [`sPrinter.bas`][proj_mod] module.


## Metadata ##

Describe the module _itself_.

  - [`MOD_NAME`][docs_met]: The name of the module.
  - [`MOD_VERSION`][docs_met]: Its current [version][sem_ver].
  - [`MOD_REPO`][docs_met]: The URL to its repository.


## Messaging ##

Generate messages by embedding data in a template…

  - [`xMessage()`][docs_msg]: Source fle<ins>**x**</ins>ible `data` from any structure, like an [array][vba_arr] or [`Collection`][vba_clx] or [`Dictionary`][vba_dix].
  - [`vMessage()`][docs_msg]: Supply anonymous <ins>**v**</ins>alues as single arguments…
  - [`iMessage()`][docs_msg]: …or <ins>**i**</ins>dentify them with name-value pairs.

…and print them to the [console][vba_cnsl].
  
  - [`xPrint()`][docs_prn]: Print the output from [`xMessage()`][docs_msg]…
  - [`vPrint()`][docs_prn]: …or from [`vMessage()`][docs_msg]…
  - [`iPrint()`][docs_prn]: …or from [`iMessage()`][docs_msg].


## Parsing ##

Break down a template into an array of its components.

  - [`Parse()`][docs_pse]: Translate the textual template into a [`ParserElement`][docs_elm] array.


## Utilities ##

Perform broadly useful tasks.

  - [`Assign()`][docs_utl]: Assign any value (scalar or objective) to a variable by [reference][vba_byrf].
  
  <br>
  
  - [`Enum_Has()`][docs_utl]: Test if an [`Enum`][vba_enm]eration combo[^enm_comb] contains one of multiple options.
  
  <br>
  
  - [`Num_Cardinal()`][docs_utl]: Represent a number as a [cardinal][card_num] like `1,234`…
  - [`Num_Ordinal()`][docs_utl]: …and as an [ordinal][ord_num] like `1,234th`.
  
  <br>
  
  - [`Arr_Rank()`][docs_utl]: Count[^max_rank] the [dimensions][vba_dimn] of an [array][vba_arr].
  - [`Arr_Length()`][docs_utl]: Get the length of an array.
  
  <br>
  
  - [`ChrX()`][docs_utl]: Safely get the character[^chr_envr] for a code, regardless of platform.
  - [`Txt_Crop()`][docs_utl]: Remove a fixed number of characters from the end(s) of a `String`.
  - [`Txt_List()`][docs_utl]: Format a set of `String`s as a bulleted list.



  [^fmt_code]: You may choose between the codes used by [`Format()`][vba_fmt] in VBA, or those used by [`TEXT()`][xl_txt] in Excel.
  [^fmt_name]: This may be a key in a [`Collection`][vba_clx] or [`Dictionary`][vba_dix] supplied to [`xMessage()`][docs_msg] in VBA; or a name paired with a value and supplied to [`iMessage()`][docs_msg].
  [^enm_comb]: With a [bitwise `Enum`eration][vba_enm2] like [`VbMsgBoxStyle`][vba_mbox], you may layer multiple options using the `+` and `Or` operators.
  [^max_rank]: In VBA an [array][vba_arr] may have at most [60 dimensions][vba_rnk].
  [^chr_envr]: The native [`Chr*()`][vba_chr] family does not reliably support Unicode on Mac.



  [c_spf]:    https://en.cppreference.com/w/c/io/fprintf
  [spf_lang]: https://en.wikipedia.org/wiki/Printf#Other_contexts
  [so_q]:     https://stackoverflow.com/q/17233701
  [proj_mod]: src/sPrinter.bas
  [docs_fmt]: docs/Formatting.md
  [docs_stx]: docs/Syntax.md
  [mod_prv]:  src/sPrinter.bas#L13-L14
  [vba_prv]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/option-private-statement
  [docs_met]: docs/Metadata.md
  [sem_ver]:  https://semver.org
  [docs_msg]: docs/Messaging.md
  [vba_arr]:  https://learn.microsoft.com/office/vba/language/concepts/getting-started/using-arrays
  [vba_clx]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object
  [vba_dix]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/dictionary-object
  [vba_cnsl]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/immediate-window
  [docs_prn]: docs/Printing.md
  [docs_pse]: docs/Parsing.md
  [docs_elm]: docs/Elements.md
  [docs_utl]: docs/Utilities.md
  [vba_byrf]: https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/procedures/passing-arguments-by-value-and-by-reference
  [vba_enm]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/enum-statement
  [card_num]: https://en.wikipedia.org/wiki/Cardinal_numeral
  [ord_num]:  https://en.wikipedia.org/wiki/Ordinal_numeral
  [vba_dimn]: https://learn.microsoft.com/dotnet/visual-basic/programming-guide/language-features/arrays/array-dimensions
  [vba_fmt]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/format-function-visual-basic-for-applications#remarks
  [xl_txt]:   https://support.microsoft.com/office/text-function-20d5ac4d-7b94-49fd-bb38-93d29371225c#ID0EDJ
  [vba_enm2]: https://www.codestack.net/visual-basic/data-structures/enumerators#flag-enumerator-multiple-options
  [vba_mbox]: https://learn.microsoft.com/office/vba/language/reference/user-interface-help/msgbox-function#settings
  [vba_rnk]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/too-many-dimensions
  [vba_chr]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/chr-function#remarks
