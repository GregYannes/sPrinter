# sPrinter #

The [`sprintf()`][c_spf] family is popular across [many programming languages][spf_lang] for conveniently displaying information in an attractive way.  Despite [vocal demand][so_q], neither VBA nor Excel support such features—until now.

<br>

Introducing the [**`sPrinter`**][proj_mod] module for Excel VBA!  Simply write a template for your message, and use curly braces `{…}` to embed data inside.

```vba
vPrint "You have a meeting with {1} {2} at {3} on {4}.", "John", "Doe", Time(), Date()
'       ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^   ^^^^^^  ^^^^^  ^^^^^^  ^^^^^^
'                      Message Template                               Data
```

> ```
> You have a meeting with John Doe at 1:30:00_PM on 1/1/26.
> ```

<br>

You can make this look even nicer, by applying [format codes][docs_fmt] to your data!

```vba
vPrint "You have a meeting with {1} {2} at {3:h:MM AM/PM} on {4:dddd, mmmm d}.", "John", "Doe", Time(), Date()
'                                             ^^^^^^^^^^        ^^^^^^^^^^^^
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



  [^fmt_code]: You may choose between the codes used by [`Format()`][vba_fmt] in VBA, or those used by [`TEXT()`][xl_txt] in Excel.
  [^fmt_name]: This may be a key in a [`Collection`][vba_clx] or [`Dictionary`][vba_dix] supplied to [`xMessage()`][docs_msg] in VBA; or a name paired with a value and supplied to [`iMessage()`][docs_msg].



  [c_spf]:    https://en.cppreference.com/w/c/io/fprintf
  [spf_lang]: https://en.wikipedia.org/wiki/Printf#Other_contexts
  [so_q]:     https://stackoverflow.com/q/17233701
  [proj_mod]: src/sPrinter.bas
  [docs_fmt]: docs/Formatting.md
  [docs_stx]: docs/Syntax.md
  [vba_fmt]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/format-function-visual-basic-for-applications#remarks
  [xl_txt]:   https://support.microsoft.com/office/text-function-20d5ac4d-7b94-49fd-bb38-93d29371225c#ID0EDJ
  [vba_clx]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/collection-object
  [vba_dix]:  https://learn.microsoft.com/office/vba/language/reference/user-interface-help/dictionary-object
  [docs_msg]: docs/Messaging.md
