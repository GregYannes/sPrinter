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



  [c_spf]:    https://en.cppreference.com/w/c/io/fprintf
  [spf_lang]: https://en.wikipedia.org/wiki/Printf#Other_contexts
  [so_q]:     https://stackoverflow.com/q/17233701
  [proj_mod]: src/sPrinter.bas
  [docs_fmt]: docs/Formatting.md
