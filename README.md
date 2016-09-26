# XlsxMdConverter

## Installation

Pull gem:

    $ git pull http://github.com/gfrikpost/xlsx_md_converter

Go to the folder:
     $ cd xlsx_md_converter

And then execute:

    $ bundle

Run IRB:

    $ bin/console


## Usage

```ruby
XLSXMdConverter.xlsx_to_md('file_path.xlsx') # => create converted file file_pathTIMESTAMP.md
```

```ruby
XLSXMdConverter.md_to_xlsx('file_path.md') # => create converted file file_pathTIMESTAMP.xlsx
```

```ruby
XLSXMdConverter.xlsx_to_simple_xlsx('file_path.xlsx') # => create file file_pathTIMESTAMP.xlsx without format and formulas
```


## License

The gem is available as open source under the terms of the [MIT License](http://opensource.org/licenses/MIT).

