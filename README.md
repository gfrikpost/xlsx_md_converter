# XlsxMdConverter

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'xlsx_md_converter'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install xlsx_md_converter

## Usage

run `bin/console`

XLSXMdConverter.xlsx_to_md('file_path.xlsx') - create converted file file_pathTIMESTAMP.md

XLSXMdConverter.md_to_xlsx('file_path.md') - create converted file file_pathTIMESTAMP.xlsx

XLSXMdConverter.xlsx_to_simple_xlsx('file_path.xlsx') - create file file_pathTIMESTAMP.xlsx without format and formulas


## License

The gem is available as open source under the terms of the [MIT License](http://opensource.org/licenses/MIT).

