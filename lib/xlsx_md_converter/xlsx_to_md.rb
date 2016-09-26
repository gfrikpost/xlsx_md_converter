require 'xlsx_md_converter/xlsx_md_converter_base'
require 'reverse_markdown'

class XLSXToMD < XLSXMdConverterBase

  private

  def process
    result = ''
    xlsx_file.each_with_pagename do |name, sheet|
      result << "<strong>#{name}</strong>"
      result << '<table>'
      result << '<tr>'
      result << table_headers(sheet.row(1).count)
      result << '</tr>'
      result << table_body(sheet)
      result << '</table>'
    end

    md_table = ReverseMarkdown.convert(result)
    out_file.puts(md_table)
    out_file.close
    super
  end

  def table_body(sheet)
    result = ''
    sheet.each do |row|
      result << '<tr>'
      row.each do |value|
        result << "<td>#{value}</td>"
      end
      result << '</tr>'
    end
    result
  end

  def table_headers(columns_count)
    headers_string = ''
    (1..columns_count).each do |column|
      headers_string << "<th>#{Roo::Utils.number_to_letter(column)}</th>"
    end
    headers_string
  end

  def out_file
    @out_file ||= File.new(output_file_path, 'w')
  end

  def output_file_extension
    'md'
  end
end
