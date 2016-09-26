require 'nokogiri'
require 'axlsx'
require 'redcarpet'

class MDToXLSX < XLSXMdConverterBase

  private

  def process
    markdown = Redcarpet::Markdown.new(Redcarpet::Render::HTML, autolink: true, tables: true)
    html = markdown.render(md_content)
    doc = Nokogiri::HTML(html)

    Axlsx::Package.new do |p|
      doc.css('table').each_with_index do |table, index|
        sheet_name = doc.css('strong')[index].text
        p.workbook.add_worksheet(name: sheet_name) do |sheet|
          table.css('tbody tr').each do |tr|
            row = []
            tr.css('td').each do |td|
              row << td.text
            end
            sheet.add_row row
          end
        end
      end
      p.serialize(output_file_path)
    end

    super
  end

  def validate_input_file
    md_file
  rescue
    @error = 'Bad input file.'
    nil
  end

  def md_file
    @md_file ||= File.new(input_file_path, 'r')
  end

  def md_content
    string = ''
    md_file.each_line do |line|
      string << line
    end
    string
  end

end
