class XLSXToSimpleXLSX < XLSXMdConverterBase

  private

  def process
    Axlsx::Package.new do |p|
      xlsx_file.each_with_pagename do |name, input_file_sheet|
        p.workbook.add_worksheet(name: name) do |out_file_sheet|
          input_file_sheet.each do |row|
            out_file_sheet.add_row row
          end
          p.serialize(output_file_path)
        end
      end
    end
    super
  end

end
