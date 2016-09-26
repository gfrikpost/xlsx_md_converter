require 'xlsx_md_converter/xlsx_to_md'
require 'xlsx_md_converter/md_to_xlsx'
require 'xlsx_md_converter/xlsx_to_simple_xlsx'

module XLSXMdConverter

  class << self

    def xlsx_to_md(file_path)
      XLSXToMD.new(file_path).convert
    end

    def md_to_xlsx(file_path)
      MDToXLSX.new(file_path).convert
    end

    def xlsx_to_simple_xlsx(file_path)
      XLSXToSimpleXLSX.new(file_path).convert
    end
  end
end
