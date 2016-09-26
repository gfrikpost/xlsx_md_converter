require 'roo'

class XLSXMdConverterBase

  def initialize(input_file_path)
    @input_file_path = input_file_path
  end

  def convert
    return puts(error) unless validate_input_file
    process
  end

  private

  def process
    puts 'file success converted'
  end

  attr_reader :input_file_path, :xlsx_file, :error

  def output_file_path
    "#{File.basename(input_file_path, '.*')}#{Time.now.to_i}.#{output_file_extension}"
  end

  def output_file_extension
    'xlsx'
  end

  def validate_input_file
    @xlsx_file = Roo::Spreadsheet.open(input_file_path, extension: :xlsx)
  rescue
    @error = 'Bad input file.'
    nil
  end
end
