require 'csv'
require 'axlsx'

class FormattedList
  def initialize(list_name)
    @spreadsheet = Axlsx::Package.new
    @workbook = @spreadsheet.workbook
    @list = list_name
  end

  def create_styles
    @workbook.styles do |style|
      @orange_cell = style.add_style(bg_color: "FFB833")
    end
  end

  def create_sheet
    @workbook.add_worksheet(name: "Formatted") do |sheet|
      CSV.foreach("./#{@list}") do |row|
        sheet.add_row(row, style: check_for_empty_flight_info(row, row.length)) if !row[0].nil?
      end
    end
  end

  def check_for_empty_flight_info(row, length)
    colored_row = Array.new(length)

    if row[9].nil? && row[11].nil?
      colored_row[4] = @orange_cell
      colored_row[5] = @orange_cell
      colored_row[9] = @orange_cell
      colored_row[11] = @orange_cell
    end

    return colored_row
  end

  def serialize(file_name)
    @spreadsheet.serialize("#{file_name}.xlsx")
  end
end

list = FormattedList.new("example.csv")
list.create_styles
list.create_sheet
list.serialize("test")
