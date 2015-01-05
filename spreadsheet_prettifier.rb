require 'csv'
require 'axlsx'

class FormattedList
  def initialize(list_name)
    @spreadsheet = Axlsx::Package.new
    @workbook = @spreadsheet.workbook
    @row_styler = RowStyler.new(@workbook)
    @list = list_name
  end

  def create_prettified_sheet
    @workbook.add_worksheet(name: "Formatted") do |sheet|
      CSV.foreach("./#{@list}") do |row|
        if $INPUT_LINE_NUMBER <= 5
          sheet.add_row(row, style: @row_styler.bold_headers)
        else
          sheet.add_row(row, style: @row_styler.empty_flight_info(row, row.length)) if correctly_formatted_row(row)
        end
      end
    end
  end

  def correctly_formatted_row(row)
    [row[0], row[3]].all? # returns false if any value is false or nil. default block {|item| item}
  end

  def serialize(file_name)
    @spreadsheet.serialize("#{file_name}.xlsx")
  end

  def debug
    puts @workbook.worksheets.name
  end
end

class RowStyler
  def initialize(workbook)
    @workbook = workbook
    create_styles
  end

  def create_styles
    @workbook.styles do |style|
      @bold_cell = style.add_style(b: true)
      @orange_cell = style.add_style(bg_color: "FFB833")
    end
  end

  def empty_flight_info(row, length)
    styled_row_array = Array.new(length)

    if row[9].nil? && row[11].nil?
      styled_row_array[4] = @orange_cell
      styled_row_array[5] = @orange_cell
      styled_row_array[9] = @orange_cell
      styled_row_array[11] = @orange_cell
    end

    return styled_row_array
  end

  def bold_headers
    return @bold_cell
  end
end

list = FormattedList.new("housing.csv")
list.create_prettified_sheet
#list.debug
list.serialize("housing")
