require 'csv'
require 'axlsx'
require 'date'

class FormattedList
  def initialize(list_name, list_type)
    @spreadsheet = Axlsx::Package.new
    @workbook = @spreadsheet.workbook
    @row_styler = RowStyler.new(@workbook)
    @list = list_name
    @list_type = list_type
  end

  def create_prettified_sheet
    if @list_type == "housing"
      prettify_for_housing
    elsif @list_type == "homestay"
      prettify_for_homestay
    else
      return false
    end
  end

  def prettify_for_housing
    @workbook.add_worksheet(name: "Formatted") do |sheet|
      CSV.foreach("./#{@list}") do |row|
        if $INPUT_LINE_NUMBER <= 5
          sheet.add_row(row, style: @row_styler.bold_headers)
        else
          sheet.add_row(row, style: @row_styler.style_housing_row(row, row.length)) if correctly_formatted_row(row)
        end
      end
    end
  end

  def prettify_for_homestay
    @workbook.add_worksheet(name: "Formatted") do |sheet|
      CSV.foreach("./#{@list}") do |row|
        if $INPUT_LINE_NUMBER == 1
          sheet.add_row(row, style: @row_styler.bold_headers)
        else
          sheet.add_row(row, style: @row_styler.style_homestay_row(row, row.length)) if correctly_formatted_row(row)
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
      @yellow_cell = style.add_style(bg_color: "FFFF00")
      @red_cell = style.add_style(bg_color: "FF0000")
      @light_red_cell = style.add_style(bg_color: "D99694")
    end
  end

  def style_housing_row(row, length)
    styled_row_array = Array.new(length)

    if row[9].nil? && row[11].nil?
      styled_row_array[4] = @yellow_cell
      styled_row_array[5] = @yellow_cell
      styled_row_array[9] = @yellow_cell
      styled_row_array[11] = @yellow_cell
    end

    return styled_row_array
  end

  def style_homestay_row(row, length)
    styled_row_array = Array.new(length)
    date = Date.parse(row[0])

    if date.wday.between?(1, 5)
      styled_row_array[0] = @red_cell
    end

    if row[11] != "Private Accommodation" && row[17].nil? && row[18].nil?
      styled_row_array[5] = @yellow_cell
      styled_row_array[6] = @yellow_cell
      styled_row_array[17] = @yellow_cell
      styled_row_array[18] = @yellow_cell
    end

    if row[11] != "Private Accommodation" && row[14].nil? && row[15].nil?
      styled_row_array[14] = @light_red_cell
      styled_row_array[15] = @light_red_cell
    end

    return styled_row_array
  end

  def bold_headers
    return @bold_cell
  end
end

list = FormattedList.new("homestay.csv", "homestay")
list.create_prettified_sheet
#list.debug
list.serialize("homestay")
