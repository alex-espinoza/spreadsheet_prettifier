require 'csv'
require 'axlsx'
require 'date'
require 'pry'

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
    elsif @list_type == "awaiting"
      prettify_for_awaiting
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

  def prettify_for_awaiting
    @workbook.add_worksheet(name: "Formatted") do |sheet|
      CSV.foreach("./#{@list}", encoding: 'windows-1251:utf-8') do |row|
        if $INPUT_LINE_NUMBER == 1
          @headers_array = row.dup
          sheet.add_row(row, style: @row_styler.bold_headers)
        else
          sheet.add_row(row, style: @row_styler.style_awaiting_row(row, row.length)) if correctly_formatted_awaiting_row(row)
        end
      end

      # Todo: 6. Separate the Residence bookings from the Homestay bookings by a few rows (not sure what this means)

      # find elements needed to get deleted, EDGE CASE: 8. is it "Finance" or "Finance Note?"
      columns_to_delete_array = @headers_array.each_index.select {|i| @headers_array[i] =~ /^Status|^Accommodation|^Finance Note|^Room|^Type/}
      @headers_array.reject!.each_with_index {|v, i| columns_to_delete_array.include?(i)}

      sheet.rows.each do |row|
        row.cells.reject!.each_with_index {|v, i| columns_to_delete_array.include?(i)}
      end

      finance_special_request_column = @headers_array.find_index("Finance Special Request Note")
      sheet.column_widths(*Array.new(@headers_array.length).insert(finance_special_request_column, 45)) # * is the splat operator, turns array of values into method arguments
    end
  end

  def correctly_formatted_row(row)
    [row[0], row[3]].all? # returns false if any value is false or nil. default block {|item| item}
  end

  def correctly_formatted_awaiting_row(row)
    if row[11] =~ /^Private Accommodation/ || row[9] =~ /^Booked|^Changed/
      row[11] =~ /^Residence|^Self-Catering/ && (row[10].nil? || row[14].nil?) ? true : false
    else
      return true
    end
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
      @finance_special_request_cell = style.add_style({alignment: {wrap_text: true}})
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

  def style_awaiting_row(row, length)
    styled_row_array = Array.new(length)

    if row[11] =~ /^Residence|^Self-Catering/ && (row[10].nil? || row[14].nil?)
      styled_row_array[4] = @red_cell
      styled_row_array[5] = @red_cell
      styled_row_array[10] = @red_cell
      styled_row_array[14] = @red_cell
    end

    if row[11] == "Half-Board Twin Room"
      styled_row_array[11] = @light_red_cell
    end

    # EDGE CASE: Will "Finance Note" ever contain information you need? Because this moves/appends everything to "Finance Special Request" and deletes "Finance Note" data
    if row[12]
      row[13].nil? ? row[13] = row[12] : row[13].concat(" #{row[12]}")
      row[12] = nil
    end

    styled_row_array[13] = @finance_special_request_cell

    return styled_row_array
  end

  def bold_headers
    return @bold_cell
  end
end

list = FormattedList.new("awaiting.csv", "awaiting")
list.create_prettified_sheet
#list.debug
list.serialize("awaiting")
