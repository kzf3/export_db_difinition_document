require "bundler"
require "date"
require "./const.rb"
require "./config.rb"
Bundler.require

client = Mysql2::Client.new(
  :host => $database[:host],
  :username => $database[:username],
  :password => $database[:password],
  :database => $database[:database]
)

show_tables = %{show tables}

all_tables = client.query(show_tables)
show_full_columns = %{show full columns from ?}
#stmt = client.prepare(show_full_columns)

package = Axlsx::Package.new
wb = package.workbook

wb.styles do |s|
  grey_cell = s.add_style :bg_color => "e6e6e6", :border => {:style => :thin, :color => "00"}, :alignment => {:horizontal => :center}
  normal_cell = s.add_style :bg_color => false, :border => {:style => :thin, :color => "00"}
  cell = s.add_style :bg_color => false
  header_style = [grey_cell, grey_cell, grey_cell, grey_cell, grey_cell]
  wb.add_worksheet(:name => 'テーブル一覧') do |sheet|
    sheet.add_row ["テーブル一覧"]
    sheet.add_row [""]
    sheet.add_row ["NO", "論理名", "物理名", "テーブル概要", "備考"], :style => header_style

    count = 0
    all_tables.each do |res|
      table = res["Tables_in_portweb"]
      sheet.add_row [count+=1, "", table, "", ""], :style => [normal_cell, normal_cell, normal_cell, normal_cell, normal_cell]
    end
  end
end

cell_position = 4

all_tables.each do |res|

  table = res["Tables_in_portweb"]
  show_full_columns = %{show full columns from #{table}}
  results = client.query(show_full_columns)
  
  wb.styles do |s|
    grey_cell = s.add_style :bg_color => "e6e6e6", :border => {:style => :thin, :color => "00"}, :alignment => {:horizontal => :center}
    normal_cell = s.add_style :bg_color => false, :border => {:style => :thin, :color => "00"}
    cell = s.add_style :bg_color => false

  
    wb.add_worksheet(:name => table.slice(0, 30)) do |sheet|
      define_row = [cell, grey_cell, normal_cell, normal_cell, grey_cell, grey_cell, normal_cell, normal_cell]
      sheet.styles.fonts.first.name = "メイリオ"
      sheet.merge_cells("A1:B1")
      sheet.add_row ["テーブル情報"]
      sheet.merge_cells("C2:D2")
      sheet.merge_cells("E2:F2")
      sheet.merge_cells("G2:H2")
      sheet.add_row ["", SYSTEM_NAME, "portweb", "", CREATE_USER, "",  "ファンチーム", ""], :style => define_row
      sheet.merge_cells("C3:D3")
      sheet.merge_cells("E3:F3")
      sheet.merge_cells("G3:H3")
      sheet.add_row ["", SCHEMA_NAME, "portweb", "", CREATE_DATE, "", Date.today.strftime("%Y/%m/%d"), ""], :style => define_row
      sheet.merge_cells("C4:D4")
      sheet.merge_cells("E4:F4")
      sheet.merge_cells("G4:H4")
      sheet.add_row ["", "論理名", "=テーブル一覧!B" + cell_position.to_s, "", MODIFIED_DATE, "", Date.today.strftime("%Y/%m/%d"), ""], :style => define_row
      sheet.merge_cells("C5:D5")
      sheet.merge_cells("E5:F5")
      sheet.merge_cells("G5:H5")
      sheet.add_row ["", "物理名", "=テーブル一覧!C" + cell_position.to_s, "", "RDBMS", "", "MySQL", ""], :style => define_row 
      sheet.merge_cells("B6:H6")
      sheet.add_row ["", "備考", "", "", "", "", "", ""], :style => [cell, grey_cell, grey_cell, grey_cell, grey_cell, grey_cell, grey_cell, grey_cell]
      sheet.merge_cells("B7:H9")
      note_row = [cell, normal_cell, normal_cell, normal_cell, normal_cell, normal_cell, normal_cell, normal_cell]
      sheet.add_row ["", "", "", "", "", "", "", ""], :style => note_row
      sheet.add_row ["", "", "", "", "", "", "", ""], :style => note_row
      sheet.add_row ["", "", "", "", "", "", "", ""], :style => note_row
      sheet.add_row [""]
      sheet.merge_cells("A11:B11")
      sheet.add_row ["カラム情報"]
      sheet.add_row ["NO", "Field", "Type", "UN", "Null", "Key", "Default", "Extra", "Comment"], :style => [grey_cell, grey_cell, grey_cell, grey_cell, grey_cell, grey_cell, grey_cell, grey_cell, grey_cell]
  
      td_row_style = [normal_cell, normal_cell, normal_cell, normal_cell, normal_cell, normal_cell, normal_cell, normal_cell, normal_cell]
      count = 0
      record = []
      not_export = ["Collation", "Privileges"]
      cell_position += 1
      results.each do |row|
        record = []
        record.push(count+=1)
        row.each do |key, value|
          if !not_export.include?(key)
            if key == "Type"
              record.push(value.split(" ")[0])
              record.push(value.split(" ")[1])
            else
              record.push(value)
            end
          end
        end
        sheet.add_row record, :style => td_row_style
      end
    end
  end
end

package.serialize('/vagrant/' + $database[:database] + 'DB定義書' + Date.today.strftime("%Y%m%d") + '.xlsx')
