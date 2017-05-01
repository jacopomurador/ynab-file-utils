require "thor"
require "roo"
require "roo-xls"
 
class YNABConverter < Thor
  desc "fab XLS_PATH [CSV_PATH]", "Convert xls account balance file of Fineco"
  def fab(xls_path, csv_path=nil)
    convert_fineco(xls_path, csv_path) { |row| [row[1],row[5],"","",row[3],row[2]] }
  end

  desc "fcc XLS_PATH [CSV_PATH]", "Convert xls credit card balance file of Fineco"
  def fcc(xls_path, csv_path=nil)
    convert_fineco(xls_path, csv_path) { |row| [row[1],row[2],"","",row[5],""] }
  end

  private

  def convert_fineco(xls_path, csv_path=nil)
    # open the excel
    xls = Roo::Spreadsheet.open(xls_path)
    
    # add header to cs
    csv_data = Array.new
    csv_data << ["Date","Payee","Category","Memo","Outflow","Inflow"]
    
    # read from xls and format csv for ynab
    i = 0
    discared = 0
    accept = 0
    xls.each do |row|
      i += 1
      if row[0]!= nil && /^(\d{2}\/\d{2}\/\d{4})$/.match(row[0].strip)
          # valid date, call the block to create the csv row
          csv_data << yield(row)
          accept += 1
      else
        # invalid date, skip the row
        discared += 1
      end
    end
    puts "Accept: #{accept} - Discared: #{discared}"

    # write csv
    csv_path = xls_path.sub(".xls",".csv") unless csv_path
    CSV.open(csv_path, "wb") do |csv|
      csv_data.each do |row|
        csv << row
      end
    end
  end
end
 
YNABConverter.start(ARGV)

