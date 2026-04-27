require 'roo'
require 'net/http'
require 'uri'
require 'openssl'
require 'csv'

def prepare_file
  begin
    xlsx = Roo::Spreadsheet.open('./filename.xlsx')
    xlsx.default_sheet = 'DATA'

    header = xlsx.row(1)
    csv_data = [header]

    (2..xlsx.last_row).each do |row_index|
      row = []
      header.map.with_index(1) do |value, col_index|
        row << xlsx.formatted_value(row_index, col_index)
      end
      csv_data << row
    end

    CSV.open('./filename.csv', 'w') do |csv|
      csv_data.each { |row| csv << row }
    end

  rescue => e
    puts "Error converting Excel file: #{e.message}"
  end
end

def send_file
  begin
    prepare_file

    file_path = './filepath'
    folder_id = "XXXXXXX"
    uri = URI("")
    uri.query = URI.encode_www_form({
      "sheetName" => "",
      "headerRowIndex" => "0",
      "primaryColumnIndex" => "0"
    })

    http = Net::HTTP.new(uri.host, uri.port)
    http.use_ssl = true
    http.verify_mode = OpenSSL::SSL::VERIFY_PEER
    store = OpenSSL::X509::Store.new
    store.set_default_paths
    store.flags = OpenSSL::X509::V_FLAG_PARTIAL_CHAIN
    http.cert_store = store

    http.open_timeout = 120
    http.read_timeout = 300
    http.write_timeout = 300

    request = Net::HTTP::Post.new(uri)
    request["Authorization"] = "Bearer XXXXXXXXXXXXXXXXXXXXXXX"
    request["smartsheet-integration-source"] = "AI,TNC,My-AI-Connector-v2"
    request["Content-Disposition"] = 'attachment; filename="somefile.csv"'
    request["Content-Type"] = "text/csv"
    request["Content-Length"] = File.size(file_path).to_s
    request.body_stream = File.open(file_path, 'rb')

    response = http.request(request)
    request.body_stream.close

    puts "Upload completed with status: #{response.code}"
  rescue => e
    puts "Error during upload: #{e.message}"
    request.body_stream&.close if defined?(request) && request&.body_stream
  end
end

send_file
