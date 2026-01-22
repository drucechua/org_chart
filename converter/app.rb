# app.rb
# You must first run 'bundle install' to ensure all gems are available

require 'dotenv/load' 
require 'convert_api'
require 'functions_framework'
require 'sinatra'           # Used implicitly by Functions Framework for parsing requests
require 'roo'
require 'caxlsx'

get "/" do
  send_file File.join(settings.root, "user.html")
end

post '/convert' do
  # Accept either <input name="pdf"> or <input name="file">
  upload = params[:pdf] || params[:file]
  halt 400, "No file uploaded" unless upload && upload[:tempfile]

  # Ensure temp folder exists
  Dir.mkdir('tmp') unless Dir.exist?('tmp')

  # Save uploaded PDF
  pdf_path   = File.join('tmp', upload[:filename])
  File.open(pdf_path, 'wb') { |f| f.write(upload[:tempfile].read) }

  base       = File.basename(upload[:filename], File.extname(upload[:filename]))
  raw_xlsx   = File.join('tmp', "#{base}.xlsx")
  pretty_xlsx= File.join('tmp', "#{base}.formatted.xlsx")

  # --- PDF -> XLSX via ConvertAPI ---
  begin
    result = ConvertApi.convert(
      'xlsx',
      { File: pdf_path, OcrMode: 'auto', SingleSheet: 'true' },
      from_format: 'pdf'
    )
    result.files.first.save(raw_xlsx)

    # --- Format the XLSX you just created ---
    format_xlsx(raw_xlsx, pretty_xlsx)

    # Return the formatted XLSX to the browser
    send_file pretty_xlsx,
      filename: "#{base}.xlsx",
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

  rescue ConvertApi::Error => e
    halt 500, "Conversion Failed: #{e.message}"
  ensure
    # Cleanup
    [pdf_path, raw_xlsx, pretty_xlsx].each { |p| File.delete(p) if p && File.exist?(p) }
  end
end




def format_xlsx(input_xlsx_path, output_xlsx_path)
  # Open the input workbook
  x = Roo::Spreadsheet.open(input_xlsx_path)
  sheet = x.sheet(0)

  # Pull all rows (arrays) and drop completely blank rows
  rows = sheet.each_row_streaming(pad_cells: true).map do |r|
    r.map { |cell|
      v = cell&.value
      v.is_a?(String) ? v.strip.gsub(/\s+/, ' ') : v
    }
  end

  rows.reject! { |r| r.nil? || r.all? { |c| c.nil? || c.to_s.strip.empty? } }
  if rows.empty?
    # Nothing to write—just copy as-is
    FileUtils.cp(input_xlsx_path, output_xlsx_path)
    return
  end

  # Use the first non-empty row as header; fill missing header names
  header = rows.shift
  header = header.map.with_index do |h, i|
    s = h.to_s.strip
    s.empty? ? "Column #{i + 1}" : s
  end

  data_rows = rows

  # Compute approximate column widths (characters) from header + data
  col_count = header.length
  max_chars = Array.new(col_count, 0)

  ( [header] + data_rows ).each do |r|
    r = (r || [])
    col_count.times do |i|
      cell = r[i]
      len = cell.nil? ? 0 : cell.to_s.length
      max_chars[i] = [max_chars[i], len].max
    end
  end

  # Convert char-count to Excel width (rough heuristic)
  widths = max_chars.map { |n| [[n * 1.1 + 2, 10].max, 60].min } # min 10, cap 60

  # Build a new, formatted workbook
  pkg = Axlsx::Package.new
  wb  = pkg.workbook

  styles = wb.styles
  header_style = styles.add_style(
    b: true,
    alignment: { horizontal: :center, vertical: :center, wrap_text: true },
    bg_color: 'EEEEEE',
    fg_color: '000000'
  )
  cell_style = styles.add_style(
    alignment: { horizontal: :left, vertical: :top, wrap_text: true }
  )

  wb.add_worksheet(name: 'Sheet1') do |ws|
    # Column widths
    widths.each_with_index { |w, i| ws.column_info[i].width = w }

    # Freeze header row
    ws.sheet_view.pane do |p|
      p.top_left_cell = 'A2'
      p.state = :frozen
      p.y_split = 1
    end

    # Header + data
    ws.add_row(header, style: Array.new(header.length, header_style), height: 22)
    data_rows.each do |r|
      # Right-pad row to header length to avoid nil style mismatch
      padded = r + Array.new([0, header.length - r.length].max, nil)
      ws.add_row(padded, style: Array.new(header.length, cell_style))
    end
  end

  pkg.serialize(output_xlsx_path)
end


# 1. AUTHENTICATION (Must use Environment Variable for production)
# Configure API key from an environment variable called CONVERT_API_SECRET
ConvertApi.configure do |config|
  config.api_credentials = ENV['CONVERT_API_SECRET']
end

# 2. DEFINE THE HTTP FUNCTION
FunctionsFramework.http("convert_pdf_handler") do |request|
  # Ensure the request is a file upload
  file_param = request.params['file'] 
  
  unless file_param && file_param[:tempfile]
    return [400, {"Content-Type" => "text/plain"}, ["Error: No file uploaded."]]
  end

  # Get the temporary uploaded file and original filename
  temp_pdf_file = file_param[:tempfile]
  original_pdf_name = file_param[:filename]
  
  # 3. CONVERSION LOGIC (YOUR EXISTING LOGIC)
  
  base_name = File.basename(original_pdf_name, File.extname(original_pdf_name))
  output_xlsx_path = "/tmp/#{base_name}.xlsx" # Save output to temp directory
  
  begin
    result = ConvertApi.convert(
      'xlsx', 
      { 
        File: temp_pdf_file.path, # Pass the path of the temporary file
        OcrMode: 'auto', 
        SingleSheet: 'true' 
      }, 
      from_format: 'pdf'
    )
    
    # Save the resulting file to the temp directory
    result.files.first.save(output_xlsx_path)

    # 4. RETURN THE FILE TO THE USER

    # Save ConvertAPI result
    result.files.first.save(output_xlsx_path)

    # NEW: create a formatted copy
    formatted_xlsx_path = "/tmp/#{base_name}.formatted.xlsx"
    format_xlsx(output_xlsx_path, formatted_xlsx_path)

    # Read formatted file
    converted_file_data = File.read(formatted_xlsx_path)

    headers = {
    "Content-Type" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "Content-Disposition" => "attachment; filename=\"#{base_name}.xlsx\""
    }

    [200, headers, [converted_file_data]]

    
    # # Read the converted file data
    # converted_file_data = File.read(output_xlsx_path)
    
    # # Set headers for file download
    # headers = {
    #   "Content-Type" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    #   "Content-Disposition" => "attachment; filename=\"#{base_name}.xlsx\""
    # }
    
    # # Return a Rack response: [status, headers, body]
    # [200, headers, [converted_file_data]]

  rescue ConvertApi::Error => e
    return [500, {"Content-Type" => "text/plain"}, ["Conversion Failed: #{e.message}"]]
  ensure
    # Clean up temporary files
    File.delete(output_xlsx_path) if File.exist?(output_xlsx_path)
    File.delete(formatted_xlsx_path) if defined?(formatted_xlsx_path) && File.exist?(formatted_xlsx_path)
  end
end