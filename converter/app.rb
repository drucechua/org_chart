# app.rb  — Sinatra-only PDF → XLSX web app

# --- TLS trust: point Ruby/OpenSSL to a known CA bundle ---
# Option A: keep a user-level bundle (ensure it exists)
ENV["SSL_CERT_FILE"] ||= File.expand_path("~/.certs/cacert.pem")
# (Alternative) Option B: ship ./certs/cacert.pem with your repo and use:
# ENV["SSL_CERT_FILE"] ||= File.expand_path(File.join(__dir__, "certs", "cacert.pem"))
ENV.delete("SSL_CERT_DIR")                    # Force using the FILE path only

# --- Load dependencies ---
require "dotenv/load"     # Load env vars from .env (e.g., CONVERT_API_SECRET)
require "sinatra"         # Minimal web framework for routes (get/post)
require "convert_api"     # ConvertAPI client (PDF → XLSX conversion)
require "roo"             # Read spreadsheets (used by formatter)
require "caxlsx"          # Write/formats .xlsx files
require "fileutils"       # File utilities (mkdir, cp, rm, etc.)
require "securerandom"    # Generate safe random tokens for temp filenames
require "open-uri"        # Simple HTTP client; used for TLS sanity check

# --- Server bind/port (works locally and on Render/Fly/etc.) ---
set :bind, "0.0.0.0"                          # Listen on all interfaces
set :port, ENV.fetch("PORT", "4567")          # Use PORT env var or 4567

# --- Strong no-cache so browsers don’t serve stale pages/assets ---
before do
  cache_control :no_store, :no_cache, :must_revalidate, max_age: 0
  headers "Pragma" => "no-cache", "Expires" => "0"
end

get "/tls_selftest" do
  require "net/http"
  uri = URI("https://v2.convertapi.com/info")
  info = {}
  begin
    Net::HTTP.start(uri.host, 443, use_ssl: true) do |http|
      info[:peer_cert_subject] = http.peer_cert.subject.to_s
      info[:peer_cert_issuer]  = http.peer_cert.issuer.to_s
      res = http.request(Net::HTTP::Get.new(uri))
      info[:status] = res.code
      info[:body_head] = res.body&.slice(0, 120)
    end
  rescue => e
    info[:error] = "#{e.class}: #{e.message}"
  end
  info[:SSL_CERT_FILE] = ENV["SSL_CERT_FILE"]
  info[:SSL_CERT_DIR]  = ENV["SSL_CERT_DIR"]
  content_type :json
  info.to_json
end




# --- ConvertAPI configuration (reads CONVERT_API_SECRET from env/.env) ---
ConvertApi.configure { |c| c.api_credentials = ENV["CONVERT_API_SECRET"] }

# --- Routes ---

# GET /  → serve your upload page
get "/" do
  send_file File.join(settings.root, "user.html"), disposition: "inline"
end

# POST /convert  → receive a PDF and return a formatted XLSX
post "/convert" do
  # Accept either <input name="pdf"> or <input name="file">
  upload = params[:pdf] || params[:file]
  halt 400, "No file uploaded" unless upload && upload[:tempfile]

  # Basic PDF guard (extension or MIME acceptable)
  ext_ok  = File.extname(upload[:filename].to_s).downcase == ".pdf"
  mime_ok = upload[:type].to_s == "application/pdf"
  halt 415, "Only PDF files are accepted" unless ext_ok || mime_ok

  # Ensure local temp directory for intermediate files
  tmpdir = File.expand_path("tmp", __dir__)
  Dir.mkdir(tmpdir) unless Dir.exist?(tmpdir)

  # Build safe, unique filenames to avoid collisions and path issues
  token       = SecureRandom.hex(8)
  base_name   = File.basename(upload[:filename].to_s, ".*")
  pdf_path    = File.join(tmpdir, "#{base_name}-#{token}.pdf")
  raw_xlsx    = File.join(tmpdir, "#{base_name}-#{token}.xlsx")
  pretty_xlsx = File.join(tmpdir, "#{base_name}-#{token}.formatted.xlsx")

  # Stream the uploaded file to disk (avoid loading whole file into RAM)
  File.open(pdf_path, "wb") { |f| IO.copy_stream(upload[:tempfile], f) }

  # Quick TLS probe to fail fast with clear error if trust store is wrong
  begin
    URI.open("https://v2.convertapi.com/info") { |io| io.read(1) }
  rescue => e
    halt 502, "TLS check failed: #{e.class}: #{e.message}\nSSL_CERT_FILE=#{ENV['SSL_CERT_FILE']}"
  end

  begin
    # PDF → XLSX via ConvertAPI (OCR auto; single sheet)
    result = ConvertApi.convert(
      "xlsx",
      { File: pdf_path, OcrMode: "auto", SingleSheet: "true" },
      from_format: "pdf"
    )
    result.files.first.save(raw_xlsx)         # Save ConvertAPI output

    # Reformat for readability (headers, widths, freeze panes, wrapping)
    format_xlsx(raw_xlsx, pretty_xlsx)

    # Send the formatted workbook to the browser as a download
    send_file pretty_xlsx,
      filename: "#{base_name}.xlsx",
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

  rescue ConvertApi::Error => e
    # ConvertAPI-specific errors (bad key, API issue, etc.)
    halt 502, "Conversion failed: #{e.message}"
  rescue => e
    # Any other server-side error (I/O, formatting)
    halt 500, "Server error: #{e.class}: #{e.message}"
  ensure
    # Defer cleanup so we don't race send_file’s streaming
    paths = [pdf_path, raw_xlsx, pretty_xlsx]
    Thread.new { sleep 2; paths.each { |p| File.delete(p) if p && File.exist?(p) } }
  end
end

# --- XLSX formatting: cleans rows, builds styles, sets widths, freezes header ---
def format_xlsx(input_xlsx_path, output_xlsx_path)
  x     = Roo::Spreadsheet.open(input_xlsx_path)  # Open converted XLSX with Roo
  sheet = x.sheet(0)                               # Use first sheet

  # Read rows, normalize whitespace in strings, and drop fully-blank rows
  rows = sheet.each_row_streaming(pad_cells: true).map do |r|
    r.map do |cell|
      v = cell&.value
      v.is_a?(String) ? v.strip.gsub(/\s+/, " ") : v
    end
  end
  rows.reject! { |r| r.nil? || r.all? { |c| c.nil? || c.to_s.strip.empty? } }

  # If no usable data remains, just copy original file and exit
  if rows.empty?
    FileUtils.cp(input_xlsx_path, output_xlsx_path)
    return
  end

  # Use first row as header; fill blank headers as "Column N"
  header = rows.shift
  header = header.map.with_index { |h, i| (h.to_s.strip.empty? ? "Column #{i + 1}" : h.to_s.strip) }
  data_rows = rows
  col_count = header.length

  # Compute simple column width heuristic from max cell lengths
  max_chars = Array.new(col_count, 0)
  ([header] + data_rows).each do |r|
    (r || []).each_with_index do |cell, i|
      len = cell.nil? ? 0 : cell.to_s.length
      max_chars[i] = [max_chars[i], len].max
    end
  end
  widths = max_chars.map { |n| [[n * 1.1 + 2, 10].max, 60].min }  # min 10, cap 60

  # Build a new XLSX with basic styling
  pkg = Axlsx::Package.new
  wb  = pkg.workbook

  styles = wb.styles
  header_style = styles.add_style(
    b: true,
    alignment: { horizontal: :center, vertical: :center, wrap_text: true },
    bg_color: "EEEEEE", fg_color: "000000"
  )
  cell_style = styles.add_style(
    alignment: { horizontal: :left, vertical: :top, wrap_text: true }
  )

  wb.add_worksheet(name: "Sheet1") do |ws|
    ws.column_widths(*widths)                   # Set column widths in one call
    ws.sheet_view.pane do |p|                   # Freeze header row
      p.top_left_cell = "A2"
      p.state = :frozen
      p.y_split = 1
    end

    ws.add_row(header, style: Array.new(header.length, header_style), height: 22)
    data_rows.each do |r|
      padded = (r || []) + Array.new([0, header.length - (r&.length || 0)].max, nil)
      ws.add_row(padded, style: Array.new(header.length, cell_style))
    end
  end

  pkg.serialize(output_xlsx_path)               # Write the final formatted XLSX
end



# # app.rb

# # --- LOAD REQUIRED LIBRARIES ---

# require 'dotenv/load' 
# require 'convert_api'
# require 'functions_framework'
# require 'sinatra'           # Used implicitly by Functions Framework for parsing requests
# require 'roo'
# require 'caxlsx'
# require 'fileutils'
# require 'openssl'
# require 'net/http'
# require 'json'


# # --- SERVER CONFIGURATION ---

# # Bind the Sinatra App: Listen locally and on hosts (Render/Fly/etc).
# set :bind, '0.0.0.0'
# set :port, ENV.fetch('PORT', '4567')


# # --- BROWSER CACHING CONTROL ---

# # Disable browser caching so updated pages / scripts are always reloaded
# before do
#   cache_control :no_store, :no_cache, :must_revalidate, max_age: 0
#   headers 'Pragma' => 'no-cache', 'Expires' => '0'
# end


# # --- TLS / SSL TRUST CONFIGURATION ---

# # Specify a known Certificate Authority bundle for HTTPS verification 
# ENV['SSL_CERT_FILE'] ||= File.expand_path("~/.certs/cacert.pem")                  # Option 1: Custom path 
# # ENV['SSL_CERT_FILE'] ||= `brew --prefix`.strip + "/etc/openssl@3/cert.pem"      # Option 2: Using homebrew shelling
# ENV.delete('SSL_CERT_DIR')                                                        # Remove SSL_CERT_DIR so Ruby uses SSL_CERT_FILE explicitly


# # --- CONVERTAPI CONFIGURATION ---
# ConvertApi.configure { |c| c.api_credentials = ENV["CONVERT_API_SECRET"] }


# # --- ROUTES ---

# # Root route: serve the upload HTML page
# get "/" do
#   send_file File.join(settings.root, "user.html")
# end

# # POST /convert → Handles PDF uploads and converts them to XLSX
# post "/convert" do

#   # Accept files uploaded under the form field name "pdf" or "file"
#   upload = params[:pdf] || params[:file]
#   halt 400, "No file uploaded" unless upload && upload[:tempfile]

#   # Basic guard to verify that the uploaded file is a PDF
#   ext_ok  = File.extname(upload[:filename].to_s).downcase == ".pdf"
#   mime_ok = upload[:type].to_s == "application/pdf"
#   halt 415, "Only PDF files are accepted" unless ext_ok || mime_ok

#   # Create tmp directory if it doesn’t exist (for intermediate files)
#   tmpdir = File.expand_path("tmp", __dir__)
#   Dir.mkdir(tmpdir) unless Dir.exist?(tmpdir)

#   # Generate random unique filenames to avoid conflicts or unsafe names
#   token       = SecureRandom.hex(8)
#   base_name   = File.basename(upload[:filename].to_s, ".*")
#   pdf_path    = File.join(tmpdir, "#{base_name}-#{token}.pdf")
#   raw_xlsx    = File.join(tmpdir, "#{base_name}-#{token}.xlsx")
#   pretty_xlsx = File.join(tmpdir, "#{base_name}-#{token}.formatted.xlsx")

#   # Stream uploaded file to disk to avoid large memory reads
#   File.open(pdf_path, "wb") { |f| IO.copy_stream(upload[:tempfile], f) }

#   # --- TLS SELF TEST ---

#   # Test an HTTPS request to ensure the CA bundle is valid and SSL works
#   begin
#     URI.open("https://v2.convertapi.com/info") { |io| io.read(1) }
#   rescue => e
#     halt 502, "TLS check failed: #{e.class}: #{e.message}\nSSL_CERT_FILE=#{ENV['SSL_CERT_FILE']}"
#   end

#   # --- MAIN CONVERSION LOGIC ---
#   begin
#     # Convert PDF → XLSX using ConvertAPI
#     result = ConvertApi.convert(
#       "xlsx",
#       { File: pdf_path, OcrMode: "auto", SingleSheet: "true" },
#       from_format: "pdf"
#     )

#     # Save the raw converted file locally
#     result.files.first.save(raw_xlsx)

#     # Reformat the resulting XLSX (adjust columns, headers, etc.)
#     format_xlsx(raw_xlsx, pretty_xlsx)

#     # Send the formatted XLSX back to the user as a file download
#     send_file pretty_xlsx,
#       filename: "#{base_name}.xlsx",
#       type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

#   rescue ConvertApi::Error => e
#   # Catch and return conversion errors from ConvertAPI
#     halt 502, "Conversion failed: #{e.message}"
#   rescue => e
#     # Catch unexpected runtime errors (e.g., file I/O)
#     halt 500, "Server error: #{e.class}: #{e.message}"
#   ensure
#     # Cleanup temporary files after a short delay (so send_file finishes)
#     paths = [pdf_path, raw_xlsx, pretty_xlsx]
#     Thread.new { sleep 2; paths.each { |p| File.delete(p) if p && File.exist?(p) } }
#   end
# end


# # --- XLSX FORMATTING FUNCTION ---

# def format_xlsx(input_xlsx_path, output_xlsx_path)
#   # Open the converted XLSX using Roo
#   x = Roo::Spreadsheet.open(input_xlsx_path)
#   sheet = x.sheet(0)

#   # Read rows, normalize whitespace, and remove empty rows
#   rows = sheet.each_row_streaming(pad_cells: true).map do |r|
#     r.map { |cell|
#       v = cell&.value
#       v.is_a?(String) ? v.strip.gsub(/\s+/, ' ') : v
#     }
#   end
#   rows.reject! { |r| r.nil? || r.all? { |c| c.nil? || c.to_s.strip.empty? } }

#   # If no usable rows exist, just copy the file as-is
#   if rows.empty?
#     FileUtils.cp(input_xlsx_path, output_xlsx_path)
#     return
#   end

#   # Use first row as header and fill blanks with “Column X”
#   header = rows.shift
#   header = header.map.with_index do |h, i|
#     s = h.to_s.strip
#     s.empty? ? "Column #{i + 1}" : s
#   end
#   data_rows = rows

#   # Calculate approximate column widths (based on max text length)
#   col_count = header.length
#   max_chars = Array.new(col_count, 0)
#   ([header] + data_rows).each do |r|
#     r = (r || [])
#     col_count.times do |i|
#       cell = r[i]
#       len = cell.nil? ? 0 : cell.to_s.length
#       max_chars[i] = [max_chars[i], len].max
#     end
#   end

#   # Convert character lengths to Excel column widths (min 10, cap 60)
#   widths = max_chars.map { |n| [[n * 1.1 + 2, 10].max, 60].min }

#   # --- BUILD FORMATTED EXCEL FILE ---

#   pkg = Axlsx::Package.new
#   wb  = pkg.workbook

#   # Define styles (header bold + gray background, normal cell left-aligned)
#   styles = wb.styles
#   header_style = styles.add_style(
#     b: true,
#     alignment: { horizontal: :center, vertical: :center, wrap_text: true },
#     bg_color: 'EEEEEE',
#     fg_color: '000000'
#   )
#   cell_style = styles.add_style(
#     alignment: { horizontal: :left, vertical: :top, wrap_text: true }
#   )

#   # Add data and formatting to worksheet
#   wb.add_worksheet(name: 'Sheet1') do |ws|
#     # Set calculated column widths
#     widths.each_with_index { |w, i| ws.column_info[i].width = w }

#     # Freeze top header row
#     ws.sheet_view.pane do |p|
#       p.top_left_cell = 'A2'
#       p.state = :frozen
#       p.y_split = 1
#     end

#     # Write header and data rows with styles
#     ws.add_row(header, style: Array.new(header.length, header_style), height: 22)
#     data_rows.each do |r|
#       padded = r + Array.new([0, header.length - r.length].max, nil)
#       ws.add_row(padded, style: Array.new(header.length, cell_style))
#     end
#   end

#   # Save the formatted workbook to output path
#   pkg.serialize(output_xlsx_path)
# end

# # --- DUPLICATE CONFIG SECTION (redundant, can remove safely) ---
# ConvertApi.configure do |c|
#   c.api_credentials = ENV['CONVERT_API_SECRET']
# end

# # --- GOOGLE CLOUD FUNCTION HANDLER (optional, not used in Sinatra) ---
# # This section defines a serverless handler for Google Cloud Functions
# # It duplicates the above Sinatra logic, but runs in a stateless function context
# # You can delete this if you only plan to use Sinatra
# FunctionsFramework.http("convert_pdf_handler") do |request|
#   # Extract uploaded file
#   file_param = request.params['file'] 
#   unless file_param && file_param[:tempfile]
#     return [400, {"Content-Type" => "text/plain"}, ["Error: No file uploaded."]]
#   end

#   temp_pdf_file = file_param[:tempfile]
#   original_pdf_name = file_param[:filename]

#   base_name = File.basename(original_pdf_name, File.extname(original_pdf_name))
#   output_xlsx_path = "/tmp/#{base_name}.xlsx"

#   begin
#     # Convert PDF → XLSX
#     result = ConvertApi.convert(
#       'xlsx', 
#       { File: temp_pdf_file.path, OcrMode: 'auto', SingleSheet: 'true' }, 
#       from_format: 'pdf'
#     )
#     result.files.first.save(output_xlsx_path)

#     # Reformat output
#     formatted_xlsx_path = "/tmp/#{base_name}.formatted.xlsx"
#     format_xlsx(output_xlsx_path, formatted_xlsx_path)

#     # Prepare response
#     converted_file_data = File.read(formatted_xlsx_path)
#     headers = {
#       "Content-Type" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#       "Content-Disposition" => "attachment; filename=\"#{base_name}.xlsx\""
#     }
#     [200, headers, [converted_file_data]]

#   rescue ConvertApi::Error => e
#     [500, {"Content-Type" => "text/plain"}, ["Conversion Failed: #{e.message}"]]
#   ensure
#     # Clean up temp files
#     File.delete(output_xlsx_path) if File.exist?(output_xlsx_path)
#     File.delete(formatted_xlsx_path) if defined?(formatted_xlsx_path) && File.exist?(formatted_xlsx_path)
#   end
# end