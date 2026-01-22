require 'convert_api'

# --- 1. CONFIGURATION AND AUTHENTICATION ---

# ⚠️ IMPORTANT: Replace 'YOUR_CONVERTAPI_SECRET' with your actual Production Token
ConvertApi.configure do |config|
  config.api_credentials = 'r1y8W8kzBgJSi72ItbHDZ9ks73AwKYE2' 
end

# --- 2. CONVERSION LOGIC (Non-hardcoded) ---

def convert_pdf_to_excel
  # Check for the input file argument
  if ARGV.empty?
    puts "❌ Error: Please provide the path to the PDF file as a command-line argument."
    puts "Usage: ruby convert.rb /path/to/your/input.pdf"
    return
  end
  
  # Get the PDF file path from the first command-line argument (ARGV[0])
  pdf_input_path = ARGV[0]
  
  # Check if the file exists
  unless File.exist?(pdf_input_path)
    puts "❌ Error: Input file not found at '#{pdf_input_path}'."
    return
  end
  
  # Dynamically determine the output file name
  # Example: 'employee.pdf' -> 'employee.xlsx'
  base_name = File.basename(pdf_input_path, File.extname(pdf_input_path))
  xlsx_output_path = "#{base_name}.xlsx"
  
  puts "Starting conversion of '#{pdf_input_path}'..."
  
  begin
    # The 'pdf' to 'xlsx' conversion request
    result = ConvertApi.convert(
      'xlsx', 
      { 
        File: pdf_input_path,
        OcrMode: 'auto', 
        SingleSheet: 'true' 
      }, 
      from_format: 'pdf'
    )
    
    # Save the resulting file
    result.files.first.save(xlsx_output_path)

    puts "✅ Conversion successful! File saved as '#{xlsx_output_path}'"
  
  rescue ConvertApi::Error => e
    puts "❌ Conversion failed: #{e.message}"
  end
end

# --- 3. EXECUTION ---
convert_pdf_to_excel