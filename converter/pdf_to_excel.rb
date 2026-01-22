require 'convert_api'

# ⚠️ IMPORTANT: Replace 'YOUR_CONVERTAPI_SECRET' with your actual key.
ConvertApi.configure do |config|
  config.api_secret = 'r1y8W8kzBgJSi72ItbHDZ9ks73AwKYE2' 
end

# Define file paths using your actual filenames
PDF_INPUT = 'local Employee List 03-10-2025.pdf'
XLSX_OUTPUT = 'converted_employee_list.xlsx'