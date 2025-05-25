# Document Text Extractor

A powerful tool that extracts text from images and PDFs using Amazon Bedrock's Claude AI models, with specialized formatting for financial documents.

## Features

- **Text Extraction**: Extract text from images (JPG, JPEG, PNG) and PDF documents
- **Multi-format Output**: Save extracted text in multiple formats:
  - Plain text (TXT)
  - Word document (DOCX)
  - Excel spreadsheet (XLSX)
- **Financial Document Intelligence**: Automatically detects and formats financial documents like invoices, receipts, and bills
- **Arabic/Urdu Support**: Includes automatic installation of Arabic and Urdu fonts for proper text rendering
- **Customizable**: Configure AWS region, profile, model, and output formats

## Prerequisites

1. **AWS Account** with access to Amazon Bedrock
2. **AWS CLI** configured with appropriate credentials
3. **Python 3.6+**
4. **Required Python packages**:
   ```
   pip install boto3 openpyxl fpdf python-docx requests
   ```

## Installation

1. Clone this repository or download the script
2. Install the required dependencies:
   ```
   pip install boto3 openpyxl fpdf python-docx requests
   ```
3. Ensure your AWS CLI is configured with appropriate credentials and permissions for Amazon Bedrock
4. (Optional) Run the font installer script to pre-install required fonts:
   ```
   python install-fonts.py
   ```

## Usage

Basic usage:
```
python image-extractor-fixed.py /path/to/your/document.jpg
```

Advanced options:
```
python image-extractor-fixed.py /path/to/your/document.pdf --region us-east-1 --profile my-aws-profile --model anthropic.claude-3-5-sonnet-20240620-v1:0 --formats txt,docx,xlsx
```

### Command Line Arguments

- `file_path`: Path to the image or PDF file (required)
- `--region`: AWS region for Bedrock (default: us-west-2)
- `--profile`: AWS profile name (default: default)
- `--model`: Bedrock model ID (default: anthropic.claude-3-5-sonnet-20240620-v1:0)
- `--formats`: Output formats as comma-separated list (default: txt,docx,xlsx)

## Font Installation

The tool includes automatic font installation for Arabic and Urdu text support. You can:

1. **Let the tool install fonts automatically** when needed (happens during first run)
2. **Pre-install fonts** using the included script:
   ```
   python install-fonts.py
   ```
3. **Install fonts manually** if automatic installation fails:
   - Amiri Regular: https://github.com/alif-type/amiri/raw/master/amiri-regular.ttf
   - Noto Sans Arabic: https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoSansArabic/NotoSansArabic-Regular.ttf
   - Noto Nastaliq Urdu: https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoNastaliqUrdu/NotoNastaliqUrdu-Regular.ttf

The fonts are installed to your user fonts directory (`~/Library/Fonts` on macOS).

## How It Works

The Document Text Extractor operates through the following process:

1. **Image/PDF Processing**: The tool reads the input file and determines its media type based on the file extension.
2. **Font Installation Check**: The tool checks for required Arabic/Urdu fonts and installs them if needed.
3. **AWS Bedrock Integration**: The file is encoded in base64 and sent to Amazon Bedrock's Claude model via the `invoke_model` API.
4. **Text Extraction**: Claude analyzes the visual content and extracts all visible text from the document.
5. **Document Type Analysis**: The extracted text is analyzed to determine if it's a financial document by searching for keywords like "invoice," "receipt," etc. in multiple languages.
6. **Structured Data Processing**: For financial documents, the text is processed into structured sections (header information, item details, totals).
7. **Format-Specific Output Generation**: The extracted text is formatted and saved in the requested output formats with appropriate styling.

The tool leverages Claude's advanced vision capabilities to accurately extract text from various document types, including those with complex layouts or multilingual content.

## Financial Document Detection and Processing

The financial document detection system works through these steps:

1. **Keyword Detection**: The tool scans the extracted text for financial keywords in multiple languages, including:
   - English: invoice, bill, receipt, statement, purchase order, etc.
   - Arabic: فاتورة, حساب, إيصال, أمر شراء, etc.

2. **Section Identification**: Once identified as a financial document, the text is processed using a state machine approach to identify three key sections:
   - **Header Section**: Contains document metadata like invoice number, date, company information
   - **Item Section**: Contains line items, quantities, prices, and descriptions
   - **Total Section**: Contains subtotals, taxes, and final amounts

3. **Pattern Recognition**: The tool uses regular expressions to identify:
   - Column headers (description, quantity, price, amount)
   - Currency amounts and numerical data
   - Key-value pairs (e.g., "Invoice Number: 12345")

4. **Intelligent Formatting**: Based on the identified structure, the tool creates:
   - Properly aligned tables for item details
   - Right-aligned numerical values
   - Bold formatting for important information like totals
   - Appropriate column widths based on content

This intelligent processing ensures that financial documents maintain their tabular structure and data relationships in all output formats.

## Performance Considerations

When using the Document Text Extractor, keep these performance factors in mind:

1. **File Size and Complexity**:
   - Larger files (>10MB) may take longer to process
   - Complex layouts with multiple columns, tables, and mixed text/graphics require more processing time
   - Consider compressing large images before processing

2. **AWS Bedrock Quotas and Limits**:
   - Be aware of your AWS Bedrock service quotas for API calls
   - Claude models have token limits (both input and output)
   - Very large documents may need to be split into multiple pages

3. **Network Considerations**:
   - Ensure stable internet connection for API calls
   - Processing time includes network latency for sending files to AWS

4. **Memory Usage**:
   - Processing large PDF files may require significant memory
   - For multi-page documents, consider processing one page at a time

5. **Cost Optimization**:
   - Claude API calls are billed based on input and output tokens
   - Consider using smaller image resolutions when possible
   - Batch processing multiple documents can be more efficient than individual calls

6. **Font Installation**:
   - First-time font installation may add processing time
   - Subsequent runs will be faster as fonts are already installed
   - The tool now installs fonts in a background thread to avoid blocking the main process

For optimal performance, we recommend processing files under 5MB and ensuring your AWS account has appropriate rate limits for your expected usage volume.

## Customizing the Claude Prompt

The tool uses a default prompt to instruct Claude on how to extract text from documents. You can customize this prompt by modifying the script:

1. **Default Prompt**: The current default prompt is simple and direct:
   ```python
   "Please extract all the text from this document."
   ```

2. **Customization Options**: You can modify the prompt in the `main()` function to provide more specific instructions, such as:
   - Focusing on specific sections of the document
   - Requesting particular formatting or organization of the extracted text
   - Asking for additional analysis of the document content
   - Specifying how to handle tables, charts, or other non-text elements

3. **Example Custom Prompts**:
   ```python
   # For financial documents
   "Please extract all text from this invoice. Organize it into sections for header information, line items, and totals."
   
   # For multilingual documents
   "Extract all text from this document, preserving both English and Arabic content. Indicate which sections are in which language."
   
   # For forms
   "Extract all form fields and their values from this document. Format as field:value pairs."
   ```

4. **Implementation**: To customize the prompt, locate this section in the code:
   ```python
   request_body = {
       "anthropic_version": "bedrock-2023-05-31",
       "max_tokens": 4000,
       "temperature": 0,
       "messages": [
           {
               "role": "user",
               "content": [
                   {
                       "type": "text",
                       "text": "Please extract all the text from this document."  # Modify this line
                   },
                   {
                       "type": "image",
                       "source": {
                           "type": "base64",
                           "media_type": media_type,
                           "data": base64.b64encode(file_content).decode('utf-8')
                       }
                   }
               ]
           }
       ]
   }
   ```

5. **Best Practices**:
   - Keep prompts clear and specific
   - Test different prompts to find what works best for your document types
   - Consider adding command-line options to select different pre-defined prompts

Customizing the prompt can significantly improve extraction quality for specific document types or use cases.

## Output Files

The tool generates output files in the same directory as the input file, with the following naming pattern:
- `[original_filename]_extracted.txt`
- `[original_filename]_extracted.docx`
- `[original_filename]_extracted.xlsx`

## AWS Permissions Required

- `bedrock:InvokeModel` permission for the Claude model being used

## Limitations

- Currently only supports Claude models from Amazon Bedrock
- PDF extraction quality depends on the PDF's content (scanned vs. digital)
- Some complex document layouts may not be perfectly preserved
- PDF output generation is currently disabled due to compatibility issues

## Troubleshooting

- **Font Issues**: 
  - If Arabic/Urdu text doesn't display correctly, run `python install-fonts.py` to install fonts
  - The tool will attempt to install fonts automatically, but may require manual installation in some cases
  - Check the console output for font installation status and errors
- **AWS Errors**: Ensure your AWS credentials have access to Amazon Bedrock and the specified model
- **Missing Dependencies**: Install any missing Python packages as prompted
- **Download Issues**: The tool uses both the `requests` library and `curl` as fallbacks for font downloads
- **PDF Generation**: PDF output is currently disabled due to compatibility issues. Use the other output formats (TXT, DOCX, XLSX) instead.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
