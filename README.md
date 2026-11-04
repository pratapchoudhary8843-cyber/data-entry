# data-entry
📁 Smart Data Entry – Intelligent File to Excel Converter

Upload Any File. Extract Data Automatically. Edit, Search & Export to Excel.

Smart Data Entry is a lightweight, intelligent data extraction and management application that converts information from different file formats into structured, editable tables.

It supports automatic OCR (Optical Character Recognition), duplicate detection, data analysis, multiple Excel sheets, and file storage — all through a modern, mobile-friendly interface.

✨ Key Features
📂 1. Upload Any File
Upload multiple files simultaneously.
Drag and drop files.
Copy and paste files directly.
Capture photos using your device camera.
Supports:
JPG, JPEG, PNG, WEBP, BMP, GIF, TIFF
PDF
Word (DOCX)
Excel (XLSX, XLSM)
CSV, TSV
JSON
TXT, MD, LOG
Other file types (stored safely).
🤖 2. Automatic Data Extraction
OCR-powered text recognition from images.
Extract text from PDF documents.
Read Word and Excel files.
Automatically detect table structures.
Identify headers and columns.
Convert extracted content into structured rows.
📊 3. Smart Data Detection

Automatically identifies:

Email addresses
Phone numbers
Dates
Numeric values
Minimum and maximum values
Sum of detected numbers
Document categories, such as invoices, receipts, forms, and data tables.
🗃️ 4. Smart Storage
Each uploaded file gets its own Excel sheet.
Original files are preserved.
Automatic duplicate file detection using SHA-1.
Store data locally in the browser or through a Python server.
IndexedDB storage in browser mode.
Excel workbook storage in server mode.
✏️ 5. Edit & Manage Data
Edit individual cells.
Rename column headers.
Add new rows.
Add new columns.
Delete unwanted rows.
Save modifications.
Preview extracted information.
🔎 6. Search & Export
Search across all stored records.
Download all records in one Excel workbook.
Export individual sheets as CSV.
Export individual sheets as JSON.
Download original uploaded files.
Copy table data directly into Excel.
📱 7. Mobile & PWA Support
Responsive mobile-friendly interface.
Installable Progressive Web App (PWA).
Camera integration.
iPhone Add to Home Screen support.
Android installation support.
Access from a phone over the same Wi-Fi network using Python server mode.
🌌 8. Modern Interface
Dark futuristic interface.
Animated particle background.
Responsive dashboard.
Live processing indicators.
Upload status notifications.
Image preview and zoom.
🛠️ Technology Stack
Technology	Purpose
HTML5	Application structure
CSS3	Responsive UI and styling
JavaScript	Application logic
IndexedDB	Browser-side storage
Flask	Python backend
Python	Server-side processing
OpenPyXL	Excel workbook management
Pillow	Image processing
NumPy	Image data processing
RapidOCR	Text recognition
Tesseract.js	Browser OCR
PDF.js	Browser PDF extraction
JSZip	DOCX extraction
SheetJS (XLSX)	Excel file processing
📥 Installation & Setup
Requirements
Python 3.9 or newer recommended.
Modern web browser.
Internet connection for initial browser-side library downloads.
1. Clone the Repository
bash
git clone https://github.com/YOUR_USERNAME/smart-data-entry.git

bash
cd smart-data-entry

2. Install Python Dependencies
bash
pip install flask pillow openpyxl numpy rapidocr-onnxruntime


For PDF support:

bash
pip install pypdf pypdfium2


Optional QR code support:

bash
pip install qrcode

3. Run the Application
bash
python app.py


The application automatically:

Finds an available port between 5000 and 5029.
Starts the Flask server.
Opens the application in your default browser.
Checks server connectivity.
Detects OCR availability.
4. Open the Application

The terminal displays an address similar to:

text
http://127.0.0.1:5000


Open the displayed address in your browser.

📱 Access from Your Mobile Phone

You can use Smart Data Entry on your phone through the same Wi-Fi network.

Steps
Start the Python server with:
bash
python app.py --phone

The terminal displays a network address and a PIN.
Connect your phone and computer to the same Wi-Fi.
Open the displayed phone access link.
Enter the PIN if requested.
Start uploading files directly from your phone.

Security: Anyone on your local network who has the access link and PIN can use the application. Do not share these credentials with untrusted people.

🌐 Browser Mode

Smart Data Entry also supports standalone browser mode.

Features:

No Python server required.
Uses IndexedDB for local storage.
OCR through Tesseract.js.
Excel processing through SheetJS.
PDF processing through PDF.js.
DOCX extraction through JSZip.

Open the application through a supported local web server or hosted website.

In browser mode, files remain in that browser's local storage rather than being uploaded to the Python backend.

📂 Project Structure
text
smart-data-entry/
│
├── app.py                  # Python Flask backend
├── index.html              # Main application interface
│
├── data/
│   ├── entries.xlsx        # Main Excel database
│   ├── files/              # Original uploaded files
│   └── photos/             # Image storage directory
│
├── manifest.webmanifest    # PWA configuration
├── sw.js                   # Service worker
├── icon-192.png            # App icon
├── apple-touch-icon.png    # iOS app icon
│
└── README.md               # Project documentation

📊 How It Works
text
       Upload File
            |
            ▼
     File Type Detection
            |
            ▼
    Automatic Extraction
            |
      ┌─────┴──────┐
      ▼            ▼
     OCR       File Reader
      │            │
      └─────┬──────┘
            |
            ▼
     Table Generation
            |
            ▼
     Smart Data Detection
            |
            ▼
      Duplicate Checking
            |
            ▼
       Data Storage
            |
       ┌────┴─────┐
       ▼          ▼
     Editing    Searching
       │          │
       └────┬─────┘
            |
            ▼
       Excel Export

💾 Data Storage

The application supports two storage modes:

Feature	Browser Mode	Python Server
Storage	IndexedDB	Excel workbook
OCR	Tesseract.js	RapidOCR / Tesseract
Image extraction	Yes	Yes
PDF extraction	Yes	Yes
Excel export	Yes	Yes
Original files	Yes	Yes
Mobile access	Browser/device dependent	Same Wi-Fi network
Backend required	No	Yes

Note: Browser storage is device- and browser-specific. Clearing browser data may delete locally stored records. Keep exported backups.

🔐 Privacy & Security
Browser mode processes files locally in the browser.
Python server mode stores uploaded files on the computer running the server.
Optional PIN protection is available for mobile network access.
Duplicate detection helps prevent accidental repeated uploads.

Avoid exposing the development server directly to the public internet without additional security controls.

⚠️ Important Notes
OCR accuracy depends on image quality, resolution, orientation, and document layout.
Complex tables may require manual corrections.
Some file formats are stored as-is when readable content cannot be extracted.
Browser mode requires internet access to download external processing libraries on first use.
The Python server currently accepts uploads up to 50 MB.
PWA installation requires appropriate web hosting and HTTPS in normal deployment scenarios.
🚀 Future Improvements
AI-powered document classification.
Multilingual OCR support.
Automatic invoice and receipt processing.
Advanced Excel analytics.
Cloud database integration.
User authentication and role management.
Bulk document processing.
Improved table and column recognition.
👨‍💻 Author

Smart Data Entry

An intelligent document-to-data management solution designed to simplify file processing, data extraction, and Excel management.

⭐ Support

If you find this project useful:

Give the repository a ⭐ Star.
Fork the project.
Report bugs through GitHub Issues.
Contribute improvements through Pull Requests.

Made with ❤️ using Python, Flask, JavaScript, and OCR technology.
