OpenWeather Automation ⛅📊
Automation system for collecting weather data, storing historical logs, generating professional reports (XLSX + PDF), and sending them automatically via email.

🚀 Overview
This project executes a complete weather-data automation pipeline:
  Loads .env configuration
  Fetches weather data from OpenWeather API
  Validates and parses the response
  Ensures the CSV log exists
  Appends the newest record
  Loads & sorts the entire historical dataset

Generates:
  Excel Report (.xlsx)
  PDF Report (.pdf)
  Sends both files by email
  Returns execution status

All steps are orchestrated by the run_automation() function inside main.py.

🧰 Technologies Used
  Python 3
  Requests → API consumption
  CSV → persistent logging
  OpenPyXL → Excel report creation
  FPDF → PDF report generation
  Dotenv → environment configuration
  SMTP (SSL) → email sending

📂 Project Structure
/project-root
│
├── main.py                 # Full automation pipeline
├── send_email.py           # Email sending + attachments
├── weather_log.csv         # Auto-created historical log
├── requirements.txt
└── .env                    # You must create this file

🔧 Installation & Setup
1️⃣ Clone the repository
  git clone https://github.com/your-username/openweather-automation.git
  cd openweather-automation

2️⃣ Create a .env file
  OPENWEATHER_API_KEY=your_api_key_here
  DEFAULT_CITY=YourCityName
  
  SMTP_USER=youremail@example.com
  SMTP_PASS=your_password
  SMTP_SERVER=smtp.gmail.com
  SMTP_PORT=465
  
  MAIL_TO=recipient@example.com
  
  ⚠️ Gmail users must enable "App Passwords" when using 2FA.

3️⃣ Install dependencies
  pip install -r requirements.txt

4️⃣ Run the automation
  python main.py

📌 Automation Pipeline (How It Works)
🔄 Executed inside run_automation():
  ✔️ Load environment variables
  ✔️ Fetch weather data
  ✔️ Abort if request fails
  ✔️ Ensure weather_log.csv exists
  ✔️ Append the new record
  ✔️ Read + sort the entire log
  ✔️ Create /reports/YYYY-MM-DD/ folder
  ✔️ Generate XLSX report
  ✔️ Generate PDF report
  ✔️ Send email with attachments
  ✔️ Return True

📄 Reports Generated
1️⃣ Excel Report (.xlsx)
  Styled header
  Zebra rows
  Borders + cell alignment
  Auto column width
  Frozen header row
  Column filters enabled
  Full historical dataset

2️⃣ PDF Report (.pdf) contains:
  Timestamp of report generation
  Total number of records
  First recorded measurement
  Last recorded measurement
  Last 3 consultations
  Clean, vertical formatting
  Files are saved under:
    /reports/YYYY-MM-DD/weather_report.xlsx
    /reports/YYYY-MM-DD/weather_report.pdf

📤 Email Sending
  The function send_weather_report():
  Builds a multipart email
  Includes a text summary (temperature + city)
  Attaches the XLSX & PDF files
  Sends everything using SMTP_SSL
  Email delivery settings come from .env.

📍 Future Improvements
  Automatic scheduling (cron / Task Scheduler)
  Multi-city reporting
  HTML dashboard with charts
  Web interface to trigger automation
  Alerts for extreme weather
  Cloud backup (S3 / GDrive)

🧑‍💻 Author
  Matheus Cartonilho
  Full-Stack & Python Developer
  Porto Velho — RO, Brazil
