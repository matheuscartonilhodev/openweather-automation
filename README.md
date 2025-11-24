OpenWeather Automation

Automation system for collecting weather data, storing historical results, and generating structured reports in XLSX and PDF formats.

🚀 Overview

This project retrieves real-time weather information from the OpenWeather API, logs each query into a CSV history file, and automatically generates two types of reports:

Excel Report (.xlsx) — with formatted headers, zebra rows, alignment rules, filters, and automatic column sizing.

PDF Report (.pdf) — containing general statistics (total records, first and last measurements) and a summary of the last three consultations.

It is designed to simulate a real-world automation workflow: API consumption → Data persistence → Report generation → Optional email sending.

🧰 Technologies Used

Python 3
Requests (API consumption)
CSV (data persistence)
OpenPyXL (Excel report generation)
FPDF (PDF report generation)
Dotenv (environment variables)
SMTP (optional) — if email sending is activated

📦 Project Structure
/project-root
│
├── main.py
├── send_email.py
├── weather_log.csv
├── requirements.txt
└── .env   (you must create this file)

🔧 Installation & Setup

Clone the repository

git clone https://github.com/your-username/openweather-automation.git
cd openweather-automation


Create the .env file
OPENWEATHER_API_KEY=your_api_key_here


Install dependencies
pip install -r requirements.txt


Run the automation
python main.py

📌 How It Works

The user runs the script.
The script asks for a city name.
The system sends a request to the OpenWeather API.
The response is validated and parsed.
The data is appended to weather_log.csv, keeping a historical timeline.

The XLSX report is created or updated:
Styled headers
Borders
Alignment rules
Zebra rows
Auto column width
Frozen header row
Column filters
The PDF report is generated containing:
Total number of records
First and last consultation timestamps
The last three weather measurements, formatted
(Optional) The data can be emailed using send_email.py.

📄 Reports Generated
XLSX Report
Full historical dataset
All columns formatted
Easy filtering and sorting
Professional table look

PDF Report
Includes:
Generation timestamp
Total records
First measurement date
Last measurement date
Last 3 consultations (datetime, city, temperature, humidity, description)

📨 Email Sending (Optional)
If configured, the project can automatically send:
The API result
Or the generated reports
Or both
The function is available in send_email.py.

📍 Future Improvements (Roadmap)
Background scheduler (CRON / Task Scheduler)
Automatic daily reports
Dashboard (HTML + JS)
Multi-city batch queries
Weather alerts system

🧑‍💻 Author
Matheus Cartonilho
Full-Stack & Python Developer
Porto Velho — RO, Brazil