 Employee Weekly Timesheet Summary & Email Notification Tool  
Project Overview  
This project is a web-based tool built with Flask (Python), HTML/CSS/JavaScript, and Pandas for data processing.  
It allows employees to upload a weekly Excel timesheet (`weekly_timesheet.xlsx`), analyzes the data, and automatically sends a summary email with:  

- Current week statistics (Category, Sub-Category, Line Items, Planned vs Actual).  
- Next week’s planned work.  
- Deviation summary (Planned vs Actual efforts).  

---

requirements 
   Flask==3.0.3
Flask-Cors==4.0.0
pandas==2.2.2
numpy==1.26.4
python-dotenv==1.0.1
openpyxl==3.1.5
How to Run the Project
Follow these steps to set up and run the Employee Weekly Timesheet Summary & Email Notification Tool:
 Clone the Project & Enter Folder
```bash
git clone https://github.com/your-username/timesheet-tool.git
cd timesheet-tool
create virtual environment
python -m venv venv
Install Dependencies
     pip install -r requirements.txt
Configure Environment Variables
Run the Flask App
    python app.py
You should see output when we click  http://127.0.0.1:5000/   
