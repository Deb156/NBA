# NBA File Monitor

A Python script that monitors watch folders and handles automated file transfers with reporting.

## Setup
1. Install dependencies: `pip install -r requirements.txt`
2. Configure `config.json` with your folder paths and email settings
3. Run: `python run.py`

## Features
- Monitors 3 configurable watch folders
- Automatic file/folder copying and moving
- Email notifications for alerts and reports
- Configurable validation and reporting intervals
- Blank folder detection
- Stuck file alerts (>1 hour)

## Configuration
Edit `config.json` to set:
- Watch folder paths
- Destination folder paths  
- Email settings
- Time intervals