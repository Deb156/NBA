import os
import json
import time
import shutil
import smtplib
import threading
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from jinja2 import Template
from openpyxl import Workbook, load_workbook
from email.mime.base import MIMEBase
from email import encoders



class FileMonitor:
    def __init__(self, config_path='config.json'):
        with open(config_path, 'r') as f:
            self.config = json.load(f)
        
        self.dropped_files = []
        self.transfer_log = []
        self.performance_tracking = {}  # Track each asset's transfer progress
        self.running = True
        self.excel_file = 'NBA_Transfer_Performance.xlsx'
        self.serial_counter = 1
        
        # Create directories if they don't exist
        for folder in self.config['watch_folders'].values():
            os.makedirs(folder, exist_ok=True)
        for folder in self.config['destination_folders'].values():
            os.makedirs(folder, exist_ok=True)

    def count_files_in_folder(self, folder_path):
        if not os.path.exists(folder_path):
            return 0
        count = 0
        for root, dirs, files in os.walk(folder_path):
            count += len(files)
        return count

    def is_transfer_ongoing(self, file_path):
        # Check if file size is changing (simple transfer detection)
        try:
            size1 = os.path.getsize(file_path)
            time.sleep(0.1)
            size2 = os.path.getsize(file_path)
            return size1 != size2
        except:
            return False

    # The script will print email alerts to console instead of sending emails when the password is set to "your_password"
    # if sender_password="your_password" in config.json

    def get_asset_type(self, filename):
        ext = os.path.splitext(filename)[1].lower()
        if ext in ['.mp4', '.avi', '.mov', '.mkv', '.wmv']:
            return 'Video'
        elif ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp']:
            return 'Image'
        elif ext in ['.mp3', '.wav', '.flac', '.aac']:
            return 'Audio'
        elif ext in ['.txt', '.doc', '.docx', '.pdf']:
            return 'Document'
        else:
            return 'File'
    
    def track_asset_drop(self, asset_name, source_folder, drop_time, is_folder=False):
        asset_key = f"{source_folder}_{asset_name}"
        asset_type = 'Folder' if is_folder else self.get_asset_type(asset_name)
        
        self.performance_tracking[asset_key] = {
            'sl_no': self.serial_counter,
            'asset_name': asset_name,
            'asset_type': asset_type,
            'source_location': source_folder,
            'upload_time': drop_time,
            'destinations': {},
            'remarks': ''
        }
        self.serial_counter += 1
    
    def update_transfer_status(self, asset_name, source_folder, dest_name, transfer_time=None):
        asset_key = f"{source_folder}_{asset_name}"
        if asset_key in self.performance_tracking:
            if transfer_time:
                self.performance_tracking[asset_key]['destinations'][dest_name] = transfer_time
            else:
                self.performance_tracking[asset_key]['destinations'][dest_name] = 'In Progress'
    
    def update_performance_data(self, asset_name, asset_type, count, start_time, transfer_times):
        end_time = max(transfer_times.values()) if transfer_times else start_time
        total_time = (end_time - start_time).total_seconds() / 60  # in minutes
        
        self.performance_data.append({
            'Asset Name': asset_name,
            'Asset Type': asset_type,
            'Count': count,
            'Start Time': start_time.strftime('%Y-%m-%d %H:%M:%S'),
            'End Time': end_time.strftime('%Y-%m-%d %H:%M:%S'),
            'Total Time (minutes)': round(total_time, 2),
            'Transfer Details': ', '.join([f"{k}: {v.strftime('%H:%M:%S')}" for k, v in transfer_times.items()])
        })
    
    def create_excel_report(self):
        try:
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            
            wb = Workbook()
            ws = wb.active
            ws.title = 'NBA Transfer Performance'
            
            # Headers with formatting
            headers = ['SL.NO', 'Asset Name', 'Asset Type', 'Source Drop Location', 
                      'Asset Upload Time', 'Destination Location', 'Asset Transferred Time', 
                      'Time Difference Analysis', 'Remarks']
            ws.append(headers)
            
            # Format headers - bold, background color, alignment
            header_font = Font(bold=True, color='FFFFFF')
            header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
            header_alignment = Alignment(horizontal='center', vertical='center')
            
            for cell in ws[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
            
            # Set column widths
            column_widths = [8, 25, 12, 20, 20, 18, 20, 20, 30]
            for i, width in enumerate(column_widths, 1):
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width
            
            # Add data rows
            for asset_key, data in self.performance_tracking.items():
                destinations = self.config['destination_mapping'][data['source_location']]
                
                for dest_name in destinations:
                    transfer_time = data['destinations'].get(dest_name, 'In Progress')
                    
                    if transfer_time != 'In Progress':
                        time_diff = (transfer_time - data['upload_time']).total_seconds() / 60
                        time_analysis = f"{time_diff:.2f} minutes"
                        transfer_time_str = transfer_time.strftime('%Y-%m-%d %H:%M:%S')
                    else:
                        time_analysis = 'In Progress'
                        transfer_time_str = 'In Progress'
                    
                    ws.append([
                        data['sl_no'],
                        data['asset_name'],
                        data['asset_type'],
                        data['source_location'],
                        data['upload_time'].strftime('%Y-%m-%d %H:%M:%S'),
                        dest_name,
                        transfer_time_str,
                        time_analysis,
                        data['remarks']
                    ])
            
            # Add borders to all cells with data
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=len(headers)):
                for cell in row:
                    cell.border = thin_border
                    if cell.row > 1:  # Data rows alignment
                        cell.alignment = Alignment(horizontal='left', vertical='center')
            
            wb.save(self.excel_file)
        except Exception as e:
            print(f"Excel creation failed: {e}")
    
    def send_email(self, subject, body, alert_type="Alert", details=None, attach_excel=False):
        try:
            if self.config['email']['sender_password'] == 'your_password':
                print(f"EMAIL ALERT: {subject}\n{body}\n")
                return
            
            # Load and render HTML template
            with open('email_template.html', 'r', encoding='utf-8') as f:
                template = Template(f.read())
            
            alert_class = "danger" if "Alert" in alert_type else "info" if "Report" in alert_type else "success"
            
            html_content = template.render(
                subject=subject,
                message=body,
                alert_type=alert_type,
                alert_class=alert_class,
                details=details,
                timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            )
                
            msg = MIMEMultipart('alternative')
            msg['From'] = self.config['email']['sender_email']
            msg['To'] = ', '.join(self.config['email']['recipients'])
            msg['Subject'] = subject
            
            msg.attach(MIMEText(body, 'plain', 'utf-8'))
            msg.attach(MIMEText(html_content, 'html', 'utf-8'))
            
            # Attach Excel file if requested
            if attach_excel and os.path.exists(self.excel_file):
                with open(self.excel_file, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename="{self.excel_file}"')
                    msg.attach(part)
            
            server = smtplib.SMTP(self.config['email']['smtp_server'], self.config['email']['smtp_port'])
            server.starttls()
            server.login(self.config['email']['sender_email'], self.config['email']['sender_password'])
            server.send_message(msg)
            server.quit()
        except Exception as e:
            print(f"Email send failed: {e}")
            print(f"EMAIL ALERT: {subject}\n{body}\n")



class WatchHandler(FileSystemEventHandler):
    def __init__(self, monitor, folder_name):
        self.monitor = monitor
        self.folder_name = folder_name

    def on_created(self, event):
        if not event.is_directory:
            file_path = event.src_path
        else:
            file_path = event.src_path
            
        # Record the drop
        file_count = 0
        if os.path.isdir(file_path):
            file_count = self.monitor.count_files_in_folder(file_path)
            if file_count == 0:
                # Send blank folder notification
                self.monitor.send_email(
                    "Blank Folder Alert",
                    f"Blank folder detected:\nSource: {self.folder_name}\nFolder: {os.path.basename(file_path)}\nTime: {datetime.now()}"
                )
        
        drop_time = datetime.now()
        asset_name = os.path.basename(file_path)
        
        self.monitor.dropped_files.append({
            'time': drop_time,
            'source_folder': self.folder_name,
            'asset_name': asset_name,
            'asset_type': 'folder' if os.path.isdir(file_path) else 'file',
            'file_count': file_count,
            'path': file_path
        })
        
        # Track asset drop for performance monitoring
        self.monitor.track_asset_drop(asset_name, self.folder_name, drop_time, os.path.isdir(file_path))
        
        # Send notification about file drop
        details = {
            'Source Folder': self.folder_name,
            'Asset Name': os.path.basename(file_path),
            'Asset Type': 'Folder' if os.path.isdir(file_path) else 'File',
            'File Count': file_count if os.path.isdir(file_path) else 1
        }
        
        self.monitor.send_email(
            f"File Drop Alert - {self.folder_name}",
            f"New {details['Asset Type'].lower()} detected in {self.folder_name} watch folder.",
            "File Drop Alert",
            details
        )

def validation_worker(monitor):
    while monitor.running:
        time.sleep(monitor.config['intervals']['validation_minutes'] * 60)
        
        # Check transfer status
        for item in monitor.dropped_files:
            if os.path.exists(item['path']):
                source_folder = item['source_folder']
                destinations = monitor.config['destination_mapping'][source_folder]
                
                transfer_status = {}
                for dest_name in destinations:
                    dest_path = monitor.config['destination_folders'][dest_name]
                    expected_file = os.path.join(dest_path, item['asset_name'])
                    transfer_status[dest_name] = os.path.exists(expected_file)
                
                # Send transfer status notification
                status_msg = f"Transfer status for {item['asset_name']} from {source_folder}:"
                transfer_times = {}
                for dest, status in transfer_status.items():
                    status_msg += f"\n{dest}: {'[OK] Transferred' if status else '[MISSING] Not Found'}"
                    if status:
                        transfer_time = datetime.now()
                        transfer_times[dest] = transfer_time
                        # Update performance tracking
                        monitor.update_transfer_status(item['asset_name'], item['source_folder'], dest, transfer_time)
                    else:
                        # Mark as in progress
                        monitor.update_transfer_status(item['asset_name'], item['source_folder'], dest)
                
                # Update performance data if transfers are complete
                if all(transfer_status.values()):
                    count = item['file_count'] if item['asset_type'] == 'folder' else 'NA'
                    monitor.update_performance_data(
                        item['asset_name'], 
                        item['asset_type'].title(), 
                        count, 
                        item['time'], 
                        transfer_times
                    )
                
                monitor.send_email(
                    f"Transfer Status - {item['asset_name']}",
                    status_msg,
                    "Transfer Status Report",
                    transfer_status
                )
        
        # Check for files older than threshold
        threshold = datetime.now() - timedelta(hours=monitor.config['alert_threshold_hours'])
        for item in monitor.dropped_files:
            if item['time'] < threshold and os.path.exists(item['path']):
                details = {
                    'Source Folder': item['source_folder'],
                    'Asset Name': item['asset_name'],
                    'Drop Time': item['time'].strftime('%Y-%m-%d %H:%M:%S'),
                    'Duration': str(datetime.now() - item['time'])
                }
                monitor.send_email(
                    "File Stuck Alert",
                    f"File/folder stuck for >1 hour in {item['source_folder']} watch folder.",
                    "Stuck File Alert",
                    details
                )

def report_worker(monitor):
    while monitor.running:
        time.sleep(monitor.config['intervals']['report_minutes'] * 60)
        
        # Generate structured report data
        dropped_summary = []
        for folder_name in monitor.config['watch_folders'].keys():
            folder_drops = [f for f in monitor.dropped_files if f['source_folder'] == folder_name]
            dropped_summary.append({'Source Folder': folder_name, 'Items Dropped': len(folder_drops)})
        
        dest_counts = []
        for dest_name, dest_path in monitor.config['destination_folders'].items():
            count = monitor.count_files_in_folder(dest_path)
            dest_counts.append({'Destination Folder': dest_name, 'File Count': count})
        
        # Create structured details for HTML template
        report_details = {
            'Dropped Files Summary': dropped_summary,
            'Destination Folder Counts': dest_counts,
            'Total Transfers': len(monitor.transfer_log)
        }
        
        report_body = f"NBA Monitor Report - {len(monitor.dropped_files)} total items monitored, {len(monitor.transfer_log)} transfers completed."
        
        monitor.send_email("NBA Monitor Report", report_body, "Monitor Report", report_details)

def performance_report_worker(monitor):
    while monitor.running:
        time.sleep(monitor.config['intervals']['performance_report_minutes'] * 60)
        
        monitor.create_excel_report()
        
        if os.path.exists(monitor.excel_file):
            wb = load_workbook(monitor.excel_file)
            ws = wb.active
            
            report_data = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0]:
                    report_data.append({
                        'Asset Name': row[0],
                        'Asset Type': row[1], 
                        'Count': row[2],
                        'Start Time': row[3],
                        'End Time': row[4],
                        'Total Time': f"{row[5]} min"
                    })
            
            report_body = f"Performance Report - Total {len(report_data)} transfers tracked. See attached Excel file."
            
            monitor.send_email(
                "NBA Performance Report", 
                report_body, 
                "Performance Report", 
                {'Total Transfers': len(report_data)},
                attach_excel=True
            )

def main():
    monitor = FileMonitor()
    
    # Start validation and report workers
    threading.Thread(target=validation_worker, args=(monitor,), daemon=True).start()
    threading.Thread(target=report_worker, args=(monitor,), daemon=True).start()
    threading.Thread(target=performance_report_worker, args=(monitor,), daemon=True).start()
    
    # Setup file watchers
    observer = Observer()
    for folder_name, folder_path in monitor.config['watch_folders'].items():
        handler = WatchHandler(monitor, folder_name)
        observer.schedule(handler, folder_path, recursive=True)
    
    observer.start()
    print("NBA File Monitor started...")
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        monitor.running = False
        observer.stop()
    
    observer.join()

if __name__ == "__main__":
    main()