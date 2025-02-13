import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkcalendar import DateEntry
import math
from tkinter.scrolledtext import ScrolledText
from tabulate import tabulate

class TimeEntry(ttk.Frame):
    """Custom time entry widget with hour, minute, and AM/PM selection"""
    def __init__(self, parent, *args, default_ampm="AM", **kwargs):
        super().__init__(parent, *args, **kwargs)
        
        self.hour_var = tk.StringVar()
        self.minute_var = tk.StringVar()
        self.ampm_var = tk.StringVar(value=default_ampm)
        
        self.hour_cb = ttk.Combobox(
            self, textvariable=self.hour_var,
            values=[f"{i:02d}" for i in range(1, 13)], width=3
        )
        self.hour_cb.grid(row=0, column=0, padx=2)
        
        ttk.Label(self, text=":").grid(row=0, column=1)
        
        self.minute_cb = ttk.Combobox(
            self, textvariable=self.minute_var,
            values=["00", "15", "30", "45"], width=3
        )
        self.minute_cb.grid(row=0, column=2, padx=2)
        
        self.ampm_cb = ttk.Combobox(
            self, textvariable=self.ampm_var,
            values=["AM", "PM"], width=3
        )
        self.ampm_cb.grid(row=0, column=3, padx=2)
        
        self.hour_var.set("08")
        self.minute_var.set("30")
        
    def get(self):
        """Return time in HH:MM AM/PM format"""
        return f"{self.hour_var.get()}:{self.minute_var.get()} {self.ampm_var.get()}"

class HolidayManager:
    def __init__(self):
        self.holidays = set()  # Store holidays as a set of datetime objects
    
    def add_holiday(self, date):
        """Add a holiday date"""
        self.holidays.add(date)
    
    def remove_holiday(self, date):
        """Remove a holiday date"""
        self.holidays.discard(date)
    
    def is_holiday(self, date):
        """Check if a date is a holiday"""
        return date in self.holidays
    
    def get_holidays(self):
        """Return list of holidays"""
        return sorted(list(self.holidays))

class TimeCardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Time Calculator")
        try:
            self.root.iconbitmap(r'C:\Users\r123t\Desktop\New folder\time-card.ico')
        except:
            pass  # If icon file not found, continue without it
        self.root.geometry("900x700")
        
        self.total_hours = 0.0
        self.total_overtime = 0.0
        self.total_regular_hours = 0.0
        
        self.holiday_manager = HolidayManager()
        
        self.setup_styles()
        self.create_ui()
    
    def setup_styles(self):
        style = ttk.Style()
        style.configure("Treeview", rowheight=25)
        style.configure("Treeview.Heading", font=('Arial', 10, 'bold'))
    
    def create_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.create_input_frame(main_frame)
        self.create_holiday_frame(main_frame)
        self.create_button_frame(main_frame)
        self.create_tree_frame(main_frame)
        
        # Footer labels
        self.vision_label = ttk.Label(self.root, text="© Dark Killer " "- Version-1.0.0", anchor="e", relief=tk.SUNKEN, padding=(5, 10))
        self.vision_label.pack(side=tk.BOTTOM, fill=tk.X)
        
        # self.copyright_label = ttk.Label(self.root, text="© Dark Killer" "- Version-1.0.0", anchor="center", relief=tk.SUNKEN, padding=(5, 25))
        # self.copyright_label.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.total_label = ttk.Label(self.root, text="Total Work Hours: 0.00 | Total Overtime: 0.00", anchor="center", relief=tk.SUNKEN, padding=(5, 2))
        self.total_label.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_bar = ttk.Label(self.root, text="Ready", anchor="center", relief=tk.SUNKEN, padding=(5, 2))
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def create_input_frame(self, parent):
        input_frame = ttk.LabelFrame(parent, text="Time Entry", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 10))

        # Date picker
        ttk.Label(input_frame, text="Date:").grid(row=0, column=0, padx=5, pady=5)
        self.date_picker = DateEntry(input_frame, width=12)
        self.date_picker.grid(row=0, column=1, padx=5, pady=5)

        # In time entry
        ttk.Label(input_frame, text="In Time:").grid(row=0, column=2, padx=5, pady=5)
        self.in_time_entry = TimeEntry(input_frame, default_ampm="AM")
        self.in_time_entry.grid(row=0, column=3, padx=5, pady=5)

        # Out time entry
        ttk.Label(input_frame, text="Out Time:").grid(row=0, column=4, padx=5, pady=5)
        self.out_time_entry = TimeEntry(input_frame, default_ampm="PM")
        self.out_time_entry.grid(row=0, column=5, padx=5, pady=5)

        # Add entry button
        ttk.Button(input_frame, text="Add Entry", command=self.add_entry).grid(row=0, column=6, padx=5, pady=5)

    def create_holiday_frame(self, parent):
        holiday_frame = ttk.LabelFrame(parent, text="Holiday Management", padding="10")
        holiday_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Holiday date picker
        ttk.Label(holiday_frame, text="Holiday Date:").grid(row=0, column=0, padx=5, pady=5)
        self.holiday_picker = DateEntry(holiday_frame, width=12)
        self.holiday_picker.grid(row=0, column=1, padx=5, pady=5)
        
        # Holiday description entry
        ttk.Label(holiday_frame, text="Description:").grid(row=0, column=2, padx=5, pady=5)
        self.holiday_description = ttk.Entry(holiday_frame, width=20)
        self.holiday_description.grid(row=0, column=3, padx=5, pady=5)
        
        # Add/Remove holiday buttons
        ttk.Button(holiday_frame, text="Add Holiday", command=self.add_holiday).grid(
            row=0, column=4, padx=5, pady=5)
        ttk.Button(holiday_frame, text="View Holidays", command=self.view_holidays).grid(
            row=0, column=5, padx=5, pady=5)
    
    def add_holiday(self):
        date = self.holiday_picker.get_date()
        description = self.holiday_description.get().strip()
        
        if not description:
            messagebox.showwarning("Warning", "Please enter a holiday description")
            return
        
        self.holiday_manager.add_holiday(date)
        messagebox.showinfo("Success", f"Added holiday: {date.strftime('%Y-%m-%d')} - {description}")
        self.holiday_description.delete(0, tk.END)
    
    def view_holidays(self):
        holidays = self.holiday_manager.get_holidays()
        
        holiday_window = tk.Toplevel(self.root)
        holiday_window.title("Holidays")
        holiday_window.geometry("400x300")
        
        # Create Treeview for holidays
        columns = ("Date", "Day of Week")
        tree = ttk.Treeview(holiday_window, columns=columns, show="headings")
        
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=150)
        
        # Add holidays to the tree
        for holiday in holidays:
            tree.insert("", "end", values=(
                holiday.strftime("%Y-%m-%d"),
                holiday.strftime("%A")
            ))
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(holiday_window, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        
        # Pack everything
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Add delete button
        def delete_selected():
            selected = tree.selection()
            if not selected:
                return
            
            for item in selected:
                values = tree.item(item)['values']
                date = datetime.strptime(values[0], "%Y-%m-%d").date()
                self.holiday_manager.remove_holiday(date)
                tree.delete(item)
        
        ttk.Button(holiday_window, text="Delete Selected", command=delete_selected).pack(pady=10)
    
    def calculate_hours(self, in_time, out_time, selected_date):
        fmt = "%I:%M %p"
        try:
            in_dt = datetime.strptime(in_time, fmt)
            out_dt = datetime.strptime(out_time, fmt)

            if out_dt < in_dt:
                out_dt += pd.Timedelta(days=1)

            time_diff = out_dt - in_dt
            total_minutes = int(time_diff.total_seconds() // 60)

            selected_date_obj = datetime.strptime(selected_date, "%Y-%m-%d").date()
            selected_day = selected_date_obj.weekday()
            
            # Check if the date is either a weekend or a holiday
            is_weekend = selected_day in [5, 6]
            is_holiday = self.holiday_manager.is_holiday(selected_date_obj)

            if is_weekend or is_holiday:
                regular_minutes = 0
                # Round down overtime to nearest 15 minutes
                ot_minutes = (total_minutes // 15) * 15
            else:
                regular_minutes = min(480, total_minutes)
                raw_ot_minutes = max(0, total_minutes - 480)
                # Apply new overtime rule: only count overtime if the 9th hour is completed
                if raw_ot_minutes < 60:
                    ot_minutes = 0
                else:
                    # Round down overtime to nearest 15 minutes
                    ot_minutes = (raw_ot_minutes // 15) * 15

            work_hours = f"{total_minutes // 60:02d}:{total_minutes % 60:02d}"
            overtime = f"{ot_minutes // 60:02d}:{ot_minutes % 60:02d}"

            return work_hours, overtime, total_minutes, regular_minutes, ot_minutes
        except ValueError:
            messagebox.showerror("Error", "Invalid time format")
            return "00:00", "00:00", 0, 0, 0
    
    def add_entry(self):
        selected_date = self.date_picker.get_date().strftime("%Y-%m-%d")
        in_time = self.in_time_entry.get()
        out_time = self.out_time_entry.get()
        
        work_hours, overtime, total_minutes, regular_minutes, ot_minutes = self.calculate_hours(
            in_time, out_time, selected_date
        )
        
        # Update totals - convert minutes to hours
        selected_date_obj = datetime.strptime(selected_date, "%Y-%m-%d").date()
        is_weekend = selected_date_obj.weekday() in [5, 6]
        is_holiday = self.holiday_manager.is_holiday(selected_date_obj)
        
        if is_weekend or is_holiday:
            # All hours count as overtime
            self.total_overtime += total_minutes / 60.0
            self.total_hours += total_minutes / 60.0
        else:
            # Split between regular and overtime
            regular_hours = regular_minutes / 60.0
            overtime_hours = ot_minutes / 60.0
            
            self.total_regular_hours += regular_hours
            self.total_overtime += overtime_hours
            self.total_hours += (regular_hours + overtime_hours)
        
        # Format display in HH:MM
        total_hours_minutes = int(self.total_hours * 60)
        overtime_minutes = int(self.total_overtime * 60)
        
        formatted_totaltime = f"{total_hours_minutes // 60:02d}:{total_hours_minutes % 60:02d}"
        formatted_overtime = f"{overtime_minutes // 60:02d}:{overtime_minutes % 60:02d}"
        
        self.total_label.config(
            text=f"Total Work Hours: {formatted_totaltime} | Total Overtime: {formatted_overtime}"
        )
        
        self.tree.insert("", "end", values=(selected_date, in_time, out_time, work_hours, overtime))

    def delete_selected(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select an entry to delete")
            return

        for item in selected_items:
            values = self.tree.item(item)['values']
            if not values or len(values) < 3:
                messagebox.showerror("Error", "Invalid entry format")
                continue

            selected_date = values[0]
            in_time = values[1]
            out_time = values[2]

            # Recalculate hours for the entry being deleted
            work_hours, overtime, total_minutes, regular_minutes, ot_minutes = self.calculate_hours(
                in_time, out_time, selected_date
            )

            # Subtract the correct amounts based on date type
            selected_date_obj = datetime.strptime(selected_date, "%Y-%m-%d").date()
            is_weekend = selected_date_obj.weekday() in [5, 6]
            is_holiday = self.holiday_manager.is_holiday(selected_date_obj)

            if is_weekend or is_holiday:
                # All hours were overtime
                self.total_overtime = max(0, self.total_overtime - (total_minutes / 60.0))
                self.total_hours = max(0, self.total_hours - (total_minutes / 60.0))
            else:
                # Split between regular and overtime
                regular_hours = regular_minutes / 60.0
                overtime_hours = ot_minutes / 60.0
                
                self.total_regular_hours = max(0, self.total_regular_hours - regular_hours)
                self.total_overtime = max(0, self.total_overtime - overtime_hours)
                self.total_hours = max(0, self.total_hours - (regular_hours + overtime_hours))

            # Delete the entry from the treeview
            self.tree.delete(item)

        # Update display with formatted totals
        total_hours_minutes = int(self.total_hours * 60)
        overtime_minutes = int(self.total_overtime * 60)
        
        formatted_totaltime = f"{total_hours_minutes // 60:02d}:{total_hours_minutes % 60:02d}"
        formatted_overtime = f"{overtime_minutes // 60:02d}:{overtime_minutes % 60:02d}"
        
        self.total_label.config(
            text=f"Total Work Hours: {formatted_totaltime} | Total Overtime: {formatted_overtime}"
        )
    
    def clear_all(self):
        if messagebox.askyesno("Confirm Clear", "Are you sure you want to clear all entries?"):
            for item in self.tree.get_children():
                self.tree.delete(item)
            self.total_hours = 0.0
            self.total_overtime = 0.0
            self.total_regular_hours = 0.0
            self.total_label.config(
                text=f"Total Work Hours: {self.total_hours:.2f} | Total Overtime: {self.total_overtime:.2f}"
            )
    
    def print_report(self):
        report_window = tk.Toplevel(self.root)
        report_window.title("Time Card Report")
        report_window.geometry("700x500")

        report_text = ScrolledText(report_window, wrap=tk.WORD, width=80, height=20)
        report_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Report Header
        report_content = "Time Card Report\n"
        report_content += "=" * 80 + "\n"
        report_content += f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"

        # Table Headers
        headers = ["Date", "Time In", "Time Out", "Work Hours", "Overtime"]
        table_data = []

        # Fetch data from treeview
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            table_data.append(values)

        # Generate table using tabulate
        if table_data:
            report_content += tabulate(table_data, headers=headers, tablefmt="grid")
        else:
            report_content += "No records found.\n"

        # Summary Section
        report_content += f"\n\nSummary:\n"
        report_content += "=" * 80 + "\n"
        report_content += f"Total Work Hours: {self.total_hours:.2f} hours\n"
        report_content += f"Total Overtime Hours: {self.total_overtime:.2f} hours\n"

        report_text.insert(tk.END, report_content)
        report_text.configure(state='disabled')

        def save_report():
            file_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text Files", "*.txt")],
                title="Save Report"
            )
            if file_path:
                with open(file_path, 'w') as f:
                    f.write(report_content)
                messagebox.showinfo("Success", "Report saved successfully!")

        ttk.Button(report_window, text="Save Report", command=save_report).pack(pady=10)
    
    def open_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")], title="Select CSV File")
        if not file_path:
            return
        try:
            df = pd.read_csv(file_path)
            # Clear existing totals
            self.total_hours = 0.0
            self.total_overtime = 0.0
            self.total_regular_hours = 0.0
            
            # Clear existing entries
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            for _, row in df.iterrows():
                work_hours, overtime, total_minutes, regular_minutes, ot_minutes = self.calculate_hours(
                    row['In Time'], row['Out Time'], row['Date']
                )
                
                selected_date_obj = datetime.strptime(row['Date'], "%Y-%m-%d").date()
                is_weekend = selected_date_obj.weekday() in [5, 6]
                is_holiday = self.holiday_manager.is_holiday(selected_date_obj)
                
                if is_weekend or is_holiday:
                    # All hours count as overtime
                    self.total_overtime += total_minutes / 60.0
                    self.total_hours += total_minutes / 60.0
                else:
                    # Split between regular and overtime
                    regular_hours = regular_minutes / 60.0
                    overtime_hours = ot_minutes / 60.0
                    
                    self.total_regular_hours += regular_hours
                    self.total_overtime += overtime_hours
                    self.total_hours += (regular_hours + overtime_hours)
                
                self.tree.insert("", "end", values=(row['Date'], row['In Time'], row['Out Time'], work_hours, overtime))
            
            # Format display in HH:MM
            total_hours_minutes = int(self.total_hours * 60)
            overtime_minutes = int(self.total_overtime * 60)
            
            formatted_totaltime = f"{total_hours_minutes // 60:02d}:{total_hours_minutes % 60:02d}"
            formatted_overtime = f"{overtime_minutes // 60:02d}:{overtime_minutes % 60:02d}"
            
            self.total_label.config(
                text=f"Total Work Hours: {formatted_totaltime} | Total Overtime: {formatted_overtime}"
            )
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while importing: {e}")

    def save_to_csv(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            title="Save CSV"
        )
        if not file_path:
            return
        
        try:
            data = []
            for item in self.tree.get_children():
                values = self.tree.item(item)['values']
                data.append({
                    "Date": values[0],
                    "In Time": values[1],
                    "Out Time": values[2],
                    "Work Hours": values[3],
                    "Overtime": values[4]
                })
            
            df = pd.DataFrame(data)
            df.to_csv(file_path, index=False)
            messagebox.showinfo("Success", "Data saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while saving: {e}")

    def create_button_frame(self, parent):
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(0, 10))

        # Create buttons
        ttk.Button(button_frame, text="Clear All", command=self.clear_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Print Report", command=self.print_report).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Open CSV", command=self.open_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Save to CSV", command=self.save_to_csv).pack(side=tk.LEFT, padx=5)

    def create_tree_frame(self, parent):
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        # Create treeview
        columns = ("Date", "Time In", "Time Out", "Work Hours", "Overtime")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        # Configure column headings and widths
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        # Add scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)

        # Pack tree and scrollbar
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Add right-click menu
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Delete", command=self.delete_selected)

        def show_context_menu(event):
            if self.tree.selection():
                self.context_menu.post(event.x_root, event.y_root)

        self.tree.bind("<Button-3>", show_context_menu)

def main():
    root = tk.Tk()
    app = TimeCardApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
