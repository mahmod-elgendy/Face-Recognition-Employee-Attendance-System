import tkinter as tk
import threading
from tkinter import messagebox 
from PIL import Image, ImageTk
import cv2
import face_recognition
import numpy as np
import csv
import datetime
import pandas as pd
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
import sys
from tkinter import simpledialog
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from bidi.algorithm import get_display
import arabic_reshaper
import ast
import sys
import dlib

# ✅ Determine base directory depending on whether running as a script or bundled .exe
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS  # When running as .exe via PyInstaller
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# ✅ Load all face recognition model files with proper path handling
shape_predictor_68_path = os.path.join(BASE_DIR, 'face_recognition_models', 'models', 'shape_predictor_68_face_landmarks.dat')
shape_predictor_5_path = os.path.join(BASE_DIR, 'face_recognition_models', 'models', 'shape_predictor_5_face_landmarks.dat')
face_rec_model_path = os.path.join(BASE_DIR, 'face_recognition_models', 'models', 'dlib_face_recognition_resnet_model_v1.dat')

# ✅ Initialize the models
predictor_68 = dlib.shape_predictor(shape_predictor_68_path)
predictor_5 = dlib.shape_predictor(shape_predictor_5_path)
face_rec_model = dlib.face_recognition_model_v1(face_rec_model_path)


pdfmetrics.registerFont(TTFont('Amiri', os.path.join(BASE_DIR, "fonts", "Amiri-Regular.ttf")))
def reshape_arabic_text(text):
    reshaped_text = arabic_reshaper.reshape(text)
    bidi_text = get_display(reshaped_text)
    return bidi_text

class Employee:
    def __init__(self, name, national_id, phone_number, vacations_per_month, face_encoding, arrival_time, checkout_time):
        self.name = name
        self.national_id = national_id
        self.phone_number = phone_number
        self.vacations_per_month = vacations_per_month
        self.face_encoding = face_encoding
        self.arrival_time = arrival_time
        self.checkout_time = checkout_time

    def to_list(self):
        return [self.name, self.national_id, self.phone_number, self.vacations_per_month, self.face_encoding, self.arrival_time, self.checkout_time]

class EmployeeAttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("نظام حضور الموظفين")
        self.root.geometry(f"{self.root.winfo_screenwidth()//2}x{self.root.winfo_screenheight()//2}")  # Set window size to 1/4 of the screen
        self.root.configure(bg="#007ACC")  # Light orange background
        self.unsaved_changes = False  # Track unsaved changes

        # Load and display an image at the top left
        self.logo = Image.open(os.path.join(BASE_DIR, "Panda_Logo.png"))
        self.logo = self.logo.resize((200, 200), Image.Resampling.LANCZOS)
        self.logo = ImageTk.PhotoImage(self.logo)
        tk.Label(root, image=self.logo, bg="#007ACC").place(x=10, y=10)

        tk.Label(root, text="الرئيسية", font=("Segoe UI", 16, "bold"), bg="#007ACC").pack(pady=10)
        
        buttons = [
            ("بدء النظام", self.open_start_system_window),
            ("إضافة موظف", self.open_add_employee_window),
            ("حذف موظف", self.open_delete_employee_window),
            ("تعديل بيانات الموظف", self.open_update_employee_window),
            ("خصم مباشر للموظف", self.direct_deduction),
            ("الشهر خلص؟", self.export_month),
            ("إدخال الحضور يدويًا", self.manual_attendance_entry),
            ("اغلاق", self.exit_application)
        ]

        for text, command in buttons:
            if text == "خصم مباشر للموظف":
                button_color = "red"  # Set color to red for the direct deduction button
            else:
                button_color = "#E6F0FA"
            tk.Button(root, text=text, width=20, command=command, font=("Segoe UI", 12), bg= button_color, fg="black", relief=tk.RAISED).pack(pady=5)

        # Handle closing the window via the "X" button
        self.root.protocol("WM_DELETE_WINDOW", self.exit_application)

    def open_add_employee_window(self):
        add_employee_window = tk.Toplevel(self.root)
        add_employee_window.title("اضافة موظف")
        add_employee_window.configure(bg="#007ACC")

        tk.Label(add_employee_window, text="أضف بيانات الموظف", font=("Segoe UI", 14), bg="#007ACC").pack(pady=10)

        def labeled_entry(label_text):
            tk.Label(add_employee_window, text=label_text, bg="#007ACC").pack()
            entry = tk.Entry(add_employee_window, width=30)
            entry.pack(pady=5)
            return entry

        name_entry = labeled_entry("الأسم:")
        national_id_entry = labeled_entry("الرقم القومي:")
        phone_number_entry = labeled_entry("رقم التليفون:")
        vacations_entry = labeled_entry("عدد الأجازات شهريا:")

        def time_input_row(label_text):
            tk.Label(add_employee_window, text=label_text, bg="#007ACC").pack()

            frame = tk.Frame(add_employee_window, bg="#007ACC")
            frame.pack(pady=5)

            hour_var = tk.StringVar(value="08")
            minute_var = tk.StringVar(value="00")
            period_var = tk.StringVar(value="AM")

            hours = [f"{h:02}" for h in range(1, 13)]
            minutes = [f"{m:02}" for m in range(0, 60)]
            periods = ["AM", "PM"]

            hour_menu = tk.OptionMenu(frame, hour_var, *hours)
            minute_menu = tk.OptionMenu(frame, minute_var, *minutes)
            period_menu = tk.OptionMenu(frame, period_var, *periods)

            hour_menu.pack(side=tk.LEFT)
            tk.Label(frame, text=":", bg="#007ACC").pack(side=tk.LEFT)
            minute_menu.pack(side=tk.LEFT)
            period_menu.pack(side=tk.LEFT)

            return hour_var, minute_var, period_var

        arrival_hour, arrival_minute, arrival_period = time_input_row("ميعاد الحضور:")
        checkout_hour, checkout_minute, checkout_period = time_input_row("ميعاد الأنصراف:")

        def save_employee_to_csv(employee):
            filepath = "data/Employees.csv"
            employee_data = employee.to_list()

            # Load existing data and remove empty rows (if any)
            if os.path.exists(filepath):
                df = pd.read_csv(filepath)
                df.dropna(how='all', inplace=True)  # Remove fully empty rows
            else:
                # Create a DataFrame if file doesn't exist
                df = pd.DataFrame(columns=[
                    "Name", "National ID", "Phone Number", "Vacations/Month", 
                    "Face Encoding", "Arrival Time", "Checkout Time"
                ])

            # Append the new employee
            new_row = pd.DataFrame([employee_data], columns=df.columns)
            df = pd.concat([df, new_row], ignore_index=True)

            # Save back
            df.to_csv(filepath, index=False)


        def submit_employee():
            name = name_entry.get()
            national_id = national_id_entry.get()
            phone_number = phone_number_entry.get()
            vacations_per_month = vacations_entry.get()

            arrival_time = f"{arrival_hour.get()}:{arrival_minute.get()} {arrival_period.get()}"
            checkout_time = f"{checkout_hour.get()}:{checkout_minute.get()} {checkout_period.get()}"

            if not all([name, national_id, phone_number, vacations_per_month]):
                messagebox.showerror("Input Error", "Please fill in all fields.")
                return

            try:
                # Validate formatted time
                arrival_time = datetime.datetime.strptime(arrival_time, "%I:%M %p").strftime("%I:%M %p")
                checkout_time = datetime.datetime.strptime(checkout_time, "%I:%M %p").strftime("%I:%M %p")
            except ValueError:
                messagebox.showerror("Time Format Error", "Invalid time entered. Please check the hour, minute, and AM/PM.")
                return

            # Face encoding process
            messagebox.showinfo("ترميز الوجه", "جارٍ الآن ترميز وجه هذا الموظف. يُرجى التأكد من أنه ينظر إلى الكاميرا.")

            video_capture = cv2.VideoCapture(0)
            if not video_capture.isOpened():
                messagebox.showerror("Camera Error", "لم يتم العثور على كاميرا أو تعذر فتح الكاميرا.")
                return

            video_capture.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
            video_capture.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
            ret, frame = video_capture.read()
            if not ret or frame is None:
                messagebox.showerror("Camera Error", "Could not access the camera.")
                video_capture.release()
                return

            rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            face_locations = face_recognition.face_locations(rgb_frame, model='hog')
            face_encodings = face_recognition.face_encodings(rgb_frame, face_locations)

            if face_encodings:
                face_encoding_list = face_encodings[0].tolist()
                employee = Employee(name, national_id, phone_number, int(vacations_per_month), face_encoding_list, arrival_time, checkout_time)
                save_employee_to_csv(employee)

                images_dir = os.path.join(os.getcwd(), "images")
                os.makedirs(images_dir, exist_ok=True)
                # Build the full image file path
                image_filename = f"{national_id}.png"
                image_path = os.path.join(images_dir, image_filename)

                # Save the image
                cv2.imwrite(image_path, frame)

                messagebox.showinfo("عملية ناجحة", "تم اضافة الموظف وحفظه بنجاح")
            else:
                messagebox.showwarning("No Face Detected", "جرب تاني, الوش مش ظاهر")

            video_capture.release()
            cv2.destroyAllWindows()

            # Clear form
            name_entry.delete(0, tk.END)
            national_id_entry.delete(0, tk.END)
            phone_number_entry.delete(0, tk.END)
            vacations_entry.delete(0, tk.END)

        tk.Button(add_employee_window, text="Submit", command=submit_employee, bg="#E6F0FA").pack(pady=20)




    def open_delete_employee_window(self):
        delete_employee_window = tk.Toplevel(self.root)
        delete_employee_window.title("حذف موظف")
        delete_employee_window.configure(bg="#007ACC")

        tk.Label(delete_employee_window, text="قائمة الموظفين", font=("Segoe UI", 16, "bold"), bg="#007ACC").pack(pady=10)

        # Scrollable canvas setup
        canvas = tk.Canvas(delete_employee_window, bg="#007ACC", highlightthickness=0)
        scrollbar = tk.Scrollbar(delete_employee_window, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        # Create a frame inside the canvas
        scroll_frame = tk.Frame(canvas, bg="#007ACC")

        # Create a window inside the canvas, tag it as 'inner'
        inner_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw", tags="inner")

        # Update scroll region and sync inner frame width to canvas width
        def on_canvas_configure(event):
            canvas.itemconfig("inner", width=event.width)  # Make frame match canvas width
            canvas.configure(scrollregion=canvas.bbox("all"))

        canvas.bind("<Configure>", on_canvas_configure)

        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Read employees from CSV using pandas
        try:
            employees_df = pd.read_csv("data/Employees.csv", encoding='utf-8-sig')
        except FileNotFoundError:
            messagebox.showwarning("No Data", "No employee data found.")
            return

        for _, emp in employees_df.iterrows():  # Use iterrows to loop through DataFrame rows
            name = emp['Name']
            national_id = emp['National ID']
            phone_number = emp['Phone Number']
            vacations_per_month = emp['Vacations']
            arrival_time = emp['Arrival_Time']
            checkout_time = emp['Checkout_Time']

            # Create a frame for each employee's box
            emp_box = tk.Frame(scroll_frame, bg="#E6F0FA", bd=2, relief=tk.RAISED)
            emp_box.pack(padx=10, pady=5, fill="x", expand=True)

            # Load the employee photo
            try:
                images_dir = os.path.join(os.getcwd(), "images")
                img_path = os.path.join(images_dir, f"{national_id}.png")
                img = Image.open(img_path)
                img = img.resize((100, 100), Image.Resampling.LANCZOS)
                img_tk = ImageTk.PhotoImage(img)

                tk.Label(emp_box, image=img_tk, bg="#E6F0FA").pack(pady=5)
                emp_box.image = img_tk  # Keep reference to avoid garbage collection
            except FileNotFoundError:
                tk.Label(emp_box, text="No Photo Available", bg="#E6F0FA").pack(pady=5)

            # Display employee data
            tk.Label(emp_box, text=f"Name: {name}", bg="#E6F0FA").pack(anchor="w", padx=5)
            tk.Label(emp_box, text=f"National ID: {national_id}", bg="#E6F0FA").pack(anchor="w", padx=5)
            tk.Label(emp_box, text=f"Phone Number: {phone_number}", bg="#E6F0FA").pack(anchor="w", padx=5)
            tk.Label(emp_box, text=f"Vacations per Month: {vacations_per_month}", bg="#E6F0FA").pack(anchor="w", padx=5)
            tk.Label(emp_box, text=f"Arrival Time: {arrival_time}", bg="#E6F0FA").pack(anchor="w", padx=5)
            tk.Label(emp_box, text=f"Checkout Time: {checkout_time}", bg="#E6F0FA").pack(anchor="w", padx=5)

            # Delete employee function
            def delete_employee(emp_name, emp_id, emp_box):  # Pass the box directly
                confirm = messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete {emp_name}?")
                if confirm:
                    # Remove the employee from the DataFrame
                    updated_employees_df = employees_df[employees_df['National ID'] != emp_id]

                    # Save updated list back to CSV
                    updated_employees_df.to_csv("data/Employees.csv", index=False, encoding='utf-8-sig')

                    # Remove from UI
                    emp_box.destroy()

            tk.Button(emp_box, text="حذف", command=lambda en=name, eid=national_id, eb=emp_box: delete_employee(en, eid, eb), bg="#E6F0FA").pack(pady=5)

    
    def open_start_system_window(self):
        loading_popup = tk.Toplevel(self.root)
        loading_popup.title("Starting...")
        loading_popup.geometry("300x100")
        loading_popup.configure(bg="#007ACC")

        tk.Label(loading_popup, text="الرجاء الانتظار... ادي السيستم وقته هو بيفتح اهو", font=("Segoe UI", 12),
                bg="#007ACC", fg="white", wraplength=280, justify="center").pack(expand=True)

        loading_popup.overrideredirect(True)
        loading_popup.update()

        popup_width = 300
        popup_height = 100
        screen_width = loading_popup.winfo_screenwidth()
        screen_height = loading_popup.winfo_screenheight()
        x = int((screen_width / 2) - (popup_width / 2))
        y = int((screen_height / 2) - (popup_height / 2))
        loading_popup.geometry(f"{popup_width}x{popup_height}+{x}+{y}")

        start_system_window = tk.Toplevel(self.root)
        start_system_window.title("Start The System")
        start_system_window.configure(bg="#007ACC")

        tk.Label(start_system_window, text="نظام الحضور اليومي", font=("Segoe UI", 16, "bold"), bg="#007ACC").pack(pady=10)

        # ✅ Stop button FIRST before camera starts
        stop_recognition = False

        def _stop_recognition_loop():
            nonlocal stop_recognition
            stop_recognition = True
            start_system_window.destroy()

        stop_button = tk.Button(start_system_window, text="إغلاق النظام", font=("Segoe UI", 12), bg="#E6F0FA", command=_stop_recognition_loop)
        stop_button.pack(pady=10)

        # ✅ Load employee data
        employees = []
        employee_info = {}

        try:
            with open("data/Employees.csv", mode='r', encoding='utf-8-sig') as file:
                reader = csv.reader(file)
                next(reader)
                for row in reader:
                    name, national_id, phone_number, vacations_per_month, face_encoding, arrival_time, checkout_time = row
                    employees.append(row)
                    employee_info[national_id] = {
                        "name": name,
                        "scheduled_arrival": arrival_time,
                        "scheduled_checkout": checkout_time
                    }
        except FileNotFoundError:
            messagebox.showwarning("No Data", "No employee data found.")
            return

        known_encodings = []
        known_names = []

        for emp in employees:
            name, national_id, phone_number, vacations_per_month, face_encoding, arrival_time, checkout_time = emp
            face_encoding = str(face_encoding).strip()
            if not face_encoding or face_encoding.lower() in ["nan", "none", "null", ""]:
                continue
            try:
                encoding_list = ast.literal_eval(face_encoding)
                if isinstance(encoding_list, list):
                    known_encodings.append(np.array(encoding_list))
                    known_names.append(name)
                else:
                    print(f"Skipping {name}: Decoded encoding is not a list")
            except Exception as e:
                print(f"Error decoding face encoding for {name}: {e}")

        video_capture = cv2.VideoCapture(0)  # Use Video for Windows for maximum compatibility
        failed_reads = 0 
        while not video_capture.isOpened():
            failed_reads += 1
            if failed_reads > 30:
                messagebox.showerror("Camera Error", "تعذر الحصول على صورة من الكاميرا.")
                break
            return

        video_capture.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
        video_capture.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)

        loading_popup.destroy()

        try:
            df = pd.read_csv("data/Monthly_Data.csv", encoding='utf-8-sig')
        except FileNotFoundError:
            df = pd.DataFrame(columns=["Date", "Name", "National ID", "Arrival Time", "Check Out Time", "Deduction", "Deduction Message", "dState"])

        while not stop_recognition:
            root.update()  

            ret, frame = video_capture.read()
            rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            face_locations = face_recognition.face_locations(rgb_frame, model='hog')
            face_encodings = face_recognition.face_encodings(rgb_frame, face_locations)

            for face_encoding, face_location in zip(face_encodings, face_locations):
                name = "Unknown"
                national_id = ""

                if known_encodings:  # Ensure list is not empty
                    face_distances = face_recognition.face_distance(known_encodings, face_encoding)
                    best_match_index = np.argmin(face_distances)

                    if face_distances[best_match_index] < 0.5:  # ✅ You can adjust this threshold
                        name = known_names[best_match_index]
                        national_id = employees[best_match_index][1]

                        current_time = datetime.datetime.now()
                        current_date = current_time.strftime("%Y-%m-%d")
                        current_time_str = current_time.strftime("%I:%M %p")

                        scheduled_arrival_str = employee_info[national_id]["scheduled_arrival"]
                        scheduled_checkout_str = employee_info[national_id]["scheduled_checkout"]

                        scheduled_arrival = datetime.datetime.strptime(scheduled_arrival_str, "%I:%M %p").time()
                        scheduled_checkout = datetime.datetime.strptime(scheduled_checkout_str, "%I:%M %p").time()

                        arrival_start = (datetime.datetime.combine(datetime.date.today(), scheduled_arrival) - datetime.timedelta(hours=3)).time()
                        arrival_end = (datetime.datetime.combine(datetime.date.today(), scheduled_arrival) + datetime.timedelta(hours=3)).time()
                        checkout_start = (datetime.datetime.combine(datetime.date.today(), scheduled_checkout) - datetime.timedelta(hours=3)).time()
                        checkout_end = (datetime.datetime.combine(datetime.date.today(), scheduled_checkout) + datetime.timedelta(hours=3)).time()

                        is_arrival = arrival_start <= current_time.time() <= arrival_end
                        is_checkout = checkout_start <= current_time.time() <= checkout_end

                        lateness = (datetime.datetime.combine(datetime.date.today(), current_time.time()) -
                                    datetime.datetime.combine(datetime.date.today(), scheduled_arrival)).total_seconds() / 60

                        deduction = ""
                        deduction_message = ""
                        dState = ""

                        if is_arrival and lateness >= 17:
                            deduction = "-20 EGP"
                            deduction_message = "وصل متأخر ب 15 دقيقة او اكتر"

                        existing_entry = df[(df["Date"] == current_date) & (df["National ID"] == national_id)]

                        if existing_entry.empty:
                            new_entry = pd.DataFrame([[current_date, name, national_id, current_time_str if is_arrival else "",
                                                    current_time_str if is_checkout else "", deduction, deduction_message, dState]],
                                                    columns=df.columns)
                            df = pd.concat([df, new_entry], ignore_index=True)
                        else:
                            if is_arrival:
                                df.loc[(df["Date"] == current_date) & (df["National ID"] == national_id), "Arrival Time"] = current_time_str
                                df.loc[(df["Date"] == current_date) & (df["National ID"] == national_id), "Deduction"] = deduction
                                df.loc[(df["Date"] == current_date) & (df["National ID"] == national_id), "Deduction Message"] = deduction_message
                            if is_checkout:
                                df.loc[(df["Date"] == current_date) & (df["National ID"] == national_id), "Check Out Time"] = current_time_str

                    top, right, bottom, left = face_location
                    cv2.rectangle(frame, (left, top), (right, bottom), (0, 255, 0), 2)
                    cv2.putText(frame, name, (left, top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 255, 0), 2)

            cv2.imshow('Face Recognition', frame)

            if cv2.waitKey(1) & 0xFF == 27:
                break

        video_capture.release()
        cv2.destroyAllWindows()

        df.dropna(how='all', inplace=True)
        df.to_csv("data/Monthly_Data.csv", index=False, encoding='utf-8-sig')


    def manual_attendance_entry(self):
        password = simpledialog.askstring("كلمة المرور", "من فضلك أدخل كلمة المرور:", show='*')
        if password != "2255":
            messagebox.showerror("كلمة المرور خاطئة", "كلمة المرور غير صحيحة، لا يمكنك الدخول.")
            return
        
        window = tk.Toplevel(self.root)
        window.title("إدخال الحضور يدويًا")
        window.configure(bg="#007ACC")

        tk.Label(window, text="إدخال الحضور يدويًا", font=("Segoe UI", 16, "bold"), bg="#007ACC", fg="white").pack(pady=10)

        form_frame = tk.Frame(window, bg="#007ACC")
        form_frame.pack(pady=10)

        # National ID
        tk.Label(form_frame, text="الرقم القومي:", bg="#007ACC", fg="white").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        national_id_entry = tk.Entry(form_frame, width=30)
        national_id_entry.grid(row=0, column=1, padx=10)

        # Date
        tk.Label(form_frame, text="التاريخ (سنة/يوم/شهر):", bg="#007ACC", fg="white").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        date_entry = tk.Entry(form_frame, width=30)
        date_entry.insert(0, datetime.datetime.now().strftime("%m/%d/%Y"))
        date_entry.grid(row=1, column=1, padx=10)

        # Arrival Time
        tk.Label(form_frame, text="وقت الحضور :", bg="#007ACC", fg="white").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        arrival_entry = tk.Entry(form_frame, width=30)
        arrival_entry.grid(row=2, column=1, padx=10)

        # Checkout Time
        tk.Label(form_frame, text="وقت الانصراف :", bg="#007ACC", fg="white").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        checkout_entry = tk.Entry(form_frame, width=30)
        checkout_entry.grid(row=3, column=1, padx=10)

        def submit_manual_attendance():
            national_id = national_id_entry.get().strip()
            date = date_entry.get().strip()
            arrival = arrival_entry.get().strip()
            checkout = checkout_entry.get().strip()

            if not (national_id and date and arrival and checkout):
                messagebox.showwarning("تنبيه", "يرجى إدخال جميع البيانات.")
                return

            # Try to get the employee's name from Employees.csv
            name = ""
            try:
                df = pd.read_csv("data/Employees.csv")
                df['National ID'] = df['National ID'].astype(str).str.strip()
                match = df[df['National ID'] == national_id.strip()]

                if not match.empty:
                    name = match.iloc[0]['Name']
                else:
                    messagebox.showwarning("لم يتم العثور", "الموظف غير موجود في قائمة الموظفين.")
                    return
            except FileNotFoundError:
                messagebox.showwarning("خطأ", "ملف الموظفين غير موجود.")
                return

            new_row = {
                "Date": date,
                "Name": name,
                "National ID": national_id,
                "Arrival Time": arrival,
                "Check Out Time": checkout,
                "Deduction": "",
                "Deduction Message": "",
                "Deduction State": ""
            }

            with open("data/Monthly_Data.csv", mode='a', newline='', encoding='utf-8-sig') as file:
                writer = csv.DictWriter(file, fieldnames=["Date", "Name", "National ID", "Arrival Time", "Check Out Time", "Deduction", "Deduction Message", "Deduction State"])
                if file.tell() == 0:
                    writer.writeheader()
                writer.writerow(new_row)

            messagebox.showinfo("تم الحفظ", f"تم حفظ الحضور لـ {name}")
            national_id_entry.delete(0, tk.END)
            arrival_entry.delete(0, tk.END)
            checkout_entry.delete(0, tk.END)

        tk.Button(window, text="حفظ الحضور", command=submit_manual_attendance, bg="#004E8C", fg="white").pack(pady=10)



    def direct_deduction(self):

        password = simpledialog.askstring("كلمة المرور", "من فضلك أدخل كلمة المرور:", show='*')
        if password != "2255":
            messagebox.showerror("كلمة المرور خاطئة", "كلمة المرور غير صحيحة، لا يمكنك الدخول.")
            return
        
        deduction_window = tk.Toplevel(self.root)
        deduction_window.title("الخصومات المباشرة")
        deduction_window.configure(bg="#007ACC")

        tk.Label(deduction_window, text="إدخال خصومات للموظفين", font=("Segoe UI", 16, "bold"), bg="#007ACC", fg="white").pack(pady=10)

        # Read employees from CSV using pandas
        try:
            employees_df = pd.read_csv("data/Employees.csv", encoding='utf-8-sig')
        except FileNotFoundError:
            messagebox.showwarning("لا يوجد بيانات", "لم يتم العثور على بيانات الموظفين.")
            return

        # Scrollable canvas setup
        canvas = tk.Canvas(deduction_window, bg="#007ACC")
        scrollbar = tk.Scrollbar(deduction_window, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg="#007ACC")

        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig("inner", width=canvas.winfo_width())  # Sync inner frame width

        canvas.bind("<Configure>", on_frame_configure)
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw", tags="inner")

        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for _, emp in employees_df.iterrows():  # Use iterrows to loop through DataFrame rows
            name = emp['Name']
            national_id = emp['National ID']

            emp_frame = tk.Frame(scroll_frame, bg="#E6F0FA", bd=2, relief=tk.RIDGE)
            emp_frame.pack(padx=10, pady=10, fill=tk.X)

            # Load the employee photo
            try:
                images_dir = os.path.join(os.getcwd(), "images")
                img_path = os.path.join(images_dir, f"{national_id}.png")
                img = Image.open(img_path)
                img = img.resize((100, 100), Image.Resampling.LANCZOS)
                img_tk = ImageTk.PhotoImage(img)

                tk.Label(emp_frame, image=img_tk, bg="#E6F0FA").pack(pady=5)
                emp_frame.image = img_tk  # Keep reference to avoid garbage collection
            except FileNotFoundError:
                tk.Label(emp_frame, text="لا توجد صورة متاحة", bg="#E6F0FA").pack(pady=5)

            tk.Label(emp_frame, text=f"الاسم: {name}", bg="#E6F0FA", anchor="w").pack(fill=tk.X, padx=10)
            tk.Label(emp_frame, text=f"الرقم القومي: {national_id}", bg="#E6F0FA", anchor="w").pack(fill=tk.X, padx=10)

            # Deduction entry
            tk.Label(emp_frame, text="الخصم (مثال: -20 EGP):", bg="#E6F0FA").pack(anchor="w", padx=10)
            deduction_entry = tk.Entry(emp_frame, width=40)
            deduction_entry.pack(padx=10, pady=5)

            # Deduction message entry
            tk.Label(emp_frame, text="رسالة الخصم:", bg="#E6F0FA").pack(anchor="w", padx=10)
            message_entry = tk.Entry(emp_frame, width=60)
            message_entry.pack(padx=10, pady=5)

            def submit_deduction(emp_data=emp, ded_entry=deduction_entry, msg_entry=message_entry):
                deduction = ded_entry.get().strip()
                message = msg_entry.get().strip()
                date_today = datetime.date.today().isoformat()

                if not deduction:
                    messagebox.showwarning("تحذير", "يرجى إدخال قيمة الخصم.")
                    return
                
                dState = "Checked"
                new_row = {
                    "Date": date_today,
                    "Name": emp_data['Name'],
                    "National ID": emp_data['National ID'],
                    "Arrival Time": "",  # Placeholder
                    "Check Out Time": "",  # Placeholder
                    "Deduction": deduction,
                    "Deduction Message": message,
                    "Deduction State": dState
                }

                # Append the new row to the CSV
                with open("data/Monthly_Data.csv", mode='a', newline='', encoding='utf-8-sig') as file:
                    writer = csv.DictWriter(file, fieldnames=["Date", "Name", "National ID", "Arrival Time", "Check Out Time", "Deduction", "Deduction Message", "Deduction State"])
                    
                    # Check if the file is empty to write the header
                    file_empty = file.tell() == 0
                    if file_empty:
                        writer.writeheader()  # Write the header if the file is empty
                    
                    writer.writerow(new_row)  # Write the new row

                messagebox.showinfo("تم الحفظ", f"تم حفظ الخصم لـ {emp_data['Name']}")
                ded_entry.delete(0, tk.END)
                msg_entry.delete(0, tk.END)

            tk.Button(emp_frame, text="حفظ الخصم", command=submit_deduction, bg="#007ACC", fg="white").pack(pady=5)



    def open_update_employee_window(self):
        update_employee_window = tk.Toplevel(self.root)
        update_employee_window.title("تعديل بيانات موظف")
        update_employee_window.configure(bg="#007ACC")

        tk.Label(update_employee_window, text="قائمة الموظفين", font=("Segoe UI", 16, "bold"), bg="#007ACC").pack(pady=10)

        # Scrollable frame setup
        canvas = tk.Canvas(update_employee_window, bg="#007ACC", highlightthickness=0)
        scrollbar = tk.Scrollbar(update_employee_window, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#007ACC")

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Load employees
        try:
            with open("data/Employees.csv", mode='r', encoding='utf-8') as file:
                reader = csv.reader(file)
                employees = list(reader)
        except FileNotFoundError:
            messagebox.showwarning("No Data", "No employee data found.")
            return

        if not employees:
            return

        header = employees[0]
        employees_data = employees[1:]

        for idx, emp in enumerate(employees_data):
            if len(emp) != 7:
                continue  # Skip bad rows

            name, national_id, phone_number, vacations_per_month, face_encoding, arrival_time, checkout_time = emp

            emp_box = tk.Frame(scrollable_frame, bg="#E6F0FA", bd=2, relief=tk.RAISED)
            emp_box.pack(side=tk.TOP, fill=tk.BOTH, padx=10, pady=10)

            # Employee image
            try:
                img_path = os.path.join("images", f"{national_id}.png")
                img = Image.open(img_path).resize((100, 100), Image.Resampling.LANCZOS)
                img_tk = ImageTk.PhotoImage(img)
                tk.Label(emp_box, image=img_tk, bg="#E6F0FA").pack(pady=5)
                emp_box.image = img_tk  # Keep reference
            except:
                tk.Label(emp_box, text="No Photo Available", bg="#E6F0FA").pack(pady=5)

            # Employee info
            tk.Label(emp_box, text=f"Name: {name}", bg="#E6F0FA").pack(anchor="w", padx=5)
            tk.Label(emp_box, text=f"National ID: {national_id}", bg="#E6F0FA").pack(anchor="w", padx=5)

            # Editable fields
            tk.Label(emp_box, text="New Phone Number:", bg="#E6F0FA").pack(anchor="w", padx=5)
            phone_entry = tk.Entry(emp_box, width=30)
            phone_entry.insert(0, phone_number)
            phone_entry.pack()

            tk.Label(emp_box, text="New Vacations per Month:", bg="#E6F0FA").pack(anchor="w", padx=5)
            vac_entry = tk.Entry(emp_box, width=30)
            vac_entry.insert(0, vacations_per_month)
            vac_entry.pack()

            def parse_time_string(t):
                try:
                    time_part, ampm = t.strip().split(" ")
                    hh, mm = time_part.split(":")
                    return hh.zfill(2), mm.zfill(2), ampm
                except:
                    return "08", "00", "AM"

            def create_time_selector(parent, label_text, default_time):
                hh, mm, ampm = parse_time_string(default_time)

                frame = tk.Frame(parent, bg="#E6F0FA")
                frame.pack(pady=5)

                tk.Label(frame, text=label_text, bg="#E6F0FA").pack(side=tk.LEFT, padx=5)

                hh_var = tk.StringVar(value=hh)
                mm_var = tk.StringVar(value=mm)
                ampm_var = tk.StringVar(value=ampm)

                hours = [str(h).zfill(2) for h in range(1, 13)]
                minutes = [str(m).zfill(2) for m in range(0, 60, 5)]
                ampm_options = ["AM", "PM"]

                tk.OptionMenu(frame, hh_var, *hours).pack(side=tk.LEFT)
                tk.Label(frame, text=":", bg="#E6F0FA").pack(side=tk.LEFT)
                tk.OptionMenu(frame, mm_var, *minutes).pack(side=tk.LEFT)
                tk.OptionMenu(frame, ampm_var, *ampm_options).pack(side=tk.LEFT)

                return hh_var, mm_var, ampm_var

            arrival_hh, arrival_mm, arrival_ampm = create_time_selector(emp_box, "New Arrival Time:", arrival_time)
            checkout_hh, checkout_mm, checkout_ampm = create_time_selector(emp_box, "New Checkout Time:", checkout_time)

            def make_update_callback(national_id, phone_entry, vac_entry,
                                    arrival_hh, arrival_mm, arrival_ampm,
                                    checkout_hh, checkout_mm, checkout_ampm):
                def update():
                    updated_phone = phone_entry.get()
                    updated_vac = vac_entry.get()
                    updated_arrival = f"{arrival_hh.get()}:{arrival_mm.get()} {arrival_ampm.get()}"
                    updated_checkout = f"{checkout_hh.get()}:{checkout_mm.get()} {checkout_ampm.get()}"

                    # Find the employee by National ID and update the row
                    employee_found = False
                    for emp in employees_data:
                        if emp[1] == national_id:
                            emp[2] = updated_phone
                            emp[3] = updated_vac
                            emp[5] = updated_arrival
                            emp[6] = updated_checkout
                            employee_found = True
                            break

                    if not employee_found:
                        messagebox.showwarning("Update Failed", f"No employee found with National ID {national_id}.")
                        return

                    with open("data/Employees.csv", mode='w', newline='', encoding='utf-8') as file:
                        writer = csv.writer(file)
                        writer.writerow(header)
                        writer.writerows(employees_data)

                    messagebox.showinfo("Success", f"Information for National ID {national_id} updated successfully.")
                return update


            tk.Button(
                emp_box,
                text="تعديل وحفظ",
                command=make_update_callback(
                    national_id,
                    phone_entry,
                    vac_entry,
                    arrival_hh, arrival_mm, arrival_ampm,
                    checkout_hh, checkout_mm, checkout_ampm
                ),
                bg="#E6F0FA"
            ).pack(pady=5)


    def check_absent_employees_and_update_vacations(self):
        emp_path = "data/Employees.csv"
        att_path = "data/Monthly_Data.csv"

        # Read employees data (now with header)
        try:
            employees_df = pd.read_csv(emp_path, encoding='utf-8-sig')
        except FileNotFoundError:
            messagebox.showerror("خطأ", "لم يتم العثور على ملف الموظفين.")
            return

        today = datetime.datetime.now().date()
        today_str = today.strftime("%m/%d/%Y")
        new_entries = []

        for _, row in employees_df.iterrows():
            national_id = str(row['National ID']).strip()
            name = row['Name']
            try:
                vacations = int(row['Vacations per Month'])
            except (ValueError, KeyError):
                vacations = 0

            try:
                attendance_df = pd.read_csv(att_path, encoding='utf-8-sig')
                attendance_df['Date'] = pd.to_datetime(attendance_df['Date'], errors='coerce').dt.date
                has_arrival = not attendance_df[
                    (attendance_df['Date'] == today) &
                    (attendance_df['National ID'].astype(str).str.strip() == national_id)
                ].empty
            except FileNotFoundError:
                has_arrival = False

            # Only apply absence if no arrival and today > attendance date
            attendance_df['Date'] = pd.to_datetime(attendance_df['Date'], errors='coerce').dt.date
            if not has_arrival and today > attendance_df['Date'].max():
                if vacations > 0:
                    employees_df.loc[employees_df['National ID'].astype(str).str.strip() == national_id, 'Vacations per Month'] = vacations - 1
                else:
                    new_row = {
                        'Date': today_str,
                        'Name': name,
                        'National ID': national_id,
                        'Arrival Time': '',
                        'Check Out Time': '',
                        'Deduction': 'خصم يومين',
                        'Deduction Message': 'مجاش ورصيد اجازاته 0',
                        "Deduction State": ''
                    }
                    new_entries.append(new_row)


        if new_entries:
            with open(att_path, mode='a', newline='', encoding='utf-8-sig') as file:
                writer = csv.DictWriter(file, fieldnames=[
                    "Date", "Name", "National ID", "Arrival Time", "Check Out Time", "Deduction", "Deduction Message", "Deduction State"
                ])
                for entry in new_entries:
                    writer.writerow(entry)

        # Save updated employees CSV with header and proper encoding
        employees_df.to_csv(emp_path, index=False, encoding='utf-8-sig')

        # Show success message
        messagebox.showinfo("تم", "تم التحقق من الغياب وتحديث الإجازات بنجاح.")


    def validation(self):
        # Load data
        filepath = "data/Monthly_Data.csv"
        df = pd.read_csv(filepath, encoding='utf-8-sig')

        # Ensure 'Date' column is in datetime format
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.date

        # Initialize a list to store validated rows
        validated_rows = []

        # Group by Date and National ID
        grouped = df.groupby(['Date', 'National ID'])

        for (date, national_id), group in grouped:
            # Get first arrival row (non-null 'Arrival Time')
            arrival_row = group[group['Arrival Time'].notna()].sort_values('Arrival Time').head(1)

            # Get last checkout row (non-null 'Check Out Time')
            checkout_row = group[group['Check Out Time'].notna()].sort_values('Check Out Time').tail(1)

            # If arrival and checkout are the same row, only add once
            if not arrival_row.empty and not checkout_row.empty and arrival_row.index[0] == checkout_row.index[0]:
                validated_rows.append(arrival_row.iloc[0].to_dict())
            else:
                if not arrival_row.empty:
                    validated_rows.append(arrival_row.iloc[0].to_dict())
                if not checkout_row.empty:
                    validated_rows.append(checkout_row.iloc[0].to_dict())

        # After processing all groups, include rows with 'Deduction State' as 'Checked'
        checked_rows = df[df['Deduction State'] == "Checked"]
        validated_rows.extend(checked_rows.to_dict(orient='records'))  # Add all checked rows

        # Build cleaned DataFrame
        cleaned_df = pd.DataFrame(validated_rows)

        # Reset index
        cleaned_df.reset_index(drop=True, inplace=True)

        # Save back or return
        cleaned_df.to_csv(filepath, index=False, encoding='utf-8-sig')


    def show_exit_prompt(self):
        exit_prompt = tk.Toplevel(self.root)
        exit_prompt.title("Exit Confirmation")
        exit_prompt.geometry("300x150")
        exit_prompt.configure(bg="#007ACC")

        tk.Label(exit_prompt, text="متأكد انك عايز تقفل؟", bg="#007ACC").pack(pady=10)

        button_frame = tk.Frame(exit_prompt, bg="#007ACC")
        button_frame.pack(pady=10)

        tk.Button(button_frame, text="ايوه", command=lambda: [exit_prompt.destroy(), self.root.destroy()],
                width=10, bg="#E6F0FA").pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="لا استني", command=exit_prompt.destroy,
                width=10, bg="#E6F0FA").pack(side=tk.RIGHT, padx=5)


    def exit_application(self):
        
        att_path = "data/Monthly_Data.csv"

        # ✅ Step 1: Check if file exists
        if not os.path.exists(att_path):
            self.show_exit_prompt()
            return

        try:
            # ✅ Step 2: Load the CSV
            df = pd.read_csv(att_path)

            # ✅ Step 3: Drop completely empty rows (all NaNs)
            df.dropna(how='all', inplace=True)

            # ✅ Step 4: Check if any valid data remains
            if df.empty:
                self.show_exit_prompt()
                return

        except Exception as e:
            print(f"Error reading CSV: {e}")
            self.show_exit_prompt()
            return

        # Set the current date in MM/DD/YYYY format
        self.validation()
        self.check_absent_employees_and_update_vacations()
        current_date = datetime.datetime.now().strftime("%m/%d/%Y")

        try:
            # Load existing attendance data
            df = pd.read_csv("data/Monthly_Data.csv")
        except FileNotFoundError:
            df = pd.DataFrame(columns=["Date", "Name", "National ID", "Arrival Time", "Check Out Time", "Deduction", "Deduction Message"])

        # Filter records for the current date
        today_records = df[df["Date"] == current_date]

        # Check for missing checkout times
        incomplete_records = today_records[today_records["Arrival Time"].notna() & today_records["Check Out Time"].isna()]

        # Define the time range to check for late checkouts
        end_of_day = datetime.datetime.strptime(current_date + " 23:59:59", "%m/%d/%Y %H:%M:%S")
        late_check_end = end_of_day + datetime.timedelta(hours=4)  # 4 hours after the day ends

        for index, row in today_records.iterrows():
            # Check if Arrival Time is a valid string before processing
            if isinstance(row["Arrival Time"], str):
                arrival_time = datetime.datetime.strptime(row["Arrival Time"], "%I:%M %p")
            else:
                continue  # Skip if not a string

            checkout_time = row["Check Out Time"]

            # If checkout_time is NaN, it means they haven't checked out
            if pd.isna(checkout_time):
                # Check if current time is past the late checkout time
                if datetime.datetime.now() > late_check_end:
                    incomplete_records = incomplete_records.append(row)

        # If there are incomplete records, show a warning message
        if not incomplete_records.empty:
            messagebox.showwarning("Incomplete Records", "Some employees have recorded arrival times but not checkout times.")

        # Proceed with the exit confirmation
        exit_prompt = tk.Toplevel(self.root)
        exit_prompt.title("Exit Confirmation")
        exit_prompt.geometry("300x150")
        exit_prompt.configure(bg="#007ACC")
        
        tk.Label(exit_prompt, text="متأكد عايز تقفل؟", bg="#007ACC").pack(pady=10)
        
        button_frame = tk.Frame(exit_prompt, bg="#007ACC")
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="أيوه", command=lambda: [exit_prompt.destroy(), self.root.destroy()], width=10, bg="#E6F0FA").pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="لأ استني", command=exit_prompt.destroy, width=10, bg="#E6F0FA").pack(side=tk.RIGHT, padx=5)

    
    def export_month(self):
        password = simpledialog.askstring("كلمة المرور", "من فضلك أدخل كلمة المرور:", show='*')
        if password != "2255":
            messagebox.showerror("كلمة المرور خاطئة", "كلمة المرور غير صحيحة، لا يمكنك الدخول.")
            return

        current_date = datetime.datetime.now()
        month = current_date.strftime("%B")
        year = current_date.year

        try:
            with open("data/Employees.csv", mode='r') as file:
                reader = csv.reader(file)
                next(reader)
                employees = list(reader)

            df = pd.read_csv("data/Monthly_Data.csv", encoding='utf-8-sig')

            for emp in employees:
                name, national_id, phone_number, vacations_per_month, face_encoding, arrival_time, checkout_time = emp
                employee_records = df[df["National ID"] == int(national_id)]

                reports_dir = os.path.join(os.getcwd(), "reports")
                os.makedirs(reports_dir, exist_ok=True)

                pdf_filename = f"{national_id}.pdf"
                pdf_path = os.path.join(reports_dir, pdf_filename)
                c = canvas.Canvas(pdf_path, pagesize=letter)
                width, height = letter

                def draw_header():
                    if os.path.exists(f"images/{national_id}.png"):
                        c.drawImage(f"images/{national_id}.png", 1 * inch, height - 1.8 * inch, width=1 * inch, height=1.3 * inch)

                    c.setFont("Amiri", 10)
                    
                    # Reshaping the Arabic text before rendering
                    reshaped_name = reshape_arabic_text(f"الاسم: {name}")
                    reshaped_national_id = reshape_arabic_text(f"الرقم القومي: {national_id}")
                    reshaped_phone = reshape_arabic_text(f"رقم الهاتف: {phone_number}")
                    reshaped_monthly_record = reshape_arabic_text("السجل الشهري:")

                    c.drawRightString(7.5 * inch, height - 1.0 * inch, reshape_arabic_text(f"الاسم: {name}"))
                    c.drawRightString(7.5 * inch, height - 1.4 * inch, reshape_arabic_text(f"الرقم القومي: {national_id}"))
                    c.drawRightString(7.5 * inch, height - 1.8 * inch, reshape_arabic_text(f"رقم الهاتف: {phone_number}"))
                    c.drawRightString(7.5 * inch, height - 2.2 * inch, reshape_arabic_text("السجل الشهري:"))


                draw_header()

                y_position = height - 2.6 * inch
                c.setFont("Amiri", 8)

                for index, row in employee_records.iterrows():
                    date = row["Date"]
                    arrival = row["Arrival Time"] if pd.notna(row["Arrival Time"]) else "-"
                    checkout = row["Check Out Time"] if pd.notna(row["Check Out Time"]) else "-"
                    deduction = row["Deduction"] if pd.notna(row["Deduction"]) else "-"
                    message = row["Deduction Message"] if pd.notna(row["Deduction Message"]) else ""

                    # Reshaping each Arabic line
                    line = f"{date} | دخول: {arrival} | خروج: {checkout} | خصم: {deduction} | سبب: {message}"
                    reshaped_line = reshape_arabic_text(line)
                    c.drawRightString(7.5 * inch, y_position, reshape_arabic_text(line))
                    y_position -= 0.2 * inch

                    if y_position < 1 * inch:
                        c.showPage()
                        draw_header()
                        y_position = height - 2.6 * inch
                        c.setFont("Amiri", 8)

                c.setFont("Amiri", 10)
                c.drawRightString(7.5 * inch, y_position - 0.4 * inch, reshape_arabic_text("الراتب الكلي: ____________________"))
                c.drawRightString(7.5 * inch, y_position - 0.8 * inch, reshape_arabic_text("توقيع الموظف: ____________________"))


                c.save()

            messagebox.showinfo("تم التصدير", "تم إنشاء تقارير الشهر بنجاح.")

        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = EmployeeAttendanceApp(root)
    root.mainloop()