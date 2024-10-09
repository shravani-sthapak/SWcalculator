from tkinter import filedialog, messagebox
import pandas as pd
import tkinter as tk

class SheetMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sheet Merger Application")

        # Variables to hold number of sheets
        self.num_quiz_sheets = tk.IntVar(value=1)
        self.num_attendance_sheets = tk.IntVar(value=1)
        self.num_viva_sheets = tk.IntVar(value=1)
        self.num_demo_sheets = tk.IntVar(value=1)

        self.demo_file_path = ""
        self.quiz_file_vars = []
        self.attendance_file_vars = []
        self.viva_file_vars = []

        self.demo_grading_var = tk.StringVar(value="Normal")
        self.quiz_grading_vars = []  # Store StringVar for each quiz sheet
        self.attendance_grading_vars = []  # Store StringVar for each attendance sheet
        self.viva_grading_vars = []  # Store StringVar for each viva sheet

        # Weightage fields
        self.demo_weightage = tk.DoubleVar(value=30)
        self.quiz_weightages = []
        self.attendance_weightages = []
        self.viva_weightages = []

        # Input fields for number of sheets
        tk.Label(root, text="Number of Demo Sheets:").grid(row=0, column=0)
        tk.Entry(root, textvariable=self.num_demo_sheets).grid(row=0, column=1)

        tk.Label(root, text="Number of Quiz Sheets:").grid(row=1, column=0)
        tk.Entry(root, textvariable=self.num_quiz_sheets).grid(row=1, column=1)

        tk.Label(root, text="Number of Attendance Sheets:").grid(row=2, column=0)
        tk.Entry(root, textvariable=self.num_attendance_sheets).grid(row=2, column=1)

        tk.Label(root, text="Number of Viva Sheets:").grid(row=3, column=0)
        tk.Entry(root, textvariable=self.num_viva_sheets).grid(row=3, column=1)

        # Submit button to create upload fields
        tk.Button(root, text="Set Sheet Counts", command=self.create_upload_fields).grid(row=4, column=0, columnspan=2)

    def create_upload_fields(self):
        # Clear previous fields
        for widget in self.root.grid_slaves():
            if int(widget.grid_info()["row"]) > 4:  # Remove previous upload fields
                widget.grid_forget()

        # Create upload fields for Demo files
        for i in range(self.num_demo_sheets.get()):
            tk.Label(self.root, text=f"Upload Demo File {i + 1}:").grid(row=5 + i, column=0)
            tk.Button(self.root, text="Browse", command=lambda idx=i: self.browse_demo_file()).grid(row=5 + i, column=1)

            # Grading criteria for Demo
            criteria_options = ["Relative Grading", "Maximum Marks", "Max 75 & Min 25", "Normal"]
            tk.Label(self.root, text="Grading Method:").grid(row=5 + i, column=2)
            for j, option in enumerate(criteria_options):
                tk.Radiobutton(self.root, text=option, variable=self.demo_grading_var, value=option).grid(row=5 + i, column=j + 3)

            # Weightage for Demo
            tk.Label(self.root, text="Weightage:").grid(row=5 + i, column=6)
            tk.Entry(self.root, textvariable=self.demo_weightage).grid(row=5 + i, column=7)

        # Create upload fields for Quiz files
        for i in range(self.num_quiz_sheets.get()):
            var = tk.StringVar()
            self.quiz_file_vars.append(var)
            tk.Label(self.root, text=f"Upload Quiz File {i + 1}:").grid(row=5 + self.num_demo_sheets.get() + i, column=0)
            tk.Button(self.root, text="Browse", command=lambda v=var: self.browse_file(v)).grid(row=5 + self.num_demo_sheets.get() + i, column=1)

            # Grading criteria for Quiz
            quiz_grading_var = tk.StringVar(value="Normal")
            self.quiz_grading_vars.append(quiz_grading_var)  # Store the StringVar for each quiz
            criteria_options = ["Relative Grading", "Maximum Marks", "Max 75 & Min 25", "Normal"]
            tk.Label(self.root, text="Grading Method:").grid(row=5 + self.num_demo_sheets.get() + i, column=2)
            for j, option in enumerate(criteria_options):
                tk.Radiobutton(self.root, text=option, variable=quiz_grading_var, value=option).grid(row=5 + self.num_demo_sheets.get() + i, column=j + 3)

            # Weightage for Quiz
            quiz_weightage_var = tk.DoubleVar(value=30)
            self.quiz_weightages.append(quiz_weightage_var)
            tk.Label(self.root, text="Weightage:").grid(row=5 + self.num_demo_sheets.get() + i, column=6)
            tk.Entry(self.root, textvariable=quiz_weightage_var).grid(row=5 + self.num_demo_sheets.get() + i, column=7)

        # Create upload fields for Attendance files
        for i in range(self.num_attendance_sheets.get()):
            var = tk.StringVar()
            self.attendance_file_vars.append(var)
            tk.Label(self.root, text=f"Upload Attendance File {i + 1}:").grid(row=5 + self.num_demo_sheets.get() + self.num_quiz_sheets.get() + i, column=0)
            tk.Button(self.root, text="Browse", command=lambda v=var: self.browse_file(v)).grid(row=5 + self.num_demo_sheets.get() + self.num_quiz_sheets.get() + i, column=1)

            # Grading criteria for Attendance
            attendance_grading_var = tk.StringVar(value="Normal")
            self.attendance_grading_vars.append(attendance_grading_var)  # Store the StringVar for each attendance
            criteria_options = ["Relative Grading", "Maximum Marks", "Max 75 & Min 25", "Normal"]
            tk.Label(self.root, text="Grading Method:").grid(row=5 + self.num_demo_sheets.get() + self.num_quiz_sheets.get() + i, column=2)
            for j, option in enumerate(criteria_options):
                tk.Radiobutton(self.root, text=option, variable=attendance_grading_var, value=option).grid(row=5 + self.num_demo_sheets.get() + self.num_quiz_sheets.get() + i, column=j + 3)

            # Weightage for Attendance
            attendance_weightage_var = tk.DoubleVar(value=20)
            self.attendance_weightages.append(attendance_weightage_var)
            tk.Label(self.root, text="Weightage:").grid(row=5 + self.num_demo_sheets.get() + self.num_quiz_sheets.get() + i, column=6)
            tk.Entry(self.root, textvariable=attendance_weightage_var).grid(row=5 + self.num_demo_sheets.get() + self.num_quiz_sheets.get() + i, column=7)

        # Create upload fields for Viva files
        for i in range(self.num_viva_sheets.get()):
            var = tk.StringVar()
            self.viva_file_vars.append(var)
            tk.Label(self.root, text=f"Upload Viva File {i + 1}:").grid(row=5 + self.num_demo_sheets.get() + self.num_quiz_sheets.get() + self.num_attendance_sheets.get() + i, column=0)
            tk.Button(self.root, text="Browse", command=lambda v=var: self.browse_file(v)).grid(row=5 + self.num_demo_sheets.get() + self.num_quiz_sheets.get() + self.num_attendance_sheets.get() + i, column=1)

            # Grading criteria for Viva
            viva_grading_var = tk.StringVar(value="Normal")
            self.viva_grading_vars.append(viva_grading_var)  # Store the StringVar for each viva
            criteria_options = ["Relative Grading", "Maximum Marks", "Max 75 & Min 25", "Normal"]
            tk.Label(self.root, text="Grading Method:").grid(row=5 + self.num_demo_sheets.get() + self.num_quiz_sheets.get() + self.num_attendance_sheets.get() + i, column=2)
            for j, option in enumerate(criteria_options):
                tk.Radiobutton(self.root, text=option, variable=viva_grading_var, value=option).grid(row=5 + self.num_demo_sheets.get() + self.num_quiz_sheets.get() + self.num_attendance_sheets.get() + i, column=j + 3)

            # Weightage for Viva
            viva_weightage_var = tk.DoubleVar(value=20)
            self.viva_weightages.append(viva_weightage_var)
            tk.Label(self.root, text="Weightage:").grid(row=5 + self.num_demo_sheets.get() + self.num_quiz_sheets.get() + self.num_attendance_sheets.get() + i, column=6)
            tk.Entry(self.root, textvariable=viva_weightage_var).grid(row=5 + self.num_demo_sheets.get() + self.num_quiz_sheets.get() + self.num_attendance_sheets.get() + i, column=7)

        # Submit button to merge sheets
        tk.Button(root, text="Merge Sheets", command=self.merge_sheets).grid(row=6 + self.num_demo_sheets.get() + self.num_quiz_sheets.get() + self.num_attendance_sheets.get() + self.num_viva_sheets.get(), column=0, columnspan=2)

    def browse_demo_file(self):
        self.demo_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
   
    def browse_file(self, variable):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            variable.set(file_path)

    def merge_sheets(self):
        try:
            # Read demo file
            demo_df = pd.read_excel(self.demo_file_path)

            # Read quiz files
            quiz_dfs = []
            for quiz_file_var, quiz_weight in zip(self.quiz_file_vars, self.quiz_weightages):
                quiz_file_path = quiz_file_var.get()
                if quiz_file_path:
                    quiz_df = pd.read_excel(quiz_file_path)
                    quiz_dfs.append(quiz_df[['Enrollment_No', 'Name', 'Quiz Marks']])  # Select necessary columns

            # Read attendance files
            attendance_dfs = []
            for attendance_file_var, attendance_weight in zip(self.attendance_file_vars, self.attendance_weightages):
                attendance_file_path = attendance_file_var.get()
                if attendance_file_path:
                    attendance_df = pd.read_excel(attendance_file_path)
                    attendance_dfs.append(attendance_df[['Enrollment_No', 'Name', 'Attendance Marks']])  # Select necessary columns

            # Read viva files
            viva_dfs = []
            for viva_file_var, viva_weight in zip(self.viva_file_vars, self.viva_weightages):
                viva_file_path = viva_file_var.get()
                if viva_file_path:
                    viva_df = pd.read_excel(viva_file_path)
                    viva_dfs.append(viva_df[['Enrollment_No', 'Name', 'Viva Marks']])  # Select necessary columns

            # Merge all dataframes
            merged_df = demo_df[['Enrollment_No', 'Name', 'Demo Marks']]  # Start with demo columns
            for quiz_df in quiz_dfs:
                merged_df = merged_df.merge(quiz_df, on=['Enrollment_No', 'Name'], how='outer')
            for attendance_df in attendance_dfs:
                merged_df = merged_df.merge(attendance_df, on=['Enrollment_No', 'Name'], how='outer')
            for viva_df in viva_dfs:
                merged_df = merged_df.merge(viva_df, on=['Enrollment_No', 'Name'], how='outer')

            # Calculate weighted marks based on grading method
            self.calculate_weighted_marks(merged_df)

            # Select only required columns for final output
            final_columns = ['Enrollment_No', 'Name', 'Demo Weighted', 'Quiz Weighted', 'Attendance Weighted', 'Viva Weighted']
            merged_df['Total Marks'] = merged_df[['Demo Weighted', 'Quiz Weighted', 'Attendance Weighted', 'Viva Weighted']].sum(axis=1)  # Calculate Total Marks
            final_df = merged_df[final_columns + ['Total Marks']]  # Add Total Marks to output

            # Save merged DataFrame to Excel
            output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if output_file:
                final_df.to_excel(output_file, index=False)
                messagebox.showinfo("Success", "Sheets merged successfully!")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def calculate_weighted_marks(self, merged_df):
        # Calculate demo weighted marks
        demo_grading = self.demo_grading_var.get()
        demo_weight = self.demo_weightage.get()
        if demo_grading == "Relative Grading":
            merged_df['Demo Weighted'] = (merged_df['Demo Marks'] / merged_df['Demo Marks'].max()) * demo_weight
        elif demo_grading == "Maximum Marks":
            merged_df['Demo Weighted'] = merged_df['Demo Marks'] * (demo_weight / 100)
        elif demo_grading == "Max 75 & Min 25":
            merged_df['Demo Weighted'] = merged_df['Demo Marks'].clip(lower=25, upper=75) * (demo_weight / 100)
        else:
            merged_df['Demo Weighted'] = merged_df['Demo Marks'] * (demo_weight / 100)  # Default to normal

        # Calculate quiz weighted marks
        for quiz_df, quiz_grading_var, quiz_weight in zip(self.quiz_file_vars, self.quiz_grading_vars, self.quiz_weightages):
            quiz_grading = quiz_grading_var.get()
            if quiz_grading == "Relative Grading":
                merged_df['Quiz Weighted'] = (merged_df['Quiz Marks'] / merged_df['Quiz Marks'].max()) * quiz_weight.get()
            elif quiz_grading == "Maximum Marks":
                merged_df['Quiz Weighted'] = merged_df['Quiz Marks'] * (quiz_weight.get() / 100)
            elif quiz_grading == "Max 75 & Min 25":
                merged_df['Quiz Weighted'] = merged_df['Quiz Marks'].clip(lower=25, upper=75) * (quiz_weight.get() / 100)
            else:
                merged_df['Quiz Weighted'] = merged_df['Quiz Marks'] * (quiz_weight.get() / 100)  # Default to normal

        # Calculate attendance weighted marks
        for attendance_df, attendance_grading_var, attendance_weight in zip(self.attendance_file_vars, self.attendance_grading_vars, self.attendance_weightages):
            attendance_grading = attendance_grading_var.get()
            if attendance_grading == "Relative Grading":
                merged_df['Attendance Weighted'] = (merged_df['Attendance Marks'] / merged_df['Attendance Marks'].max()) * attendance_weight.get()
            elif attendance_grading == "Maximum Marks":
                merged_df['Attendance Weighted'] = merged_df['Attendance Marks'] * (attendance_weight.get() / 100)
            elif attendance_grading == "Max 75 & Min 25":
                merged_df['Attendance Weighted'] = merged_df['Attendance Marks'].clip(lower=25, upper=75) * (attendance_weight.get() / 100)
            else:
                merged_df['Attendance Weighted'] = merged_df['Attendance Marks'] * (attendance_weight.get() / 100)  # Default to normal

        # Calculate viva weighted marks
        for viva_df, viva_grading_var, viva_weight in zip(self.viva_file_vars, self.viva_grading_vars, self.viva_weightages):
            viva_grading = viva_grading_var.get()
            if viva_grading == "Relative Grading":
                merged_df['Viva Weighted'] = (merged_df['Viva Marks'] / merged_df['Viva Marks'].max()) * viva_weight.get()
            elif viva_grading == "Maximum Marks":
                merged_df['Viva Weighted'] = merged_df['Viva Marks'] * (viva_weight.get() / 100)
            elif viva_grading == "Max 75 & Min 25":
                merged_df['Viva Weighted'] = merged_df['Viva Marks'].clip(lower=25, upper=75) * (viva_weight.get() / 100)
            else:
                merged_df['Viva Weighted'] = merged_df['Viva Marks'] * (viva_weight.get() / 100)  # Default to normal


if __name__ == "__main__":

    root = tk.Tk()
    app = SheetMergerApp(root)
    root.mainloop()
