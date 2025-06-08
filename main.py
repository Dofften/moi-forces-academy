from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib import colors
import pandas as pd
import sqlite3
from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


sheet_names = ["1-9"]

# Custom dimensions and design settings
CARD_WIDTH = 3.37 * inch
CARD_HEIGHT = 2.12 * inch
MARGIN_X = 0.5 * inch
MARGIN_Y = 0.5 * inch
SPACING_X = 0.25 * inch
SPACING_Y = 0.25 * inch

CARDS_PER_ROW = 2
CARDS_PER_COLUMN = 4


def connect_db():
    """Connect to the SQLite database and return the connection object."""
    conn = sqlite3.connect("students.db")
    return conn


def intialize_db():
    """Initialize the database and create the students table if it doesn't exist."""
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS students (
            ADMNO TEXT,
            NAME TEXT,
            GRADE TEXT,
            STREAM TEXT,
            PROCESSED INTEGER DEFAULT 0,
            PRIMARY KEY (ADMNO, NAME)
        )
    """
    )
    conn.commit()
    conn.close()


def store_new_students(new_df):
    """Store new students from the DataFrame into the database."""
    conn = connect_db()
    cursor = conn.cursor()

    for _, row in new_df.iterrows():
        cursor.execute(
            """
            INSERT OR IGNORE INTO students (ADMNO, NAME, GRADE, STREAM, PROCESSED)
            VALUES (?, ?, ?, ?, 0)
        """,
            (row["ADMNO"], row["NAME"], row["GRADE"], row["STREAM"]),
        )
    conn.commit()
    conn.close()


def get_unprocessed_students():
    """Retrieve unprocessed students from the database."""
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM students WHERE PROCESSED = 0")
    new_students = cursor.fetchall()
    conn.close()

    # Convert result to DataFrame
    columns = ["ADMNO", "NAME", "GRADE", "STREAM", "PROCESSED"]
    return pd.DataFrame(new_students, columns=columns)


def mark_students_as_processed(admnos):
    """Mark students as processed in the database."""
    conn = connect_db()
    cursor = conn.cursor()
    cursor.executemany(
        "UPDATE students SET PROCESSED = 1 WHERE ADMNO = ?",
        [(admno,) for admno in admnos],
    )
    conn.commit()
    conn.close()

def nuke_database(db_path="students.db"):
    """Delete all data from the database."""
    confirm = messagebox.askyesno("Confirm Delete", "Are you absolutely sure you want to delete ALL data?")
    if not confirm:
        return

    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        # Get all tables
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # Drop all tables
        for table_name in tables:
            cursor.execute(f"DROP TABLE IF EXISTS {table_name[0]}")

        conn.commit()
        conn.close()
        intialize_db()
        messagebox.showinfo("Database Nuked", "All tables have been deleted. The database is now empty.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to nuke database: {e}")

def draw_card(c, x, y, name, admno, grade, stream, validity):
    """Draw a single student card on the canvas."""
    # Background and Border
    # c.setFillColorRGB(0.9, 0.9, 0.9)  # Light gray background
    c.rect(x, y, CARD_WIDTH, CARD_HEIGHT, fill=0)

    # School Name Header
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(
        x + CARD_WIDTH / 2, y + CARD_HEIGHT - 0.3 * inch, "MOI FORCES ACADEMY"
    )

    # Draw the logo (optional, update with your actual logo file path)
    logo_path = "Moi_forces_academy.jpeg"  # Replace with actual logo file path
    if logo_path:
        c.drawImage(
            logo_path,
            x + CARD_WIDTH - 0.5 * inch,
            y + CARD_HEIGHT - 0.45 * inch,
            width=0.4 * inch,
            height=0.4 * inch,
        )

    # "MEAL CARD" Box
    c.setFillColor(colors.lightgrey)
    c.rect(
        x + 0.1 * inch,
        y + CARD_HEIGHT - 0.8 * inch,
        CARD_WIDTH - 0.2 * inch,
        0.35 * inch,
        fill=1,
    )
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(x + CARD_WIDTH / 2, y + CARD_HEIGHT - 0.67 * inch, "MEAL CARD")

    # Add the student details
    # text_x = x + 0.2 * inch
    text_y = y + CARD_HEIGHT - 1 * inch

    # # Name
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(x + CARD_WIDTH / 2, text_y, f"NAME: {name}")
    # Draw the name with dynamic font size or text wrapping
    # max_width = CARD_WIDTH - 0.4 * inch  # Set a max width for text
    # draw_text_with_dynamic_font_size(c, f"NAME: {name}", text_x, text_y, max_width)

    # Adm No
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(x + CARD_WIDTH / 2, text_y - 0.3 * inch, f"ADMNO: {admno}")

    # Grade and stream
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(
        x + CARD_WIDTH / 2, text_y - 0.6 * inch, f"GRADE: {grade} {stream}"
    )

    # Validity
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(
        x + CARD_WIDTH / 2, text_y - 0.9 * inch, f"VALIDITY: {validity}"
    )


def create_dataframe(input_file):
    """Create a DataFrame from the specified Excel file."""
    # Initialize an empty list to store dataframes
    df_list = []
    # Add a column to each dataframe for the sheet name
    for sheet in sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet, index_col=[0])
        df["Sheet"] = sheet  # Add the sheet name as a new column
        # Convert 'ADMNO' column to string to remove decimals
        df["ADMNO"] = df["ADMNO"].astype(str).str.replace(r"\.0$", "", regex=True)

        # Replace missing or empty ADMNO with generated unique ID
        # df["ADMNO"] = df["ADMNO"].apply(
        #     lambda x: (
        #         x if pd.notna(x) and x.strip() != "" else f"AUTO-{uuid.uuid4().hex[:6]}"
        #     )
        # )

        df_list.append(df)

    # Concatenate the dataframes
    return pd.concat(df_list, ignore_index=True)


def generate_pdf(data, output_file, validity: str):
    """Generate a PDF with student Meal cards from the provided DataFrame."""

    c = canvas.Canvas(output_file, pagesize=letter)
    PAGE_WIDTH, PAGE_HEIGHT = letter

    x = MARGIN_X
    y = PAGE_HEIGHT - MARGIN_Y - CARD_HEIGHT

    card_count = 0

    for index, row in data.iterrows():
        draw_card(
            c,
            x,
            y,
            row["NAME"],
            row["ADMNO"],
            row["GRADE"],
            row["STREAM"],
            validity=validity,
        )

        x += CARD_WIDTH + SPACING_X
        card_count += 1

        if card_count % CARDS_PER_ROW == 0:
            x = MARGIN_X
            y -= CARD_HEIGHT + SPACING_Y

        if card_count % (CARDS_PER_ROW * CARDS_PER_COLUMN) == 0:
            c.showPage()
            x = MARGIN_X
            y = PAGE_HEIGHT - MARGIN_Y - CARD_HEIGHT

    c.save()


def updated_cards(new_data_file, output_file, validity):
    """Update the database with new students and generate Meal cards."""
    # Read new data from the Excel sheet
    new_df = create_dataframe(new_data_file)

    # Store new students in the database
    store_new_students(new_df)

    # Get unprocessed students from the database
    unprocessed_students = get_unprocessed_students()

    # If there are unprocessed students, generate ID cards
    if not unprocessed_students.empty:
        # print(f"Found {len(unprocessed_students)} new students to process.")
        generate_pdf(unprocessed_students, output_file, validity=validity)
        # Mark the processed students as processed
        mark_students_as_processed(unprocessed_students["ADMNO"].tolist())
        messagebox.showinfo(
            title="Message",
            message=f"Found {len(unprocessed_students)} new students to process.\n Meal cards saved at {output_file}",
        )
    else:
        messagebox.showinfo("Message", "No new students to process.")
        # print("No new students to process.")


# GUI Functions
def select_file():
    """Open a file dialog to select an Excel file."""
    file_path = filedialog.askopenfilename(
        title="Select Excel File", filetypes=(("Excel files", "*.xlsx"),)
    )
    if file_path:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, file_path)


def on_generate():
    """Handle the "Generate Bulk Meal Cards" button click."""
    input_file = entry_file_path.get()
    validity = validity_period.get()
    output_file = f"meal_cards({datetime.now().strftime('%Y%m%dT%H%M%S')}).pdf"
    if not input_file:
        messagebox.showerror("Error", "Please select an Excel file.")
        return

    updated_cards(input_file, output_file, validity)

def sort_column(tree, col, reverse):
    """Sort the treeview by the specified column."""
    data = [(tree.set(k, col), k) for k in tree.get_children("")]
    try:
        data.sort(key=lambda t: int(t[0]) if t[0].isdigit() else t[0], reverse=reverse)
    except ValueError:
        data.sort(reverse=reverse)

    for index, (val, k) in enumerate(data):
        tree.move(k, "", index)

    tree.heading(col, command=lambda: sort_column(tree, col, not reverse))

def view_all_records():
    """View all student records in a new window."""
    def delete_selected():
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("No selection", "Please select a record to delete.")
            return

        values = tree.item(selected[0])["values"]
        admno, name = values[0], values[1]

        confirm = messagebox.askyesno(
            "Confirm Delete", f"Delete record for {name} ({admno})?"
        )
        if confirm:
            conn = connect_db()
            cursor = conn.cursor()
            cursor.execute("DELETE FROM students WHERE ADMNO = ? AND NAME = ?", (admno, name))
            conn.commit()
            conn.close()
            tree.delete(selected[0])
            messagebox.showinfo("Deleted", f"Record for {name} deleted.")

    def edit_selected():
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("No selection", "Please select a record to edit.")
            return

        values = tree.item(selected[0])["values"]
        admno, name, grade, stream, processed = values

        def save_changes():
            new_name = entry_name.get()
            new_grade = entry_grade.get()
            new_stream = entry_stream.get()

            conn = connect_db()
            cursor = conn.cursor()
            cursor.execute(
                """
                UPDATE students
                SET NAME = ?, GRADE = ?, STREAM = ?
                WHERE ADMNO = ?
                """,
                (new_name, new_grade, new_stream, admno),
            )
            conn.commit()
            conn.close()

            tree.item(selected[0], values=(admno, new_name, new_grade, new_stream, processed))
            edit_window.destroy()
            messagebox.showinfo("Updated", f"Record for {admno} updated.")

        edit_window = tk.Toplevel(view_window)
        edit_window.title(f"Edit Record: {admno}")
        edit_window.geometry("300x200")

        tk.Label(edit_window, text="Name:").pack()
        entry_name = tk.Entry(edit_window)
        entry_name.insert(0, name)
        entry_name.pack()

        tk.Label(edit_window, text="Grade:").pack()
        entry_grade = tk.Entry(edit_window)
        entry_grade.insert(0, grade)
        entry_grade.pack()

        tk.Label(edit_window, text="Stream:").pack()
        entry_stream = tk.Entry(edit_window)
        entry_stream.insert(0, stream)
        entry_stream.pack()

        tk.Button(edit_window, text="Save Changes", command=save_changes).pack(pady=10)
    
    def generate_mealcard_selected():
        """Generate meal card for the selected student."""
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("No selection", "Please select a record to generate meal card.")
            return

        values = tree.item(selected[0])["values"]
        admno, name, grade, stream, processed = values
        validity = ""

        generate_window = tk.Toplevel(view_window)
        generate_window.title(f"{name} - {admno}")
        generate_window.geometry("300x200")

        tk.Label(generate_window, text="Validity:").pack()
        entry_name = tk.Entry(generate_window)
        entry_name.insert(0, validity)
        entry_name.pack()
        def generate_pdf_and_save():
            validity = entry_name.get()
            if not validity:
                messagebox.showerror("Error", "Please enter a validity period.")
                return
            output_file = f"{name}_{admno}_{datetime.now().strftime('%Y%m%dT%H%M%S')}.pdf"
            data = pd.DataFrame({"ADMNO": [admno], "NAME": [name], "GRADE": [grade], "STREAM": [stream]})
            generate_pdf(data, output_file, validity=validity)
            messagebox.showinfo("Meal Card Generated", f"Meal card saved as {output_file}.")
        tk.Button(generate_window, text="Generate Meal Card", command=generate_pdf_and_save).pack(pady=10)

    # -- Main Window --
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("SELECT ADMNO, NAME, GRADE, STREAM, PROCESSED FROM students")
    records = cursor.fetchall()
    conn.close()

    view_window = tk.Toplevel(root)
    view_window.title("All Student Records")
    view_window.geometry("800x500")

    columns = ("ADMNO", "NAME", "GRADE", "STREAM", "PROCESSED")
    tree = ttk.Treeview(view_window, columns=columns, show="headings")
    tree.pack(fill=tk.BOTH, expand=True)

    for col in columns:
        tree.heading(col, text=col, command=lambda _col=col: sort_column(tree, _col, False))
        tree.column(col, width=120, anchor="center")

    for row in records:
        tree.insert("", tk.END, values=row)

    scrollbar = ttk.Scrollbar(view_window, orient="vertical", command=tree.yview)
    tree.configure(yscroll=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # Action Buttons
    btn_frame = tk.Frame(view_window)
    btn_frame.pack(pady=10)

    tk.Button(btn_frame, text="Edit Selected", command=edit_selected).pack(side=tk.LEFT, padx=10)
    tk.Button(btn_frame, text="Delete Selected", command=delete_selected).pack(side=tk.LEFT, padx=10)
    tk.Button(btn_frame, text="Generate Meal Card", command=generate_mealcard_selected).pack(side=tk.LEFT, padx=10)
    tk.Button(btn_frame, text="Delete All Data", fg="white", bg="red", command=nuke_database).pack(side=tk.LEFT, padx=10)

# GUI Setup
root = tk.Tk()

root.title("MFA Meal Cards Generator")

# File Selection
frame_file = tk.Frame(root)
frame_file.pack(pady=10)

label_file_path = tk.Label(frame_file, text="Excel File: ")
label_file_path.grid(row=0, column=0, padx=10, pady=10)

entry_file_path = tk.Entry(frame_file, width=40)
entry_file_path.grid(row=0, column=1, padx=10, pady=10)

label_validity_period = tk.Label(frame_file, text="Validity: ")
label_validity_period.grid(row=2, column=0, padx=10, pady=10)

validity_period = tk.Entry(frame_file, width=40)
validity_period.grid(row=2, column=1, padx=10, pady=10)

button_browse = tk.Button(frame_file, text="Browse", command=select_file)
button_browse.grid(row=0, column=2, padx=10, pady=10)

# Generate Button
button_generate = tk.Button(
    root, text="Generate Bulk Meal Cards", command=on_generate, width=20
)
button_generate.pack(pady=20)

button_view_records = tk.Button(
    root, text="View All Records", command=view_all_records, width=20
)
button_view_records.pack(pady=10)


def main():
    """Main function to run the application."""
    intialize_db()

    # Run the application
    root.mainloop()


# Usage
# intialize_db()
# input_file = input("Please input filename: ")
# validity = input("Please input validity: ")
# output_file = f"meal_cards({datetime.now().strftime('%Y%m%dT%H%M%S')}).pdf"
# updated_cards(input_file, output_file, validity)

if __name__ == "__main__":
    main()
