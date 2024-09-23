from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib import colors
import pandas as pd
import sqlite3
from datetime import datetime


sheet_names = ['1-6','7E','7S','7A','7K','7L','7M','8E','8S','8A','8L','8M']

# Custom dimensions and design settings
CARD_WIDTH = 3.37 * inch
CARD_HEIGHT = 2.12 * inch
MARGIN_X = 0.5 * inch
MARGIN_Y = 0.5 * inch
SPACING_X = 0.25 * inch
SPACING_Y = 0.25 * inch

CARDS_PER_ROW = 2
CARDS_PER_COLUMN = 4

# Connect to SQLite database
def connect_db():
    conn = sqlite3.connect('students.db')
    return conn

def create_students_table():
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS students (
            ADMNO TEXT PRIMARY KEY,
            NAME TEXT,
            GRADE TEXT,
            STREAM TEXT,
            PROCESSED INTEGER DEFAULT 0
        )
    ''')
    conn.commit()
    conn.close()

def store_new_students(new_df):
    conn = connect_db()
    cursor = conn.cursor()

    for _, row in new_df.iterrows():
        cursor.execute('''
            INSERT OR IGNORE INTO students (ADMNO, NAME, GRADE, STREAM, PROCESSED)
            VALUES (?, ?, ?, ?, 0)
        ''', (row['ADMNO'], row['NAME'], row['GRADE'], row['STREAM']))
    conn.commit()
    conn.close()

def get_unprocessed_students():
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM students WHERE PROCESSED = 0')
    new_students = cursor.fetchall()
    conn.close()

    # Convert result to DataFrame
    columns = ['ADMNO', 'NAME', 'GRADE', 'STREAM', 'PROCESSED']
    return pd.DataFrame(new_students, columns=columns)

def mark_students_as_processed(admnos):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.executemany('UPDATE students SET PROCESSED = 1 WHERE ADMNO = ?', [(admno,) for admno in admnos])
    conn.commit()
    conn.close()

def draw_card(c, x, y, name, admno, grade, stream, validity):
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
    c.drawCentredString(x + CARD_WIDTH / 2, text_y - 0.6 * inch, f"GRADE: {grade} {stream}")

    # Validity
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(x + CARD_WIDTH / 2, text_y - 0.9 * inch, f"VALIDITY: {validity}")


# Function to adjust font size for long text
# def draw_text_with_dynamic_font_size(
#     c, text, x, y, max_width, max_font_size=10, min_font_size=6
# ):
#     font_size = max_font_size
#     c.setFont("Helvetica-Bold", font_size)

#     # Check the width of the text at the current font size
#     while (
#         font_size >= min_font_size
#         and c.stringWidth(text, "Helvetica-Bold", font_size) > max_width
#     ):
#         font_size -= 1  # Reduce font size if the text is too wide

#     c.setFont("Helvetica-Bold", font_size)
#     c.drawString(x, y, text)

def create_dataframe(input_file):
    # Initialize an empty list to store dataframes
    df_list = []
    # Add a column to each dataframe for the sheet name
    for sheet in sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet, index_col=[0])
        df['Sheet'] = sheet  # Add the sheet name as a new column
        # Convert 'ADMNO' column to string to remove decimals
        df['ADMNO'] = df['ADMNO'].astype(str).str.replace(r'\.0$', '', regex=True)
        df_list.append(df)

    # Concatenate the dataframes
    return pd.concat(df_list, ignore_index=True)

def generate_pdf(data, output_file, validity: str):

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


def updated_cards(new_data_file, output_file):
    # Read new data from the Excel sheet
    new_df = create_dataframe(new_data_file)

    # Store new students in the database
    store_new_students(new_df)

    # Get unprocessed students from the database
    unprocessed_students = get_unprocessed_students()

    # If there are unprocessed students, generate ID cards
    if not unprocessed_students.empty:
        print(f"Found {len(unprocessed_students)} new students to process.")
        generate_pdf(unprocessed_students, output_file, validity="25/10/2024")

        # Mark the processed students as processed
        mark_students_as_processed(unprocessed_students['ADMNO'].tolist())
    else:
        print("No new students to process.")

# Usage
create_students_table()
input_file = 'data.xlsx'
output_file = f"meal_cards({datetime.now().strftime('%d/%m')}).pdf"
updated_cards(input_file, output_file)

# # Usage
# # input_file = "TERM3 MEAL CARDS 1-6(updates).xlsx"
# input_file = 'data.xlsx'
# output_file = f"id_cards({datetime.now().strftime('%m-%d')}).pdf"
# # generate_pdf(input_type='file', data=input_file, output_file=output_file, validity="25/10/2024")
# updated_cards(input_file, output_file)
