import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import os

# Function to generate a report card for a single student
def generate_report_card(student_id, name, scores_df, total_score, avg_score):
    filename = f"report_card_{student_id}.pdf"
    doc = SimpleDocTemplate(filename, pagesize=letter)
    styles = getSampleStyleSheet()
    
    
    # Title
    title = Paragraph(f"<b>Report Card</b>", styles["Title"])
    
    # Student information
    student_info = Paragraph(f"<b>Name:</b> {name}<br/><b>Student ID:</b> {student_id}", styles["Normal"])
    
    # Scores summary
    summary = Paragraph(f"<b>Total Score:</b> {total_score}<br/><b>Average Score:</b> {avg_score:.2f}", styles["Normal"])
    
    # Subject-wise scores table
    data = [["Subject", "Score"]] + scores_df.values.tolist()
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    # Content flow
    elements = [title, Spacer(1, 12), student_info, Spacer(1, 12), summary, Spacer(1, 12), table]
    
    # Build the PDF
    doc.build(elements)
    print(f"Report card generated: {filename}")

# Main script
def main():
    # File path
    input_file = "student_scores.xlsx"
    
    # Check if file exists
    if not os.path.exists(input_file):
        print(f"Error: File '{input_file}' not found.")
        return

    try:
        # Read the Excel file
        df = pd.read_excel(input_file)
        df.columns = df.columns.str.strip()  # Removes extra spaces from column names

        # Validate required columns
        required_columns = {"Student ID", "Name", "Subject", "Score"}
        if not required_columns.issubset(df.columns):
            print("Error: Missing required columns in the Excel file.")
            return

        # Group data by student
        grouped = df.groupby(["Student ID", "Name"])

        for (student_id, name), group in grouped:
            total_score = group["Score"].sum()
            avg_score = group["Score"].mean()
            scores_df = group[["Subject", "Score"]]
            generate_report_card(student_id, name, scores_df, total_score, avg_score)

    except Exception as e:
        print(f"Error processing the file: {e}")

if __name__ == "__main__":
    main()
