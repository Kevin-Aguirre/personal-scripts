from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
import sys


def format_template_for_company(pdfText, arguments):
	curr_date = datetime.now()
	formatted_date = curr_date.strftime("%B %-d, %Y")

	comp_name, addr_line1, addr_line2, position = arguments[1], arguments[2], arguments[3], arguments[4]
	newTxt = pdfText.replace("[COMPANY NAME]", comp_name).replace("[COMPANY ADDRESS LINE 1]", addr_line1).replace("[COMPANY ADDRESS LINE 2]", addr_line2).replace("[POSITION]", position).replace("[DATE]", formatted_date)
	return newTxt

def create_pdf(formatted_pdf_text, output_path = "KevinAguirre_coverletter.pdf"):
	document = SimpleDocTemplate(output_path, pagesize=letter)
	styles = getSampleStyleSheet()
		# Define a centered style
	centered_style = styles['Normal'].clone('CenteredStyle')
	centered_style.alignment = 1  # 1 for center alignment
	centered_style.fontName = "Times-Roman"
	centered_style.leading = 14
	centered_style.fontSize = 28  # Set font size for centered text

	    # Define a right-aligned style
	right_aligned_style = styles['Normal'].clone('RightAlignedStyle')
	right_aligned_style.alignment = 2  # 2 for right alignment
	right_aligned_style.fontName = "Times-Roman"
	right_aligned_style.leading = 14
	right_aligned_style.fontSize = 12  # Set font size for right-aligned text


	default_style = styles['Normal']
	default_style.fontName = "Times-Roman"
	default_style.fontSize = 12
	default_style.leading = 14

	flowables = []

	lines = formatted_pdf_text.split('\n')

	for i, line in enumerate(lines):
		paragraph = None
		if i == 0: 
			paragraph = Paragraph(line, centered_style)
		elif i == 1:
			centered_style.fontSize = 10 
			paragraph = Paragraph(line, centered_style)
		elif len(lines) - 2 <= i <= len(lines) - 1:
			paragraph = Paragraph(line, right_aligned_style)
		else: 
			paragraph = Paragraph(line, default_style)

		flowables.append(paragraph)

		
		if i == 0:
			flowables.append(Spacer(1, 20))  # Add some space after each part
		
		if not(5 <= i < 7):
			flowables.append(Spacer(1, 6))  # Add some space after each part

		
	document.build(flowables)
       
def main(): 
	coverTemplatePath = 'main.txt'
	output_path = "KevinAguirre_coverletter.pdf"


	content = ''
	with open(coverTemplatePath, 'r') as file: content = file.read()
	
	formatted_pdf_text = format_template_for_company(content, sys.argv)
	create_pdf(formatted_pdf_text, output_path)

if __name__ == '__main__':
    main()
    


