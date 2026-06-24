Build a complete production-ready Python certificate generation system.

PROJECT GOAL

Generate PNG certificates automatically from an Excel file and a certificate template image.

The system should read every row from Excel, populate the certificate fields, generate a unique certificate number, and export one PNG certificate per participant.

TECH STACK

* Python 3.9+
* pandas
* Pillow (PIL)
* openpyxl

INSTALLATION

pip3 install pandas pillow openpyxl

PROJECT STRUCTURE

certificate-generator/

├── template.png
├── Football Merit Certificates.xlsx
├── generate.py
├── fonts/
│   └── Arial.ttf
└── output/

EXCEL COLUMNS

The Excel file contains the following columns exactly:

* Full Name
* Father Name
* Mother Name
* DOB
* Aadhaar Number
* Position
* Sport Event
* Sport
* District

CERTIFICATE NUMBER

Generate a unique certificate number for each row.

Format:

0001
0002
0003
0004

Use row order.

The generated certificate number should be rendered near:

HOA/HSG/M/0001

OUTPUT FILE NAME FORMAT

0001_Deepshikha.png
0002_Rahul_Kumar.png
0003_Amit_Singh.png

DATA MAPPING

Map Excel columns to certificate fields:

Full Name
→ Mr./Ms. __________

Father Name
→ Father's Name __________

Mother Name
→ Mother's Name __________

DOB
→ Date of Birth __________

Aadhaar Number
→ Aadhaar No. __________

Position
→ secured ______ position

Sport Event
→ event in the ______

Sport
→ sports discipline

District
→ held from 2nd November to 8th November 2025 at ______ District

Certificate Number
→ HOA/HSG/M/0001

IMPORTANT REQUIREMENTS

1. DO NOT hardcode field coordinates.

2. Create a calibration mode.

Example:

CALIBRATION_MODE = True

When enabled:

* Draw colored rectangles for every field.
* Draw field labels.
* Draw coordinate guides.
* Generate only first certificate.

Example labels:

FULL_NAME
FATHER_NAME
MOTHER_NAME
DOB
AADHAAR
POSITION
SPORT_EVENT
SPORT
DISTRICT
CERTIFICATE_NO

This mode helps identify exact locations visually.

3. Create a coordinate configuration section.

Example:

FIELD_CONFIG = {
"full_name": {
"x": 0,
"y": 0,
"max_width": 400
}
}

All field positions must come from configuration.

4. Create helper function:

fit_text_to_width()

Requirements:

* Automatically reduce font size
* Prevent overflow
* Keep text inside allotted area
* Works for long names

5. Preserve original image resolution.

Never resize template image.

6. Aadhaar number must always be treated as string.

Do not allow scientific notation.

7. Format DOB as:

DD-MM-YYYY

8. Create output folder automatically.

9. Print progress:

Generated:
0001_Deepshikha.png

10. Exception handling:

* Missing Excel file
* Missing template image
* Missing font
* Missing columns
* Invalid data

11. Use clean modular code.

Functions:

load_excel()
generate_certificate()
draw_field()
fit_text_to_width()
create_calibration_preview()
main()

CALIBRATION WORKFLOW

Phase 1:

* Enable CALIBRATION_MODE
* Generate only one certificate
* Show colored field boxes
* Manually identify locations

Phase 2:

* Update coordinates
* Disable calibration mode

CALIBRATION_MODE = False

* Generate all certificates

FINAL OUTPUT

When CALIBRATION_MODE is False:

Read every row from Excel.

Generate:

output/
├── 0001_Deepshikha.png
├── 0002_Himanshi.png
├── 0003_Khushbu.png
...
├── 2600_xxxxx.png

At completion print:

---

Certificates Generated: XXXX
Output Folder: output
---------------------

Generate complete runnable generate.py code with comments and production-quality structure.
