import gspread
from oauth2client.service_account import ServiceAccountCredentials
from weasyprint import HTML, CSS
from weasyprint.fonts import FontConfiguration
from string import Template
import logging
from typing import List


def underline_vocab(target: str, vocabs: List[str]) -> str:
    for vocab in vocabs:
        if vocab in target:
            uvocab = f"<u>{vocab}</u>"
            newtarget = target.replace(vocab, uvocab)
            return newtarget
    return target


# Function to get data from google Sheet
def get_data_for_level(level: str) -> List:
    # Fetch data from Google Sheet
    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/spreadsheets",
             "https://www.googleapis.com/auth/drive.file",
             "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open("Copy of ビジネス英語研修_online_対訳付き").worksheet(f"level_{level}")
    return sheet.get_all_values()


# Function to create template mapping

# Function to open HTML template file and substitute vars

# Function to output PDF

# Log WeasyPrint output
logger = logging.getLogger('weasyprint')
logger.addHandler(logging.FileHandler('/tmp/weasyprint.log'))

# There are 4 levels
# levels = [1, 2, 3, 4]
levels = [2]
# There are 6 units per level
# units = [1, 2, 3, 4, 5, 6]
units = [1]
# There are 4 lessons per unit, but the last is a review unit with no materials
lessons = [1, 2, 3]
# lessons = [3]

for level in levels:
    print(f'Level {level}')
    # Create HTML template
    font_config = FontConfiguration()
    template_filename = 'EB-presentation-template.html'
    with open(template_filename, "r") as template_file:
        template_file_contents = template_file.read()
    template_string = Template(template_file_contents)
    css_filename = "presentation.css"
    with open(css_filename, "r") as css_file:
        css_string = css_file.read()

    # output_path = f'/Users/cbunn/Documents/Employment/5 Star/Google Drive/All Stars Second Edition/All Stars Second Edition/Worksheets/Level {level}/'
    output_path = f'/Users/cbunn/projects/elitebusiness/output/Level {level}/'

    data = get_data_for_level(level)

    # col = sheet.col_values(3)
    # cell = sheet.cell(2, 2).value
    # print("Column 3:")
    # print(col)
    # print("Cell 2,2:")
    # print(cell)

    # Loop through all Units and Lessons
    for unit in units:
        for lesson in lessons:
            # Set the row based on the unit and lesson, column to 1
            # *NOTE* Google Sheets and gspread start numbering of rows
            #  and columns at 1, while the Python dict begins at 0,0
            row = 1 + ((unit - 1) * 13) + ((lesson - 1) * 4)
            column = 1
            # Unit titles are only listed next to the first lesson, so
            #  we need a special variable for that
            unit_row = 1 + ((unit - 1) * 13)

            # Create substitution mapping
            template_mapping = dict()
            template_mapping["level"] = level
            # These are used for the page header
            template_mapping["unit"] = unit
            template_mapping["lesson"] = lesson
            print(f'Level: {level}, Unit: {unit}, Lesson: {lesson}')
            template_mapping["unit_title"] = data[unit_row][column]
            template_mapping["lesson_title"] = data[row][column + 1]
            template_mapping["dialogue_a1_en"] = data[row][column + 4]
            template_mapping["dialogue_b1_en"] = data[row + 1][column + 4]
            template_mapping["dialogue_a2_en"] = data[row + 2][column + 4]
            template_mapping["dialogue_b2_en"] = data[row + 3][column + 4]
            template_mapping["dialogue_a1_jp"] = data[row][column + 5]
            template_mapping["dialogue_b1_jp"] = data[row + 1][column + 5]
            template_mapping["dialogue_a2_jp"] = data[row + 2][column + 5]
            template_mapping["dialogue_b2_jp"] = data[row + 3][column + 5]
            target_a_en = data[row][column + 7]
            target_b_en = data[row + 1][column + 7]
            vocab1_en = data[row][column + 8]
            vocab2_en = data[row + 1][column + 8]
            vocab3_en = data[row + 2][column + 8]
            vocab4_en = data[row + 3][column + 8]
            template_mapping["target_a_en"] = underline_vocab(target_a_en, [vocab1_en, vocab2_en, vocab3_en, vocab4_en])
            template_mapping["target_b_en"] = underline_vocab(target_b_en, [vocab1_en, vocab2_en, vocab3_en, vocab4_en])
            template_mapping["vocab1_en"] = data[row][column + 8]
            template_mapping["vocab2_en"] = data[row + 1][column + 8]
            template_mapping["vocab3_en"] = data[row + 2][column + 8]
            template_mapping["vocab4_en"] = data[row + 3][column + 8]
            template_mapping["vocab1_jp"] = data[row][column + 9]
            template_mapping["vocab2_jp"] = data[row + 1][column + 9]
            template_mapping["vocab3_jp"] = data[row + 2][column + 9]
            template_mapping["vocab4_jp"] = data[row + 3][column + 9]
            template_mapping["extension1_en"] = data[row][column + 10]
            template_mapping["extension2_en"] = data[row + 1][column + 10]
            template_mapping["extension3_en"] = data[row + 2][column + 10]
            template_mapping["extension1_jp"] = data[row][column + 11]
            template_mapping["extension2_jp"] = data[row + 1][column + 11]
            template_mapping["extension3_jp"] = data[row + 2][column + 11]

            # Substitute
            template_filled = template_string.safe_substitute(template_mapping)
            html = HTML(string=template_filled)

            # The numbers used in the filename DO NOT need to be zero filled
            f_level = str(level)
            f_unit = str(unit)
            f_lesson = str(lesson)

            html.write_pdf(f'{output_path}EB{f_level}U{f_unit}L{f_lesson}.pdf')
