import gspread
from oauth2client.service_account import ServiceAccountCredentials
from weasyprint import HTML, CSS
from weasyprint.fonts import FontConfiguration
from string import Template
import logging

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

    # Fetch data from Google Sheet
    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/spreadsheets",
             "https://www.googleapis.com/auth/drive.file",
             "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open("ビジネス英語研修_online_対訳付き").worksheet(f"level_{level}")
    data = sheet.get_all_values()
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
            template_mapping["dialogue_a1_jp"] = "悩んでいるみたいですね。"
            template_mapping["dialogue_b1_jp"] = "そうなんですよ。このデータの作成手伝ってくれませんか。"
            template_mapping["dialogue_a2_jp"] = "もちろんです。どうしましたか。"
            template_mapping["dialogue_b2_jp"] = "どうしても（データが）合わないんですよ。"
            # template_mapping["dialogue_a1_jp"] = data[row][column + 5]
            # template_mapping["dialogue_b1_jp"] = data[row + 1][column + 5]
            # template_mapping["dialogue_a2_jp"] = data[row + 2][column + 5]
            # template_mapping["dialogue_b2_jp"] = data[row + 3][column + 5]
            template_mapping["target_a_en"] = data[row][column + 7]
            template_mapping["target_b_en"] = data[row + 1][column + 7]
            template_mapping["vocab1_en"] = data[row][column + 8]
            template_mapping["vocab2_en"] = data[row + 1][column + 8]
            template_mapping["vocab3_en"] = data[row + 2][column + 8]
            template_mapping["vocab4_en"] = data[row + 3][column + 8]
            template_mapping["vocab1_jp"] = "データ"
            template_mapping["vocab2_jp"] = "報告書"
            template_mapping["vocab3_jp"] = "図式"
            template_mapping["vocab4_jp"] = "プレゼンテーション"
            template_mapping["extension1_en"] = data[row][column + 11]
            template_mapping["extension2_en"] = data[row + 1][column + 11]
            template_mapping["extension3_en"] = data[row + 2][column + 11]
            template_mapping["extension1_jp"] = "このデータの作成を手伝ってくれませんか。"
            template_mapping["extension2_jp"] = "こちらのデータの作成を手伝っていただけませんか。"
            template_mapping["extension3_jp"] = "お願いなんだけれど、このデータの作成を手伝ってくれない？"

            # template_mapping["unit_title"] = "Asking for help"
            # template_mapping["dialogue_a1_en"] = "You look stressed."
            # template_mapping["dialogue_b1_en"] = "Yeah, I am. Can you help me with this data?"
            # template_mapping["dialogue_a2_en"] = "Of course. What's the problem?"
            # template_mapping["dialogue_b2_en"] = "It just doesn't add up."
            # template_mapping["dialogue_a1_jp"] = "悩んでいるみたいですね。"
            # template_mapping["dialogue_b1_jp"] = "そうなんですよ。このデータの作成手伝ってくれませんか。"
            # template_mapping["dialogue_a2_jp"] = "もちろんです。どうしましたか。"
            # template_mapping["dialogue_b2_jp"] = "どうしても（データが）合わないんですよ。"
            # template_mapping["target_a_en"] = "Can you help me with this <u>data</u>?"
            # template_mapping["target_b_en"] = "Of course."
            # template_mapping["vocab1_en"] = "data"
            # template_mapping["vocab2_en"] = "report"
            # template_mapping["vocab3_en"] = "schematic"
            # template_mapping["vocab4_en"] = "presentation"
            # template_mapping["vocab1_jp"] = "データ"
            # template_mapping["vocab2_jp"] = "報告書"
            # template_mapping["vocab3_jp"] = "図式"
            # template_mapping["vocab4_jp"] = "プレゼンテーション"
            # template_mapping["extension1_en"] = "Can you give me a hand with this <u>data</u>?"
            # template_mapping["extension2_en"] = "Would you mind helping me with this <u>data</u>?"
            # template_mapping["extension3_en"] = "Would you be a dear and help me with this <u>data</u>?"
            # template_mapping["extension1_jp"] = "このデータの作成を手伝ってくれませんか。"
            # template_mapping["extension2_jp"] = "こちらのデータの作成を手伝っていただけませんか。"
            # template_mapping["extension3_jp"] = "お願いなんだけれど、このデータの作成を手伝ってくれない？"

            # Substitute
            template_filled = template_string.safe_substitute(template_mapping)
            html = HTML(string=template_filled)

            # The numbers used in the filename DO NOT need to be zero filled
            f_level = str(level)
            f_unit = str(unit)
            f_lesson = str(lesson)

            html.write_pdf(f'{output_path}EB{f_level}U{f_unit}L{f_lesson}.pdf')
