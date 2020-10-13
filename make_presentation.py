import gspread
from oauth2client.service_account import ServiceAccountCredentials
from weasyprint import HTML, CSS
from weasyprint.fonts import FontConfiguration
from string import Template
import logging

# Log WeasyPrint output
logger = logging.getLogger('weasyprint')
logger.addHandler(logging.FileHandler('/tmp/weasyprint.log'))

levels = [1]
# units = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
units = [1]
# units = [4, 6, 10, 14]
lessons = [1]

for level in levels:
    print(f'Level {level}')
    # Create HTML template
    font_config = FontConfiguration()
    template_filename = f'BN{level}-presentation-template.html'
    with open(template_filename, "r") as template_file:
        template_file_contents = template_file.read()
    template_string = Template(template_file_contents)
    css_filename = "presentation.css"
    with open(css_filename, "r") as css_file:
        css_string = css_file.read()

    # output_path = f'/Users/cbunn/Documents/Employment/5 Star/Google Drive/All Stars Second Edition/All Stars Second Edition/Worksheets/Level {level}/'
    output_path = f'/Users/cbunn/Documents/Employment/5 Star/Google Drive/Business Next/test-output/Level {level}/'

    # Fetch data from Google Sheet
    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/spreadsheets",
             "https://www.googleapis.com/auth/drive.file",
             "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open("businessnext-test").worksheet("level_1")
    data = sheet.get_all_values()

    # Set the starting point of the gspread output
    row = 1
    column = 3

    # Loop through all Units and Lessons
    #  (range() needs a +1 because it stops at the number before)
    for unit in units:
        for lesson in lessons:
            # Create substitution mapping
            template_mapping = dict()
            template_mapping["level"] = level
            # These are used for the page header
            template_mapping["unit"] = unit
            template_mapping["lesson"] = lesson
            print(f'Level: {level}, Unit: {unit}, Lesson: {lesson}')

            template_mapping["unit_title"] = "Asking for help"
            template_mapping["dialogue_a1_en"] = "You look stressed."
            template_mapping["dialogue_b1_en"] = "Yeah, I am. Can you help me with this data?"
            template_mapping["dialogue_a2_en"] = "Of course. What's the problem?"
            template_mapping["dialogue_b2_en"] = "It just doesn't add up."
            template_mapping["dialogue_a1_jp"] = "悩んでいるみたいですね。"
            template_mapping["dialogue_b1_jp"] = "そうなんですよ。このデータの作成手伝ってくれませんか。"
            template_mapping["dialogue_a2_jp"] = "もちろんです。どうしましたか。"
            template_mapping["dialogue_b2_jp"] = "どうしても（データが）合わないんですよ。"
            template_mapping["target_a_en"] = "Can you help me with this <u>data</u>?"
            template_mapping["target_b_en"] = "Of course."
            template_mapping["vocab1_en"] = "data"
            template_mapping["vocab2_en"] = "report"
            template_mapping["vocab3_en"] = "schematic"
            template_mapping["vocab4_en"] = "presentation"
            template_mapping["vocab1_jp"] = "データ"
            template_mapping["vocab2_jp"] = "報告書"
            template_mapping["vocab3_jp"] = "図式"
            template_mapping["vocab4_jp"] = "プレゼンテーション"
            template_mapping["extension1_en"] = "Can you give me a hand with this <u>data</u>?"
            template_mapping["extension2_en"] = "Would you mind helping me with this <u>data</u>?"
            template_mapping["extension3_en"] = "Would you be a dear and help me with this <u>data</u>?"
            template_mapping["extension1_jp"] = "このデータの作成を手伝ってくれませんか。"
            template_mapping["extension2_jp"] = "こちらのデータの作成を手伝っていただけませんか。"
            template_mapping["extension3_jp"] = "お願いなんだけれど、このデータの作成を手伝ってくれない？"

            # Substitute
            template_filled = template_string.safe_substitute(template_mapping)
            html = HTML(string=template_filled)

            # The numbers used in the filename need to be zero filled
            f_level = str(level)
            f_unit = str(unit).zfill(2)
            f_lesson = str(lesson).zfill(2)

            html.write_pdf(f'{output_path}bn{f_level}U{f_unit}L{f_lesson}.pdf')
