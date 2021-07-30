import gspread
from oauth2client.service_account import ServiceAccountCredentials
from weasyprint import HTML
import pathlib
import logging
import jinja2


# Find vocabulary in target sentence and underline it
def underline_vocab(target: str, vocabs: list[str]) -> str:
    for vocab in vocabs:
        if vocab in target:
            uvocab = f"<u>{vocab}</u>"
            newtarget = target.replace(vocab, uvocab)
            return newtarget
    return target


# Get data from google Sheet (one level at a time)
def get_data_for_level(level: str) -> list[str]:
    # Fetch data from Google Sheet
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open("Copy of ビジネス英語研修_online_対訳付き").worksheet(f"level_{level}")
    return sheet.get_all_values()


# Create template mapping for one lesson at a time
def create_template_mapping(
    data: list, level: int, unit: int, lesson: int
) -> dict[str, str]:
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
    template_mapping["static_path"] = pathlib.Path(__file__).parent.resolve() / "static/"
    template_mapping["level"] = level
    # These are used for the page header
    template_mapping["unit"] = unit
    template_mapping["lesson"] = lesson
    print(f"Level: {level}, Unit: {unit}, Lesson: {lesson}")
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
    template_mapping["target_a_en"] = underline_vocab(
        target_a_en, [vocab1_en, vocab2_en, vocab3_en, vocab4_en]
    )
    template_mapping["target_b_en"] = underline_vocab(
        target_b_en, [vocab1_en, vocab2_en, vocab3_en, vocab4_en]
    )
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
    return template_mapping


# Use Jinja to render a template into HTML
def render_template(template_mapping: dict[str, str]) -> str:
    template_loader = jinja2.FileSystemLoader(searchpath="./templates/")
    template_env = jinja2.Environment(loader=template_loader)
    TEMPLATE_FILE = "base.html"
    template = template_env.get_template(TEMPLATE_FILE)
    output_text = template.render(template_mapping)
    return output_text

# Output PDF
def output_pdf(contents: str, filename: str):
    # Log WeasyPrint output
    logger = logging.getLogger("weasyprint")
    logger.addHandler(logging.FileHandler("/tmp/weasyprint.log"))
    # Create Weasyprint HTML object
    html = HTML(string=contents)
    # Output PDF via Weasyprint
    html.write_pdf(filename)


def main(
    levels: list,
    units: list,
    lessons: list,
    output_path: str = pathlib.Path(__file__).parent.parent.absolute() / "output/",
):
    for level in levels:
        print(f"Level {level}")
        # Get data from Google Sheet
        data = get_data_for_level(level)
        # Loop through all Units and Lessons
        for unit in units:
            for lesson in lessons:
                # create mapping dict
                template_mapping = create_template_mapping(
                    data=data, level=level, unit=unit, lesson=lesson
                )
                presentation_contents = render_template(template_mapping)
                # Create final path if it doesn't exist
                pathlib.Path(f"{output_path}/Level {level}/").mkdir(parents=True, exist_ok=True)
                print(f"Output path: {output_path}/Level {level}/")
                # Output PDF
                output_filename = (
                    f"{output_path}/Level {level}/EB{level}U{unit}L{lesson}.pdf"
                )
                output_pdf(contents=presentation_contents, filename=output_filename)


# # There are 4 levels
# # levels = [1, 2, 3, 4]
# levels = [1, 2, 3]
# # There are 6 units per level
# # units = [1, 2, 3, 4, 5, 6]
# units = [2]
# # There are 4 lessons per unit, but the last is a review unit with no materials
# lessons = [1, 2, 3]
# # lessons = [3]
# # Filename for HTML Template
# template_filename = 'EB-presentation-template.html'
# # output_path = f'/Users/cbunn/Documents/Employment/5 Star/Google Drive/All Stars Second Edition/All Stars Second Edition/Worksheets/Level {level}/'
# output_path = '/Users/cbunn/projects/elitebusiness/output/'
if __name__ == "__main__":
    # There are 4 levels, but level 4 isn't ready
    # levels = [1, 2, 3]
    # units = [1, 2, 3, 4, 5, 6]
    # lessons = [1, 2, 3]
    levels = [2]
    units = [4]
    lessons = [1, 2, 3]
    main(levels, units, lessons)
