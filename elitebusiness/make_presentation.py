# standard library imports
from pathlib import Path
from typing import Optional
import logging

# third party imports
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from weasyprint import HTML
import jinja2
import typer


app = typer.Typer(help="Create presentation files for Elite lessons")


def underline_vocab(target: str, vocabs: list[str]) -> str:
    """Underline the words in a target sentence that are part of the lesson's vocabulary if they are found.

    Args:
        target (str): A target sentence to search.
        vocabs (list[str]): A list of that lesson's vocabulary.

    Returns:
        str: A new target sentence with HTML to underline the vocabulary within it, or ehe same target sentence if the vocabulary isn't found..
    """
    for vocab in vocabs:
        if vocab in target:
            uvocab = f"<u>{vocab}</u>"
            newtarget = target.replace(vocab, uvocab)
            return newtarget
    return target


def get_data_for_level(level: str) -> list[list[str]]:
    """Get data from Google Sheet for a given level.

    Args:
        level (str): The grade level of the data to get.

    Returns:
        list[list[str]]: The contents of the sheet in list form (two-dimensional).
    """
    # Fetch data from Google Sheet
    sheetname = "Copy of ビジネス英語研修_online_対訳付き"
    credsfile = Path(__file__).parent.parent.resolve() / "creds.json"
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(credsfile, scope)
    client = gspread.authorize(creds)
    sheet = client.open(sheetname).worksheet(f"level_{level}")
    return sheet.get_all_values()


def create_template_mapping(
    data: list[list[str]], level: int, unit: int, lesson: int
) -> dict[str, str]:
    """Create a dict to map the template variables to their values for one lesson.

    Args:
        data (list): The data for a level in a two-dimensional list.
        level (int): The level number.
        unit (int): The unit number.
        lesson (int): The lesson number.

    Returns:
        dict[str, str]: A dict which maps template variables to to their values.
    """
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
    template_mapping["static_path"] = Path(__file__).parent.resolve() / "static/"
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


def render_template(template_mapping: dict[str, str]) -> str:
    """Render a Jinja template into HTML.

    Args:
        template_mapping (dict[str, str]): A dict containing a map of template variables and their values.

    Returns:
        str: The HTML content to use to make a PDF.
    """
    template_loader = jinja2.FileSystemLoader(searchpath="./templates/")
    template_env = jinja2.Environment(loader=template_loader)
    TEMPLATE_FILE = "base.html"
    template = template_env.get_template(TEMPLATE_FILE)
    output_text = template.render(template_mapping)
    return output_text


# Output PDF
def output_pdf(contents: str, filename: str) -> None:
    """Create a PDF from HTML using Weasyprint.

    Args:
        contents (str): The source HTML.
        filename (str): The filename of the PDF to create.
    """
    # Log WeasyPrint output
    logger = logging.getLogger("weasyprint")
    logger.addHandler(logging.FileHandler("/tmp/weasyprint.log"))
    # Create Weasyprint HTML object
    html = HTML(string=contents)
    # Output PDF via Weasyprint
    html.write_pdf(filename)


@app.command()
def presentations(
    levels: Optional[list[int]] = typer.Option(
        None,
        "--level",
        "-L",
        help="Specify a level to create a presentation for. Can be repeated for multiple levels.",
        show_default="all levels",
    ),
    units: Optional[list[int]] = typer.Option(
        None,
        "--unit",
        "-u",
        help="Specify a unit to create a presentation for. Can be repeated for multiple units.",
        show_default="all units",
    ),
    lessons: Optional[list[int]] = typer.Option(
        None,
        "--lesson",
        "-l",
        help="Specify a lesson to create a presentation for. Can be repeated for multiple lessons.",
        show_default="all levels",
    ),
    output_path: str = typer.Option(
        Path(__file__).parent.parent.resolve() / "output/",
        "--outputpath",
        "-o",
        help="The path where the PDF files will be saved.",
    ),
) -> None:
    """Generate the PDF files for each lesson."""
    # Check if levels/units/lessons were set or if they should use defaults
    if not levels:
        levels = [1, 2, 3]
    if not units:
        units = [1, 2, 3, 4, 5, 6]
    if not lessons:
        lessons = [1, 2, 3]
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
                Path(f"{output_path}/Level {level}/").mkdir(parents=True, exist_ok=True)
                print(f"Output path: {output_path}/Level {level}/")
                # Output PDF
                output_filename = (
                    f"{output_path}/Level {level}/EB{level}U{unit}L{lesson}.pdf"
                )
                output_pdf(contents=presentation_contents, filename=output_filename)


if __name__ == "__main__":
    app()
