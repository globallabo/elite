# Elite Business English Materials Generation

![Google Sheets + Jinja + Python = PDF](/GSplusJinja.png)

The Elite materials can be automatically generated using data from Google Sheets and formatting from Jinja HTML templates.

These scripts make it possible to automatically generate a full set of curriculum materials. The templates are made using Jinja for fine control of layout and typography. The lesson content, such as vocabulary words and target sentences, is kept in a Google Sheets spreadsheet for best visibility to the team and so any team member can easily make changes to the content. The Weasyprint package for Python is used to convert the HTML and CSS to PDF files which are then distributed to the team and printed for use in the classroom.

This system makes it easy to create and maintain a large and complex set of materials. For example, in Elite, there are three distinct levels, each with six units of three lessons each. Any member of the team can change the lesson content. And if the appearance or layout needs to be changed, only the template needs to be changed and the scripts will output all the materials in seconds. Doing this by hand would require editing up to 54 separate documents.

## Example Output

![Elite text pages](/elite-text-example.png)

## Usage

### Create Presentation Slides

The materials are presented as slides which can be printed or used in an online presentation or lesson. To create them, run:

```bash
python elite/makepresentation.py
```

By default, this will make presentation slides for all levels, units and lessons. And it will save them to a folder named "output" at the root of the package.

If you wish to only process specific levels, use the `--level / -L` option. For example, if you want to only make slides for Level 2:

```bash
python elite/makepresentation.py --level 2
```

To make slides for multiple levels, repeat the option. For example, to make slides for both Levels 1 and 2:

```bash
python elite/makepresentation.py -L 1 -L 2
```

The same applies to units, using the `--unit / -u` option. While this can be used independently of the level option, typically, they'll be used together. For example, if you want to only make slides for Level 2, Unit 4:

```bash
python elite/makepresentation.py --level 2 --unit 4
```

To make slides for multiple units, repeat the option. For example, to make slides for both Units 3 and 4 of Level 1:

```bash
python elite/makepresentation.py -L 1 -u 3 -u 4
```

And these also apply to lessons, using the `--lesson / -l` option (note that level uses upper case, and lesson uses lower case). While this can be used independently of the level and unit options, typically, they'll be used together. For example, if you want to only make slides for Level 2, Unit 4, Lesson 1:

```bash
python elite/makepresentation.py --level 2 --unit 4 --lesson 1
```

To make slides for multiple lessons, repeat the option. For example, to make slides for both Lessons 2 and 3 of Level 1, Unit 2:

```bash
python elite/makepresentation.py -L 1 -u 2 -l 2 -l 3
```

If you wish to save your slide PDFs to another directory, use the `--outputpath / -o` option. For example, to save to your home directory:

```bash
python elite/makepresentation.py -o ~/
```
