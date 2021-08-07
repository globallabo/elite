# Elite Business English Program

The Elite materials can be automatically generated using data from Google Sheets and formatting from Jinja HTML templates.

## Usage

### Create Presentation Slides

The materials are presented as slides which can be printed or used in an online presentation or lesson. To create them, run:

```bash
python elitebusiness/makepresentation.py
```

By default, this will make presentation slides for all levels, units and lessons. And it will save them to a folder named "output" at the root of the package.

If you wish to only process specific levels, use the `--level / -L` option. For example, if you want to only make slides for Level 2:

```bash
python elitebusiness/makepresentation.py --level 2
```

To make slides for multiple levels, repeat the option. For example, to make slides for both Levels 1 and 2:

```bash
python elitebusiness/makepresentation.py -L 1 -L 2
```

The same applies to units, using the `--unit / -u` option. While this can be used independently of the level option, typically, they'll be used together. For example, if you want to only make slides for Level 2, Unit 4:

```bash
python elitebusiness/makepresentation.py --level 2 --unit 4
```

To make slides for multiple units, repeat the option. For example, to make slides for both Units 3 and 4 of Level 1:

```bash
python elitebusiness/makepresentation.py -L 1 -u 3 -u 4
```

And these also apply to lessons, using the `--lesson / -l` option (note that level uses upper case, and lesson uses lower case). While this can be used independently of the level and unit options, typically, they'll be used together. For example, if you want to only make slides for Level 2, Unit 4, Lesson 1:

```bash
python elitebusiness/makepresentation.py --level 2 --unit 4 --lesson 1
```

To make slides for multiple lessons, repeat the option. For example, to make slides for both Lessons 2 and 3 of Level 1, Unit 2:

```bash
python elitebusiness/makepresentation.py -L 1 -u 2 -l 2 -l 3
```

If you wish to save your slide PDFs to another directory, use the `--outputpath / -o` option. For example, to save to your home directory:

```bash
python elitebusiness/makepresentation.py -o ~/
```
