import itertools
from collections import namedtuple
from pathlib import Path

import typer
import xlsxwriter
import yaml

# Some of the colours are a bit hard to distinguish.
colours = [
    "#FF1493",  # deeppink
    # '#FF69B4',  # hotpink
    "#FFE4E1",  # mistyrose
    "#E6E6FA",  # lavender
    # '#B0E0E6',  # powderblue
    "#7FFFD4",  # aquamarine
    "#ADFF2F",  # greenyellow
    "#FFFFE0",  # lightyellow
    "#FFD700",  # gold
    "#FF7F50",  # coral
]

# Represents a single entry in an index.
Entry = namedtuple(
    "Entry", ["book", "book_colour", "chapter", "chapter_colour", "page", "topic"]
)


def create_workbook(
    course_content_path: Path = typer.Argument(
        ..., help="Location of YAML document containing SANS course content"
    ),
):
    """Creates an Excel workbook containing SANS course contents and index.

    The workbook is written to the current directory. It's filename is the same
    as COURSE_CONTENT_PATH, except it's file extension is .xlsx.
    """
    with open(course_content_path) as file:
        course_content = yaml.safe_load(file)

    # The workbook is written to the current directory with the same filename
    # as the course content filename, except with extension .xlsx.
    workbook = xlsxwriter.Workbook(f"{course_content_path.stem}.xlsx")

    # Create cell formats with backgrounds colours set to our colour values.
    formats = [workbook.add_format({"bg_color": colour}) for colour in colours]

    # Book titles are bold, chapter names are italic.
    book_title_format = workbook.add_format({"bold": True})
    chapter_name_format = workbook.add_format({"italic": True})

    # Cycle through every third format for better colour contrast between successive
    # cell formats.
    book_formats = itertools.islice(itertools.cycle(formats), 0, None, 3)
    chapter_formats = itertools.islice(itertools.cycle(formats), 1, None, 3)

    # Holds a sequence of index entries to be sorted alphabetically.
    entries = []

    contents = workbook.add_worksheet("Contents")

    row = 0
    for book in course_content:
        # Each book has a title and chapters.
        title, chapters = book.popitem()

        # All entries have the same book colour.
        book_format = next(book_formats)

        # Write the book's title.
        contents.write(row, 0, title, book_title_format)
        row += 1

        for chapter in chapters:
            # Each chapter has a name and indices.
            name, indices = chapter.popitem()

            # All of this chapter's entries have the same chapter colour.
            chapter_format = next(chapter_formats)

            # Add an empty cell with the book's colour.
            contents.write(row, 0, None, book_format)

            # Write the chapter's name.
            contents.write(row, 1, name, chapter_name_format)
            row += 1

            for course_content in indices:
                # Each index has has a page number and topic.
                page, topic = course_content.popitem()

                # Add empty cells with book colour and chapter colour.
                contents.write(row, 0, None, book_format)
                contents.write(row, 1, page, chapter_format)

                # Write the topic.
                contents.write(row, 2, topic)
                row += 1

                # Collect this entry.
                entries.append(
                    Entry(title, book_format, name, chapter_format, page, topic)
                )

    index = workbook.add_worksheet("Index")

    # Write index entries in alphabetical order.
    for row, (book, book_format, chapter, chapter_format, page, topic) in enumerate(
        sorted(entries, key=lambda topic: topic.topic.lower())
    ):
        index.write(row, 0, topic)
        index.write(row, 1, page, chapter_format)
        index.write(row, 2, chapter, chapter_format)
        index.write(row, 3, book, book_format)

    workbook.close()


if __name__ == "__main__":
    typer.run(create_workbook)
