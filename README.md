# sans-index

Generates a colourful Excel workbook from a YAML file containing SANS course content.

Read this [blog post](https://brodie.id.au/blog/sans-course-index.html)
for more information about this project.


## Quick start

1. Install Python 3.8+.
2. Install [Poetry](https://python-poetry.org/docs/).
3. Clone this repository.
4. Run `poetry install`.
5. Run `poetry shell`.
6. Generate an Excel workbook (`sans-course.xlsx`) from a YAML file (`sans-course.yml`):

       python sans-index.py sans-course.yml


## SANS course outline

Create a YAML document of SANS course content. The root of the document is a sequence of the course's books.

```yaml
- <book>
- <book>
- <book>
```

Books map a book title to a sequence of chapters.

```yaml
- Threat intelligence:
    - <chapter>
    - <chapter>
    - <chapter>
```

Chapters map a chapter name to a sequence of topics/keywords.

```yaml
- Threat intelligence:
    - Case study - Stuxnet:
        - <topic>
        - <topic>
        - <topic>
```

Topics map a page number to topic.

```yaml
- Threat intelligence:
    - Case study - Stuxnet:
        - 10: Stuxnet
        - 10: Case study - Stuxnet
        - 10: Iran's nuclear program
    - Introduction to active defense and incident response:
        - 18: Sliding scale of cyber security
        - 18: Architecture
        - 18: Passive defense
```


## `sans-index.py`

    % python sans-index.py --help

    Usage: sans-index.py [OPTIONS] COURSE_CONTENT_PATH

    Creates an Excel workbook containing SANS course contents and index.

    The workbook is written to the current directory. It's filename is the same
    as COURSE_CONTENT_PATH, except it's file extension is .xlsx.

    Arguments:
    COURSE_CONTENT_PATH  Location of YAML document containing SANS course
                        content  [required]

    Options:
    --install-completion [bash|zsh|fish|powershell|pwsh]
                                    Install completion for the specified shell.
    --show-completion [bash|zsh|fish|powershell|pwsh]
                                    Show completion for the specified shell, to
                                    copy it or customize the installation.
    --help                          Show this message and exit.
