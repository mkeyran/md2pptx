# md2pptx

md2pptx is a Python tool that converts Markdown files to PowerPoint presentations.

## Installation

1. Ensure you have Python 3.10 or later installed.
2. Install Poetry if you haven't already:
   ```
   curl -sSL https://install.python-poetry.org | python3 -
   ```
3. Clone this repository:
   ```
   git clone https://github.com/yourusername/md2pptx.git
   cd md2pptx
   ```
4. Install the project dependencies using Poetry:
   ```
   poetry install
   ```

## Usage

You can run the tool using Poetry:

```
poetry run md2pptx INPUT_FILE OUTPUT_FILE [OPTIONS]
```

Or, if you've activated the Poetry shell:

```
md2pptx INPUT_FILE OUTPUT_FILE [OPTIONS]
```

Arguments:
- `INPUT_FILE`: Path to the input Markdown file
- `OUTPUT_FILE`: Path to the output PowerPoint file

Options:
- `--width FLOAT`: Slide width in inches (default: 16.0)
- `--height FLOAT`: Slide height in inches (default: 9.0)

Example:
```
poetry run md2pptx input.md output.pptx --width 13.33 --height 7.5
```

## Markdown Format

The tool supports the following Markdown elements:

1. Slides are separated by `---` on a new line.
2. The first line of each slide is treated as the slide title.
3. Headers are created using `#`, `##`, `###`, etc.
4. Bullet points are created using `-` or `*`.
5. Numbered lists are created using `1.`, `2.`, etc.
6. Images are inserted using the standard Markdown syntax: `![alt text](image_url)`

Example:

```markdown
# Slide 1 Title

- Bullet point 1
- Bullet point 2

## Subheader

1. Numbered item 1
2. Numbered item 2

![Image description](https://example.com/image.jpg)

---

# Slide 2 Title

Content for slide 2
```

## Limitations

- The tool currently does not support advanced PowerPoint features like animations or transitions.
- Image positioning is fixed and may need manual adjustment in the resulting PowerPoint file.


## Testing

To run the tests, use the following command:

```
poetry run pytest
```

This will run all the tests in the `tests` directory. You can also run specific test files or functions by providing their names to pytest.

...

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. When contributing, please ensure that you add appropriate tests for any new functionality or bug fixes.
