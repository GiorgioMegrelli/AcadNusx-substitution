# AcadNusx-substitution

The Python script to replace 'AcadNusx' font in DOCX elements

### Usage

_packages in `requirements.txt` should be installed_

```shell
python main.py --help
```

```shell
Usage: main.py [OPTIONS]

Options:
  -i, --input-file TEXT   Input file.  [required]
  -o, --output-file TEXT  Output file.  [required]
  --help                  Show this message and exit.
```

Example:

```shell
python main.py -i document.docx -o document_new.docx
```

### Examples:

**#1**

_The incorrect word in **AcadNusx**:_

![Example 1 in AN](media/example_1_an.png)

==>

_The correct word in **Georgian** after conversion:_

![Example 1 in Geo](media/example_1_geo.png)

**#2**

_The incorrect word in **AcadNusx**:_

![Example 2 in AN](media/example_2_an.png)

==>

_The correct word in **Georgian** after conversion:_

![Example 2 in Geo](media/example_2_geo.png)

### Structure

```text
- chars.py    # Characters' mapping
- errors.py   # Error classes
- filters.py  # Filters to sieve
- main.py     # Main runner script
```
