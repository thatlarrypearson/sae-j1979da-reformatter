# SAE 1979DA Reformatter

Reformat [SAE OBD Interface Standard SAE-J1979](https://www.sae.org/standards/content/j1979da_202104/) from Excel to Word making the OBD standard information more readable, understandable and usable.

The digital Excel version of the [SAE J1979 standard](https://www.sae.org/standards/content/j1979da_202104/) can be purchased from SAE and is required to use this software.

## Motivation

The [SAE OBD Interface Standard SAE-J1979](https://www.sae.org/standards/content/j1979da_202104/) is difficult to read and work with.  A better method for organizing a subset of the data and placing it into a printable document was needed.  Now users like myself can easily pull an extract of the data in a usable easy to read format.

## Features

- Supports ```Annex B - Parameter IDs``` sheet transformation to Word
- Supports ```Annex G - InfoType IDs``` sheet transformation to word
- Supports three methods for specifying desired OBD commands:
  - ANNEX B hexadecimal parameter IDs in the form of 0x99
  - ANNEX G hexadecimal parameter IDs in the form of 0x99
  - Command names known by either ```python-OBD``` or ```telemetry-OBD``` such as ```SPEED```
- Commands are automatically placed in ANNEX B or G sections
- Each command starts a new page
- Word document file name can be specified and is included in the word document header
- Supports OBD commands returning multiple results
- Pages are numbered
- Resulting word document is editable and can be pasted into other documents

![Sample Output Page](docs/Page4.png)

## Usage

```bash
$ python3.8 -m j1929_reformatter.reformatter --help
usage: reformatter [-h] [--commands COMMANDS] [--annex_b ANNEX_B]
                   [--annex_g ANNEX_G] [--word WORD] [--xlsx XLSX] [--verbose]

Reformat Excel OBD Interface Standard SAE-J1979 to Word

optional arguments:
  -h, --help           show this help message and exit
  --commands COMMANDS  Command name list to include in Word report generation.
                       Command names come from 'telemetry-obd' (including
                       'python-OBD') package. Comma separated list. e.g.
                       "SPEED,RPM,FUEL_RATE".
  --annex_b ANNEX_B    Annex B Parameter IDs (PID) list to include in Word
                       report generation. Comma separated list. e.g.
                       "0x01,0x0F,0xAF".
  --annex_g ANNEX_G    Annex G Info Type IDs (PID) list to include in Word
                       report generation. Comma separated list. e.g.
                       "0x01,0x0F,0xAF".
  --word WORD          Word output file. File can be either a full or relative
                       path name. If the file already exists, it will be
                       overwritten. Defaults to 'SAE-J1979DA.docx'.
  --xlsx XLSX          Excel version of SAE standard document J1979DA file
                       name. Defaults to "SAE J1979DA_202104.xlsx".
  --verbose            Turn verbose output on. Default is off.

$
```

## Installation

### Required Supporting Libraries

- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [python-docx](https://python-docx.readthedocs.io/en/latest/)
- [python-OBD](https://github.com/brendan-w/python-OBD)
  - Requires manual install from source
- [telemetry-obd](https://github.com/thatlarrypearson/telemetry-obd-log-to-csv)
  - Requires manual install from source

```bash
# Python pip install support
python3.8 -m pip install --upgrade --user pip
python3.8 -m pip install --upgrade --user wheel setuptools markdown build

# Required libraries

# By default openpyxl does not guard against quadratic blowup or billion laughs xml attacks.
# To guard against these attacks install defusedxml.
python3.8 -m pip install --user openpyxl jinja2 python-docx docxtpl

# Recommended: install lxml library on Linux systems
sudo apt-get install -y libxml2 libxslt-dev

# Recommended: install Python lxml library
python3.8 -m pip install --user lxml

# Recommended when using images in report
python3.8 -m pip install --user pillow

# Recommended: install python-OBD from source (github repository)
git clone https://github.com/brendan-w/python-OBD.git
cd python-OBD
python3.8 -m build
python3.8 -m pip install --user dist/dist/obd-0.7.1-py3-none-any.whl
cd

# Recommended: install telemetry-obd from source (github repository)
git clone https://github.com/thatlarrypearson/telemetry-obd.git
cd telemetry-obd
python3.8 -m build
python3.8 -m pip install --user dist/telemetry_obd-0.2-py3-none-any.whl
cd

# Recommended: install sae-j1979da-reformatter from source (github repository)
git clone https://github.com/thatlarrypearson/sae-j1979da-reformatter.git
cd sae-j1979da-reformatter
python3.8 -m build
python3.8 -m pip install --user dist/j1979_reformatter-????-py3-non-any.whl
cd
```

## Limitations

- Most OBD commands have a table describing the data returned through an OBD interface.  The tables don't have borders but they should.
- The program is USA biased in that only USA regulatory terms are printed out.
  - Fix: replace ```'US OBD Regulatory term used'``` with another column heading.
  - To see alternative column headings, turn on verbose mode (```--verbose```).  A representation of the dictionary of dictionaries will be printed out.
- The program only reformats annex B and G.

## See Also

- [Telemetry OBD Data To CSV File](https://github.com/thatlarrypearson/telemetry-obd-log-to-csv)

## License

[MIT License](LICENSE)
