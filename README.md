# XLSXtoLabels
Script to convert MS Excel Spreadsheets to labels using LaTeX and Perl.

## Dependencies

### General
- pdflatex
- Perl (I used version 5.28)

### Perl modules
- `Spreadsheet::ParseXLSX`
- `Getopt::Long`
- `Pod::Usage`
- `Term::ANSIColor`
- `Config::IniFiles`

### LaTeX packages
- `extarticle`
- `inputenc`
- `labels`

## Usage
1. Create an Excel spreadsheet as shown in the Examples folder
2. Run the script by using the command `./script.pl`. The usage information with availible options will be shown when no arguments are given.
3. The output are two files, one is the TeX-file and the other one is the PDF-file that is compiled from the source code of the TeX-file.

## Predefined dimensions
The file `config.ini` is used to predefine the dimensions of the labels. At this moment, the only labels that are predefined are from the package `Avery J8163`. This can be used by passing the argument `--type J8163` when calling the script. `J8163` corresponds to the section name in `config.ini`.

Feel free to contribute and add more dimensions of other packages to the `config.ini` file.