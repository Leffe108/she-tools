# she-tools
Tools for [S:t Hans Extreme](https://she.lundsok.se) - a trail running race in Lund

# result.py
Requirements: Python 3.8 (no additional libraries)

Extracts a subset of the data from [IOF 3.0 XML](https://orienteering.sport/iof/it/data-standard-3-0/) ResultList from [UsynligO](https://usynligo.no/) app and saves it to CSV for Excel.

**Usage**:
* `py result.py UsynligO_result_file.xml` (creates a CSV file for Excel in languages using , as decimal separator)
* `py result.py --en UsynligO_result_file.xml` (creates a CSV file for English Excel)
