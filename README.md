# `excel2tmx`
<details>

  ## Python version
  
  A [Python script](excel2tmx.py) to convert XLS(X) files to TMX.  
  The script is intended to convert spreadsheets created by [`﻿﻿write_project2excel.groovy`](https://github.com/capstanlqc/omegat-scripts/blob/master/write_project2excel.groovy) in OmegaT, but can work with any Excel file if the column headers are in the row 2.

  ```bash
  usage: python3 excel2tmx.py [-h] --sl SL --tl TL [--sheet-pattern SHEET_PATTERN] [--alttype {id,context}] [--omt] <file_path>
  The following arguments are required: <file_path>, --sl, --tl
  ```
  | **Command line arguments:** | Explanation |
  |-----------------------------|-------------|
  |`--sl`: | source language code |
  |`--tl`: | target language code |
  |`<file_path>`: | path to the input Excel file | 
  | *Optional arguments:* |       |
  |`--sheet-patern`: | a regex to specify which sheet(s) to process |
  |`--alttype`: | argument to specify how alternative segments are identified and written to TMX (<prop type="id">, or <prop type="prev"> and <prop type="next">). Supported values: `id` and `context` (defaults to `id`). If segment ID is not found in the Excel file, the segment is treated as if `context` was specified, even with `alttype` set to `id`. |
  |`--omt`: | argument (without value) to control the output location. If set, the output is `../tm/excel2tmx`, otherwise `../excel2tmx_output` |

  **Python dependencies:**
  `pandas`, `xlrd`  
  ```bash
  pip install pandas xlrd
  ```

  ## Groovy version (for OmegaT)
  A [Groovy script](excel2tmx.groovy) to convert XLS(X) files to TMX in OmegaT.  
  The script doesn't take any command line arguments. Source and target languages, as well as the output location are defined by the current OmegaT project.  
  The script expects Excel files to be located in `<project>/script_input`, the resultant TMX files are placed into `<project>/tm/excel2tmx/`. The script processes all Excel files found in the input folder.  
  
  `alttype` is defined in the script itself, accepts two values: `id` and `context`.  
  `sheetPattern` is defined in the script itself, defaults to `~/.*/` (processes every sheet).  
  
  The rest of the functionality is identical to the Python version.  
  For the first run, the computer should be connected to the internet to download required libraries. Once the dependencies are downloaded, connection is not required.
</details>
