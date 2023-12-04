# JsonToExcel_pyton
program for extracting json files and write the to excel

# Usage
This program is desiged for converting a json file obtaind from ABB (edge-insight) to extract TsReffrences and tags to a exel file.

1. open a CMD/Terminal window
1. navigating to the repository paht (eks. "C:/JsonToExcel_pyton")
2. make sure python is installed
3. type:
   ```bash
   C:\JsonToExcel_pyton> python JsonToExcel_pyton.py <json-file-name> <excel-file-name>
   ```
   eks.
   ```bash
   C:\JsonToExcel_pyton> python JsonToExcel_pyton.py AC800.json "VeasTagTsRef.xlsx"
   ```
4. The script will return a excel-file containing your tags and ts-refs in seperated collumns.

├└── another_folder
    ├── file6.txt ├
    └── file7.txt
