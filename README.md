
![Updated](https://img.shields.io/badge/updated-2025-11-17-blue)
# Piping Comparison Script

## Overview
The **Piping Comparison Script** is a Python-based tool designed to compare and match piping component descriptions between Excel sheets. It extracts and analyzes key attributes such as size, material, type, and other specifications to identify matching components based on predefined criteria. This tool is particularly useful for engineers and project managers working on inventory and piping data, enabling efficient matching and comparison.

---

## Features
- **Data Preprocessing:** Cleans and filters input data to remove invalid entries.
- **Attribute Extraction:** Extracts detailed attributes like size, type, material, SCH, and class using regex.
- **Custom Matching Logic:** Implements flexible criteria for matching components based on type-specific requirements.
- **Multi-Sheet Comparison:** Compares descriptions across multiple sheets and tracks remaining quantities.
- **Detailed Output:** Generates updated Excel files with consumed matches and remaining quantities.
- **Comprehensive Logging:** Outputs detailed logs for comparisons, highlighting matched and unmatched components.

---

## Requirements
Ensure the following dependencies are installed before running the script:

- `pandas`: For data manipulation.
- `openpyxl`: For handling Excel files.
- `xlrd`: For reading older Excel formats.
- `re`: For regex-based data extraction.
- `os`: For file path management.
- `json`: For loading mapping configurations.

Install dependencies using pip:

```bash
pip install pandas openpyxl xlrd
```

---

## Setup

### 1. Input File Requirements
- Place the input Excel file in the directory specified in the script.
- The Excel file should contain at least:
  - A sheet named `Sheet1` with columns:
    - `Description` (mandatory)
    - `QUANTITY` (mandatory)
  - Additional sheets with columns:
    - `Description` (mandatory)
    - `CURRENT` (mandatory for inventory tracking)

### 2. Configuration File
- A `SCH.txt` file containing size mappings must be located in the same directory as the script.

Example JSON format for `SCH.txt`:
```json
{
    "0.125": {
        "10": 1.24,
        "30": 1.45,
        "STD": 1.73,
        "40": 1.73,
        "XS": 2.41,
        "80": 2.41,
        "XXS": 1.73,
        "10S": 1.24,
        "40S": 1.73,
        "80S": 2.41
    },
    "0.25": {
        "10": 1.65,
        "30": 1.85,
        "STD": 2.24,
        "40": 2.24,
        "XS": 3.02,
        "80": 3.02,
        "10S": 1.65,
        "40S": 2.24,
        "80S": 3.02
    }}

```

---

## Usage

### Running the Script
1. Ensure the input Excel file and `SCH.txt` are in place.
2. Run the script using:

```bash
python check_inventory_description_matcher_with-remaning.py
```

### Methods
The script executes two primary methods:
1. **`display_comparison()`**:
   - Outputs comparison results in the console.
   - Saves a detailed log to `comparison_output.txt`.

2. **`run_comparison()`**:
   - Updates `Sheet1` with matched descriptions and remaining quantities.
   - Saves results to an output file, e.g., `out.xlsx`.
   - Updates inventory sheets with adjusted `CURRENT` values.

---

## File Outputs
- **`out.xlsx`:** Updated main sheet with columns:
  - `Consumed Matches`: Details of matched inventory items.
  - `Remaining Quantity`: Remaining unmatched quantity.
- **Updated Inventory Files:**
  - Separate updated sheets for each inventory, saved with the prefix `updated_`.
- **`comparison_output.txt`:** Log file with detailed comparison results.

---

## Customization
### Adjust Matching Criteria
Modify the `compare_descriptions` method to customize matching logic. Criteria include:
- Sizes, SCH, and class tolerances.
- Material and connection type matching.

### Preprocessing Rules
Update `_preprocess_sheets` to:
- Exclude specific descriptions.
- Adjust filtering logic for invalid entries.

---

## Example Input & Output

### Input
| Description             | QUANTITY |
|-------------------------|----------|
|    12 IN NS ELBOW A234-WPB 90-LR SMLS BW TO ASME B16.9 SCH20  | 10       |

### Output (in `out.xlsx`)
| Description             | QUANTITY | Consumed Matches                | Remaining Quantity |
|-------------------------|----------|---------------------------------|--------------------|
|    12 IN NS ELBOW A234-WPB 90-LR SMLS BW TO ASME B16.9 SCH20  | 10       | InventorySheet1: 5; Sheet2: 3   | 2                  |

---

## Troubleshooting
- **Missing `SCH.txt` File:** Ensure the size map file exists and contains valid JSON.
- **Incorrect File Paths:** Verify the input file paths match those in the script.
- **Unmatched Descriptions:** Check criteria in `compare_descriptions` for overly strict conditions.

---

## Contribution
Contributions are welcome! Submit issues or pull requests on GitHub to suggest enhancements or report bugs.

---

## License
This project is licensed under the MIT License. See the LICENSE file for details.

---

## Author
Hossein Izadi

**Support Email:** arkittoe@gmail.com


## Recent Updates

- Enhanced functionality
- Improved documentation
- Bug fixes

## Contributing

Contributions are welcome! Please read our contributing guidelines.

---

*Last updated: 2025-11-24 08:21:20*
