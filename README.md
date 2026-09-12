# Excel to AASX

A local tool for converting Excel assembly data into AAS 3.0 packages.

## Quick Start

1. Install **Python 3.12** on Windows.
2. Run `setup.bat` to install dependencies.
3. Run `scripts/start.bat` to open the web interface.
4. Upload your `.xlsx` file and adjust field mappings if needed.
5. Click **Export Draft**, then download the `.aasx` file.

## Notes

- Excel files must follow the [input format](docs/INPUT_FORMAT.md).
- Draft exports allow missing information and record it in the output.
- Exported packages include the original Excel file.
- Press **Ctrl+C** in the terminal to stop the application.
