# AutoInsertWordHeader

A VBA macro for Microsoft Word that automatically inserts your name and the current date into the document header.

## Table of Contents

- [Introduction](#introduction)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Customization](#customization)
- [Contributing](#contributing)
- [License](#license)

## Introduction

Manually adding your name and the current date to the header of every Word document can be repetitive. The `AutoInsertWordHeader` macro streamlines this process by automating the insertion, enhancing your workflow efficiency.

## Features

- **Automatic Insertion**: Adds your name and the current date to the header with a single command.
- **Customisable**: Easily modify the macro to fit your formatting preferences.
- **Optional Execution**: Assign a keyboard shortcut to run the macro when needed, ensuring flexibility.

## Installation

To integrate the `AutoInsertWordHeader` macro into Microsoft Word:

1. **Access the Developer Tab**:
   - Open Word.
   - Navigate to `File` > `Options` > `Customise Ribbon`.
   - Check the `Developer` option to enable the Developer tab.

2. **Add the Macro**:
   - Go to the `Developer` tab.
   - Click on `Visual Basic` to open the VBA editor.
   - In the editor, insert a new module:
     - Right-click on any existing module or the project tree.
     - Select `Insert` > `Module`.
   - Copy and paste the macro code from the [AutoInsertWordHeader repository](https://github.com/Gabrielhj17/AutoInsertWordHeader/blob/main/code/AutoInsertHeader.bas) into the module.
   - Save and close the VBA editor.

For a detailed guide on creating and running macros, refer to [Microsoft's official documentation](https://support.microsoft.com/en-us/office/create-or-run-a-macro-c6b99036-905c-49a6-818a-dfb98b7c3c9c).

## Usage

Once the macro is installed:

1. **Assign a Keyboard Shortcut** (Optional but recommended):
   - Navigate to `File` > `Options` > `Customise Ribbon`.
   - Click on `Customise` next to `Keyboard shortcuts`.
   - In the `Categories` list, select `Macros`.
   - In the `Macros` list, choose `AutoInsertHeader` (or the name you've given the macro).
   - Press the desired keyboard shortcut (e.g., `Ctrl` + `Shift` + `H`).
   - Click `Assign`, then `Close`.

2. **Run the Macro**:
   - Open any Word document.
   - Use the assigned keyboard shortcut or run the macro from the `Developer` tab.
   - Your name and the current date will be inserted into the header.

## Customisation

The macro is designed for flexibility:

- **Name**: By default, the macro uses a placeholder for the name. Replace `'Your Name'` in the code with your actual name.
- **Date Format**: Adjust the date format by modifying the `Format(Date, "MMMM d, yyyy")` function in the macro code to your preferred format.
- **Alignment**: The current alignment settings may need adjustment based on your template. Modify the alignment properties in the macro to suit your needs.

Feel free to enhance the macro further. Contributions are welcome!

## Contributing

Contributions to improve the macro are appreciated. To contribute:

1. Fork the repository.
2. Create a new branch:
   ```bash
   git checkout -b feature/YourFeature
   ```
3. Commit your changes:
   ```bash
   git commit -m "Add YourFeature"
   ```
4. Push to the branch:
   ```bash
   git push origin feature/YourFeature
   ```
5. Open a Pull Request.

## License

This project is licensed under the MIT License. See the [LICENSE](https://github.com/Gabrielhj17/AutoInsertWordHeader/blob/main/LICENSE) file for details.

