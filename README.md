# RobotNotasDeExclusion

This repository contains an example of Robotic Process Automation (RPA) using Python. It demonstrates how to automate the process of generating Word documents from data extracted from an Excel file. The automation utilizes the `openpyxl` library for Excel manipulation and the `python-docx` library for Word document manipulation.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Project Structure](#project-structure)
- [Usage](#usage)
- [Customization](#customization)
- [Contributing](#contributing)
- [License](#license)

## Prerequisites

Before you begin, make sure you have the following installed:

- Python 3.x
- Required Python libraries: `openpyxl` and `python-docx`

You can install the required libraries using the following command:

```sh
pip install openpyxl python-docx
```

## Project Structure

The project is organized as follows:

```
my_rpa_project/
│
├── main.py
├── actions/
│   ├── __init__.py
│   └── automation_functions.py
│
├── input_data/
│   └── resumen exclusiones.xlsx
│
├── note_templates/
│   ├── template1.docx
│   ├── template2.docx
│   └── template3.docx
│
└── output_data/
```

- `main.py`: Main script to run the RPA automation process.
- `actions/`: Contains functions for data extraction and document generation.
- `input_data/`: Contains the input Excel file.
- `note_templates/`: Contains Word template files for document generation.
- `output_data/`: The generated Word documents will be saved here.

## Usage

1. Place your Excel file with data in the `input_data/` directory.
2. Create your Word template files and place them in the `note_templates/` directory.
3. Run the RPA automation by executing the `main.py` script:

```sh
python main.py
```

Generated Word documents will be saved in the `output_data/` directory.

## Customization

- Customize the `template_paths` list in `main.py` to match your template filenames.
- Adjust the placeholder names in your templates and functions as needed.
- Modify the functions in `automation_functions.py` to add more custom automation steps.

## Contributing

Contributions are welcome! If you find issues or have suggestions, please open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).
```

You can copy and paste this Markdown content into a `.md` file in your GitHub repository to create your README.
