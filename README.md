**N43-to-Excel**

This software is a PHP-based parser designed to convert Norma 43 (N43) files to Excel format. It simplifies the process of converting financial data from N43 files into a more accessible Excel format, aiding in data analysis and manipulation.

### Usage:

To utilize this software, follow these steps:

1. Ensure you have PHP installed on your system.
2. Clone this repository to your local machine or download the source code.
3. Open your command line interface.
4. Navigate to the directory containing the `n43_to_excel.php` file.
5. Execute the following command:

```
php n43_to_excel.php <input_n43_file> <output_excel_file>
```

Replace `<input_n43_file>` with the path to your N43 file and `<output_excel_file>` with the desired path for the generated Excel file.

### Example:

```
php n43_to_excel.php input.n43 output.xlsx
```

This command will convert the `input.n43` file to `output.xlsx`.

### Requirements:

- PHP installed on your system.
- Execute `composer install` to install the https://github.com/PHPOffice/PhpSpreadsheet library.

### License:

This software is distributed under the [MIT License](LICENSE).

### Contributions:

Contributions are welcome! If you find any bugs or have suggestions for improvements, feel free to open an issue or submit a pull request.

### Disclaimer:

This software is provided as-is without any warranties. Use at your own risk.

For more information or support, contact the-sir-arratoiak: https://github.com/the-sir-arratoiak
