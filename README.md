<p align="center">
  <img src="img/Excel-SimpleDatePicker_header.png" alt="Simple DatePicker displayed in light and dark mode on an Excel worksheet" width="90%">
</p>

<div align="center">
  <h1>Excel - Simple DatePicker</h1>
A simple and lightweight DatePicker for Excel, built entirely in VBA.
</div>

## 📝 Features

- Opens automatically when double-clicking a cell formatted as **Short Date**
- Fully responsive and automatically scales to different sizes
- Supports **light and dark mode**
- Supports multiple UI languages based on Excel's interface language
- Allows easy **day, month, and year selection**

> **Resizing:** To change the DatePicker size, modify the `DATEPICKER_WIDTH` constant in `DP_modLayout`. The layout, controls, and font sizes scale automatically.

## 🔒 Security

The `.xlsm` file included in this repository is a working example of the VBA DatePicker.

Because it is a macro-enabled Excel file, Excel may display a security warning when you open it.  
This is normal for files containing VBA macros.

The complete VBA source code is provided in this repository for anyone who wants to inspect it.

## ⚙️ Installation

The DatePicker requires **Microsoft Excel with VBA support**. Macros must be enabled to use it.

The VBA source is provided as exported VBA components. **Importing the components is recommended over copying and pasting the code**, as this preserves the required module, class, and UserForm structure.

### Import the VBA components

1. Open your Excel workbook and press `Alt + F11` to open the VBA editor.
2. In the **Project Explorer**, select your workbook's VBA project.
3. Import the components from the corresponding folders in `src`:
   - **Class Modules** → import the `.cls` files
   - **Forms** → import the `.frm` files
   - **Modules** → import the `.bas` files
4. When importing a UserForm, keep its corresponding `.frx` file in the same folder as the `.frm` file.
5. From `Microsoft_Excel_Objects`, copy the provided code into the corresponding Excel objects in your workbook, such as `ThisWorkbook`.
6. Save your workbook as an `.xlsm` file.
7. Enable macros when prompted.

> **Important:** Do not rename the imported VBA components unless you also update the code accordingly.

The `ThisWorkbook` code is required for the DatePicker to open when double-clicking a cell formatted as **Short Date**.

### Using the included demo

A ready-to-use example is available in [`demo/Excel-SimpleDatePicker.xlsm`](demo/Excel-SimpleDatePicker.xlsm).

## 🖥️ Compatibility

The DatePicker is designed for desktop versions of Microsoft Excel with VBA support.  
Older Excel versions may be compatible, but have not been tested against every release.

### macOS

The DatePicker supports Excel for Mac.  
For automatic dark-mode detection, copy `DatePickerTheme.applescript` from `src/Mac/` to:

`~/Library/Application Scripts/com.microsoft.Excel/`

Create the `com.microsoft.Excel` folder if it does not already exist.  
If the script is unavailable or cannot be executed, the DatePicker safely falls back to Light Mode.


## 📌 Remarks

The DatePicker is designed to work with cells formatted as **Short Date**.  
**Long Date** formatting is not supported.

Excel Online includes a built-in date picker, but this feature is currently not available in the desktop version of Excel.

## 💙 Support

If you find a bug, open an issue.  
If you like this small project, leave a ⭐.

Thanks!
