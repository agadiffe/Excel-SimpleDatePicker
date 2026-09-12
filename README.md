<p align="center">
  <img src="img/Excel-SimpleDatePicker_header.png" alt="Simple DatePicker displayed in light and dark mode on an Excel worksheet" width="90%">
</p>

<div align="center">
  <h1>Excel - Simple DatePicker</h1>
A simple and lightweight DatePicker for Excel, built entirely in VBA.
</div>

## 📝 Features

- Opens automatically when double-clicking a cell formatted as a Date
- Supports light and dark modes with automatic detection
- Supports multiple languages based on Excel's display language
- Responsive layout with adjustable size

> **Resizing:** To change the DatePicker size, modify the `DP_DATEPICKER_WIDTH` constant in `DP_modLayout`.

## ⚙️ Installation

The DatePicker requires **Microsoft Excel with VBA support and macros enabled**.

### Import the VBA components

> **Note:** If Windows has blocked the downloaded files, unblock them before importing.

1. Open your Excel workbook and press `Alt + F11` to open the VBA editor.
2. In the **Project Explorer**, select your workbook's VBA project.
3. Import the components from the corresponding folders in `src`:
   - **Class Modules**: import the `.cls` files
   - **Forms**: import the `.frm` files (keep each `.frx` file in the same folder)
   - **Modules**: import the `.bas` files
4. Open the `ThisWorkbook` module and add the following code to the `Workbook_SheetBeforeDoubleClick` event:
    ```vba
    If DP_HandleDatePickerDoubleClick(Target) Then
        Cancel = True ' Prevent other Excel actions
    End If
    ```
5. Save your workbook as an `.xlsm` file.

### Using the included demo

A ready-to-use example is available in `demo/Excel-SimpleDatePicker.xlsm`.

## 🔒 Security

The demo workbook contains VBA macros, so Excel may display its usual security warning when you open it.

If Windows has blocked the downloaded file, Excel may prevent the macros from running without offering an option to enable them.  
To unblock it, right-click the file, select **Properties**, check **Unblock**, then confirm.

The complete VBA source code is available in this repository for review.

## 🎨 Theme

The DatePicker also supports Excel for Mac.  
For the optional automatic dark-mode detection, copy `DatePickerTheme.applescript` from `src/Mac/` to:

`~/Library/Application Scripts/com.microsoft.Excel/`

## 📌 Remarks

Supported formats include Excel's Short Date formats, the system's Long Date format, and several common Long Date formats not tied to the system locale.  
Time-only formats are not supported, and custom date formats may not trigger the DatePicker.

The double-click behavior and the cell formats that trigger the DatePicker can be customized in `DP_modHelpers`.

Excel Online includes a built-in date picker, but this feature is currently not available in the desktop version.

## 💙 Support

If you find a bug, open an issue.  
If you like the project, leave a ⭐.  
Thanks.
