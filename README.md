
# ExportAsPDF

## Overview

`ExportAsPDF` is a Python script designed to convert structured JSON data into a well-formatted PDF document. It supports advanced formatting, such as styled text, hyperlinks, images, and more. The script detects the installed office suite (Microsoft Office or LibreOffice) to perform the conversion from DOCX to PDF.

`ExportAsPDF` was created as a module for the [Minimal Text Editor (Lite)](https://github.com/micilini/MinimalTextEditorLite) application.

---

## Features

- Converts JSON blocks into a Word document (`.docx`) and then into a PDF.
- Supports:
  - Headers, paragraphs, and lists (ordered and unordered).
  - Quotes, warnings, and checklists.
  - Tables, images, and code blocks.
  - Inline formatting (bold, italic, underline, links, highlights).
- Automatically resizes images to fit the page.
- Uses either Microsoft Office or LibreOffice for DOCX-to-PDF conversion.

---

## Requirements

### Python Libraries

Install the required Python libraries using the following command:

```bash
pip install python-docx docx2pdf pillow
```

### System Requirements

- **Microsoft Office** or **LibreOffice** installed on the system.
- Python 3.6 or higher.

---

## Usage

### Important Notes
- This script is designed to be called programmatically by other applications. It should not be opened directly by the user.
- The script accepts a single argument: the file path of the JSON input.

### Command-Line Execution

1. Save the JSON data to a local file.
2. Pass the file path as an argument when invoking the executable.
3. Capture the generated `.pdf` file output from the standard output.

---

---

## Example Usage

### Using the Executable in C#
Below is an example of how to call `ExportAsDoc.exe` from a C# application:

```csharp
// Save JSON data to a temporary file
string tempJsonFilePath = Path.Combine(Path.GetTempPath(), "data.json");
File.WriteAllText(tempJsonFilePath, jsonData);

// Execute the process asynchronously
var result = await Task.Run(() =>
{
    var processStartInfo = new ProcessStartInfo
    {
        FileName = "Modules\\Export\\ExportAsPDF.exe",
        Arguments = $"\"{tempJsonFilePath}\"",
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        UseShellExecute = false,
        CreateNoWindow = true
    };

    using (var process = new Process { StartInfo = processStartInfo })
    {
        process.Start();

        using (var memoryStream = new MemoryStream())
        {
            process.StandardOutput.BaseStream.CopyTo(memoryStream);
            process.WaitForExit();

            if (process.ExitCode != 0)
            {
                var error = process.StandardError.ReadToEnd();
                throw new Exception(error);
            }

            return memoryStream.ToArray(); // Returns the binary DOCX file
        }
    }
});
```

## Conversion Details

### Office Suite Detection

The script automatically detects and uses an installed office suite:
- **Microsoft Office**: Conversion is handled via `docx2pdf`.
- **LibreOffice**: Conversion is handled via a subprocess call to LibreOffice's CLI.

If no office suite is detected, the script will terminate with an error.

---

## Troubleshooting

- **Error**: `LibreOffice is not installed or not found in the specified paths.`
  - Ensure LibreOffice or Microsoft Office is installed and accessible from the system PATH.

- **Error**: `pip install command fails.`
  - Ensure Python and pip are correctly installed.

---

## Contributing

Feel free to open issues or submit pull requests to improve the script.

---

## License

This script is open-source and available under the MIT License.
