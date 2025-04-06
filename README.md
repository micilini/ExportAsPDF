
# ExportAsPDF

## Overview

`ExportAsPDF` is a Python script designed to convert structured JSON data into a well-formatted PDF document. It supports advanced formatting, such as styled text, hyperlinks, images, and more. The script uses [ReportLab](https://www.reportlab.com/) to make conversion from JSON to PDF directly.

`ExportAsPDF` was created as a module for the [Minimal Text Editor (Lite)](https://github.com/micilini/MinimalTextEditorLite) application.

---

## Features

- Converts JSON blocks into a PDF (`.pdf`).
- Supports:
  - Headers, paragraphs, and lists (ordered and unordered).
  - Quotes, warnings, and checklists.
  - Tables, images, and code blocks.
  - Inline formatting (bold, italic, underline, links, highlights).
- Automatically resizes images to fit the page.
- Use ReportLab to generate PDF files.

---

## Requirements

### Python Libraries

Install the required Python libraries using the following command:

```bash
pip install reportlab pillow pyinstaller
```

---

## How to Generate the Executable (`ExportAsPDF.exe`)

### Requirements
- Python 3.6 or higher
- `pip` package manager
- The following Python packages:
  - `reportlab`
  - `pillow`
  - `pyinstaller`

> 💡 `reportlab` is used for PDF generation. `pillow` handles image rendering (e.g. embedded base64 or SVG-converted PNGs).

### Create the Executable

Use `pyinstaller` to package the script into an executable:

   ```bash
   pyinstaller --onefile --distpath ./dist --name ExportAsPDF --add-data "assets;assets" ExportAsPDF.py
   ```
   - The `--onefile` flag ensures the executable is a single file.
   - The `--name` flag specifies the output executable's name.
   - The `--add-data` flag ensure that all files inside `assets` folder are inside the executable

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

## Example Usage

### Using the Executable in C#
Below is an example of how to call `ExportAsPDF.exe` from a C# application:

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

            return memoryStream.ToArray(); // Returns the binary PDF file
        }
    }
});
```

---

## Contributing

Feel free to open issues or submit pull requests to improve the script.

---

## License

This script is open-source and available under the MIT License.
