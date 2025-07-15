# D365 F&O Label Extractor & Manager

A powerful PowerShell script to automate the management of labels in Microsoft Dynamics 365 Finance & Operations projects.

This tool goes beyond simple extraction, providing features to automatically create labels from hardcoded text, diagnose missing label definitions, and maintain consistency across your codebase.

## Table of Contents
- [Why Use This Script?](#why-use-this-script)
- [Key Features](#key-features)
- [Requirements](#requirements)
- [Installation](#installation)
- [How to Use](#how-to-use)
  - [Parameters](#parameters)
  - [Command Examples](#command-examples)
- [Workflows (Use Cases)](#workflows-use-cases)
  - [1. Automating Initial Label Creation](#1-automating-initial-label-creation)
  - [2. Auditing and Maintaining an Existing Project](#2-auditing-and-maintaining-an-existing-project)
- [How to Contribute](#how-to-contribute)
- [License](#license)

## Why Use This Script?

Label management in D365 F&O projects can be a manual, repetitive, and error-prone process:
- Creating and referencing labels for every text property is time-consuming.
- It's easy to forget to create a label definition in the text file, causing "orphan labels" (`@STC:MyLabel`) that don't appear in the UI.
- Hardcoded text can spread throughout the project, making translation and maintenance difficult.
- Finding all the places where a label is used can be challenging.

This script is designed to solve all these problems by automating the most tedious tasks and providing powerful diagnostic tools.

## Key Features

- **Label Extraction**: Reads entire projects (`.rnrproj`) or solutions (`.sln`) and extracts all label references (`@STC:...`).
- **Automatic Label Creation**: Converts hardcoded text found in properties (`<Label>`, `<HelpText>`, etc.) into label references, generating the corresponding ID and value in your label file.
- **Clean ID Generation**: Creates automatic label IDs in `PascalCase` format (e.g., `MyObjectNamePropertyName`), free of special characters.
- **Duplicate Prevention**: Before creating a new label, it checks if a label with the same *value* or *ID* already exists, reusing it to keep the label file clean.
- **Orphan Label Diagnosis**: Identifies and reports which labels are referenced in the code but lack a definition in the label file.
- **Flexible Usage**: Allows specifying the full project path or just the name, using a default path for automatic resolution.
- **Versatile Output**: Displays results in the terminal or exports them to a report file.

## Requirements

- **PowerShell 5.1** or higher.
- Read/write access to your **Project/Solution** directory and the **Model Path** (`PackagesLocalDirectory`).

## Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/andrrff/d365utils.git
    ```
2.  Or, simply download the `Extract-D365Labels.ps1` file to your computer.

## How to Use

Open a PowerShell terminal, navigate to the directory where the script is located, and run it with the required parameters.

### Parameters

| Parameter | Type | Description | Required |
| :--- | :--- | :--- | :--- |
| `-ProjectPath` | `[string]` | Path to the `.sln` or `.rnrproj` file, or just the project name (will be resolved using `DefaultProjectsPath`). | Yes |
| `-ModelPath` | `[string]` | Path to the root of your model (e.g., `K:\AosService\PackagesLocalDirectory\STC\STC`). | Yes |
| `-CompareLabelFile` | `[string]` | Path to the label file (`.txt`) to be read and/or updated. Required for `-AutoCreateLabels` and `-FindMissingLabels`. | No |
| `-ShowInTerminal` | `[switch]` | If present, displays the found labels directly in the terminal. | No |
| `-OutputFile` | `[string]` | The file path where the list of extracted labels will be saved. | No |
| `-AutoCreateLabels` | `[switch]` | Activates the mode that converts hardcoded text into labels and updates the label file. Requires `-CompareLabelFile`. | No |
| `-FindMissingLabels` | `[switch]` | Activates the diagnostic mode that checks for label references without a definition in the label file. Requires `-CompareLabelFile`. | No |
| `-ShowDebug` | `[switch]` | Displays detailed debug information during execution. | No |
| `-DefaultProjectsPath` | `[string]` | The default path to your projects folder. Used when a simple name is passed to `-ProjectPath`. | No |

### Command Examples

> **Note:** All examples use the project name `STCTestLabelsHelper` and the label file path `K:\AosService\PackagesLocalDirectory\STC\STC\AxLabelFile\LabelResources\en-us\STC.en-us.label.txt`. Please adapt them to your environment.

**1. Simple Extraction**
Extracts all existing labels and displays them in the terminal.
```powershell
.\Extract-D365Labels.ps1 -ProjectPath "STCTestLabelsHelper" -ModelPath "K:\..." -ShowInTerminal
```

**2. Full Workflow (Recommended)**
Converts hardcoded text, finds orphan labels, and displays the results.
```powershell
.\Extract-D365Labels.ps1 `
    -ProjectPath "STCTestLabelsHelper" `
    -ModelPath "K:\AosService\PackagesLocalDirectory\STC\STC" `
    -CompareLabelFile "K:\...\STC.en-us.label.txt" `
    -AutoCreateLabels `
    -FindMissingLabels `
    -ShowInTerminal
```

## Workflows (Use Cases)

### 1. Automating Initial Label Creation

You've just created a new object and added text directly to its properties.

**Situation (`MyNewForm.xml` file):**
```xml
...
<HelpText>This is a help text.</HelpText>
<Label>My New Form</Label>
...
```

**Command:**
```powershell
.\Extract-D365Labels.ps1 -ProjectPath "MyProject" -CompareLabelFile "K:\...\STC.en-us.label.txt" -AutoCreateLabels
```

**Result:**
1.  The script will display in the console:
    ```
    [AUTO-LABEL] Text 'This is a help text.' converted to new label: @STC:MyNewFormHelpTextForm
    [AUTO-LABEL] Text 'My New Form' converted to new label: @STC:MyNewFormLabelForm
    ```
2.  The `MyNewForm.xml` file will be **automatically modified** to:
    ```xml
    ...
    <HelpText>@STC:MyNewFormHelpTextForm</HelpText>
    <Label>@STC:MyNewFormLabelForm</Label>
    ...
    ```
3.  The `STC.en-us.label.txt` file will be **updated** with the new lines:
    ```
    MyNewFormHelpTextForm=This is a help text.
    MyNewFormLabelForm=My New Form
    ```

### 2. Auditing and Maintaining an Existing Project

You suspect your project has labels referenced in the code that were forgotten in the label file.

**Situation (`ExistingTable.xml` file):**
```xml
...
<Label>@STC:CustomerName</Label> <!-- This label does not exist in the .txt file -->
...
```

**Command:**
```powershell
.\Extract-D365Labels.ps1 -ProjectPath "MyProject" -CompareLabelFile "K:\...\STC.en-us.label.txt" -FindMissingLabels
```

**Result:**
1.  The script will display a clear diagnostic report:
    ```
    [WARNING] 1 LABEL NOT FOUND IN THE LABEL FILE:
    ==========================================================

    - Missing Label: @STC:CustomerName
      Referenced in:
        - Object: ExistingTable (AxTable)
          File: K:\AosService\PackagesLocalDirectory\STC\STC\AxTable\ExistingTable.xml
    ```
With this information, you can confidently add the `CustomerName` label to your `.txt` file.

## How to Contribute

Contributions are welcome! If you have an idea for a new feature or find a bug:
1.  **Fork** this repository.
2.  Create a new branch (`git checkout -b feature/my-feature`).
3.  Make your changes and commit them (`git commit -am 'Add new feature X'`).
4.  Push to your branch (`git push origin feature/my-feature`).
5.  Create a new **Pull Request**.

## License

This project is licensed under the [MIT License](LICENSE.txt).
