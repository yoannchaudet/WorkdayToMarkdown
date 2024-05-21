# WorkdayFeedbackToMarkdown

A very quick utility to streamline extracting peer feedback from Workday.

## Usage

1. Connect to Workday and download all peer feedback you are interested in (as XLSX files)
2. Then point the utility either at a single file or a folder containing all your files

   At the root of this repository:

    ```bash
    dotnet run --project WorkdayToMarkdown --file <path to a  folder or an XLSX file>
    ```

    A temporary file will be created containing all the peer feedback  in a digestible way. It is meant to be imported in any tool that  understands Markdown your heart desire (Obsidian, GitHub, etc.).

## Note

This is mostly hardcoded for whatever my employer has setup.