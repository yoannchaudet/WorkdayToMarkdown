# WorkdayFeedbackToMarkdown

A very quick utility to streamline extracting peer feedback from Workday.

## Prerequisites

To compile/run that you need the .NET SDK. [Install it for your platform](https://dotnet.microsoft.com/en-us/download/dotnet/8.0) or if you are on MacOS and using Homebrew, run:

```bash
brew install --cask dotnet-sdk
```

## Usage

1. Connect to Workday and download all peer feedback you are interested in (as XLSX files)

   <img width="933" alt="Download feedback as an Excel file" src="https://github.com/yoannchaudet/WorkdayToMarkdown/assets/14911070/b11d2f50-c8f3-4777-8ef3-b38a7eca7cb3">

2. Then point the utility either at a single file or a folder containing all your files

   At the root of this repository:

    ```bash
    dotnet run --project WorkdayToMarkdown --file <path to a folder or an XLSX file>
    ```

    A temporary file will be created containing all the peer feedback  in a digestible way. It is meant to be imported in any tool that  understands Markdown your heart desire (Obsidian, GitHub, etc.).

## Note

This is mostly hardcoded for whatever my employer has setup.
