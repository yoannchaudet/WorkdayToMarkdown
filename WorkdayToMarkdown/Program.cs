using System.CommandLine;
using System.Text;
using ExcelDataReader;
using Microsoft.Extensions.Logging;

namespace WorkdayToMarkdown;

internal static class Program
{
    // CLI parsing
    private static async Task<int> Main(string[] args)
    {
        var fileOption = new Option<FileInfo?>(
            "--file",
            "Feedback file or folder")
        {
            IsRequired = true
        };
        var dateOption = new Option<DateOnly?>(
            "--since",
            "The date at which to start for collecting feedback");

        var rootCommand = new RootCommand("Convert a Workday feedback file export (XLSX) to a markdown file");
        rootCommand.AddOption(fileOption);
        rootCommand.AddOption(dateOption);

        rootCommand.SetHandler((file, date) =>
            {
                // Create logger
                using var factory = LoggerFactory.Create(builder => builder.AddConsole());
                var logger = factory.CreateLogger("Program");

                // Default to 6 months ago
                date ??= DateOnly.FromDateTime(DateTime.Now).AddMonths(-6);

                logger.LogInformation($"Cutoff date: {date}");

                // Read all feedback (either in a folder or in a file)
                IEnumerable<Feedback> feedback;
                if (Directory.Exists(file!.FullName))
                    feedback = Directory.GetFiles(file.FullName, "*.xlsx")
                        .SelectMany(path => ReadFeedbackFile(logger, new FileInfo(path)));
                else
                    feedback = ReadFeedbackFile(logger, file);

                // Filter by cutoff date, then group by receiver (skip receivers with no in-range feedback)
                var groups = feedback
                    .Where(f => DateOnly.FromDateTime(f.Date) >= date.Value)
                    .OrderBy(f => f.To)
                    .GroupBy(f => f.To)
                    .ToDictionary(g => g.Key!, g => g.ToList());

                // Output markdown
                var tempFile = Path.GetTempFileName();
                using var writer = new StreamWriter(tempFile, false, Encoding.UTF8);
                writer.WriteLine("# Peer feedback");
                writer.WriteLine();
                writer.WriteLine($"Generated on {DateTime.Now}");
                writer.WriteLine();
                foreach (var receiver in groups.Keys)
                {
                    writer.WriteLine($"## {receiver}");
                    writer.WriteLine();

                    string? from = null;
                    DateTime? on = null;

                    foreach (var currentFeedback in groups[receiver])
                    {
                        if (currentFeedback.From != from || currentFeedback.Date != on)
                        {
                            from = currentFeedback.From;
                            on = currentFeedback.Date;
                            writer.Write($"### {currentFeedback.From}, {currentFeedback.Date:yyyy-MM-dd}");
                            writer.WriteLine(currentFeedback.IsConfidential ? " 🔒" : "");
                            writer.WriteLine();
                        }

                        writer.WriteLine($"{currentFeedback.Question}");
                        writer.WriteLine(GetMarkdownQuote(currentFeedback.Response));
                        writer.WriteLine();
                    }
                }

                // Logging
                logger.LogInformation($"Markdown file written in {tempFile}");
            },
            fileOption, dateOption);
        return await rootCommand.InvokeAsync(args);
    }

    // Helper to quote a Markdown text
    private static string GetMarkdownQuote(string? text)
    {
        return text != null ? $"> {text!.Trim().Replace("\n", "\n> ")}" : "";
    }

    // Convert the feedback file into markdown
    private static IEnumerable<Feedback> ReadFeedbackFile(ILogger logger, FileInfo feedbackFile)
    {
        // Load old encoding for ExcelDataReader
        // See https://github.com/ExcelDataReader/ExcelDataReader?tab=readme-ov-file#important-note-on-net-core
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Logging
        logger.LogInformation($"Reading {feedbackFile.FullName}");

        // Go through the feedback file. Two schemas are supported:
        //   Legacy (10 cols): [0]=To, [2]=Date, [5]=From, [7]=Question, [8]=Response, [9]=Confidential
        //   Current (8 cols): [0]=Date, [1]=To, [2]=From, [5]=Question, [6]=Response (no confidential column)
        using var stream = feedbackFile.OpenRead();
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var expectedHeaders = 2;
        var line = 0;
        while (reader.Read())
        {
            line++;

            int dateCol, toCol, fromCol, questionCol, responseCol, confidentialCol;
            switch (reader.FieldCount)
            {
                case 10 when reader.GetFieldType(2) == typeof(DateTime):
                    (dateCol, toCol, fromCol, questionCol, responseCol, confidentialCol) = (2, 0, 5, 7, 8, 9);
                    break;
                case 8 when reader.GetFieldType(0) == typeof(DateTime):
                    (dateCol, toCol, fromCol, questionCol, responseCol, confidentialCol) = (0, 1, 2, 5, 6, -1);
                    break;
                default:
                    // Allow up to 2 bad schema rows (for the headers)
                    expectedHeaders--;
                    if (expectedHeaders < 0)
                    {
                        var rowValues = new List<string>();
                        for (var i = 0; i < reader.FieldCount; i++)
                            rowValues.Add(
                                $"{reader.GetValue(i)?.ToString() ?? string.Empty} (type: {reader.GetFieldType(i)?.ToString() ?? string.Empty})");
                        var rawLine = string.Join(",", rowValues);
                        logger.LogWarning($"Unexpected bad schema row at line {line}: {rawLine}");
                    }
                    continue;
            }

            // Extract the columns we care about
            var date = reader.GetDateTime(dateCol);
            var receiver = reader.GetString(toCol);
            var giver = reader.GetString(fromCol);
            var question = reader.GetString(questionCol);
            var response = reader.IsDBNull(responseCol) ? null : reader.GetString(responseCol);
            var confidential = confidentialCol >= 0
                && !reader.IsDBNull(confidentialCol)
                && reader.GetString(confidentialCol).ToLowerInvariant().Equals("yes");

            // Skip rows with no actual response text (e.g. unanswered optional questions)
            if (string.IsNullOrWhiteSpace(response))
                continue;

            // Return the feedback
            yield return new Feedback
            {
                Date = date,
                From = giver,
                To = receiver,
                Question = question,
                Response = response,
                IsConfidential = confidential
            };
        }
    }
}