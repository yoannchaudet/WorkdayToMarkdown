namespace WorkdayToMarkdown;

public class Feedback
{
    public string From { get; set; }
    public string To { get; set; }
    public DateTime Date { get; set; }
    public string Question { get; set; }
    public string FeedbackResponse { get; set; }
    public bool IsConfidential { get; set; }
}