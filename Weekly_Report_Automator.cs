using DocumentFormat.OpenXml.Packaging;
using System.Text;

class Weekly_Report_Automator
{
    static DateTime GetFirstSaturday()
    {
        return DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek - 1);
    }

    static DateTime GetLastFriday()
    {
        return GetFirstSaturday().AddDays(13);
    }

    static DateTime GetFirstWeek()
    {
        return GetFirstSaturday().AddDays(-21);
    }

    static Dictionary<string, string> GenerateDayPlaceholders()
    {
        // Define Greek Day format
        string[] greekDays = { "ΚΥΡ", "ΔΕΥ", "ΤΡΙ", "ΤΕΤ", "ΠΕΜ", "ΠΑΡ", "ΣΑΒ" };

        Dictionary<string, string> placeholderDates = new Dictionary<string, string>();

        var firstSaturday = GetFirstSaturday();
        for (int i = 0; i < 14; i++)
        {
            DateTime dayOfWeek = firstSaturday.AddDays(i);
            placeholderDates[$"$D{i}"] = $"{greekDays[(int)dayOfWeek.DayOfWeek]}\n{dayOfWeek:dd MMM yyyy}".ToUpper();
        }
        return placeholderDates;
    }

    static Dictionary<string, string> GenerateWeekPlaceholders()
    {
        Dictionary<string, string> placeholdersWeek = new Dictionary<string, string>();

        var firstWeek = GetFirstWeek();
        for (int i = 0; i < 6; i++)
        {
            DateTime firstDayOfWeek = firstWeek.AddDays(i * 7);
            DateTime lastDayOfWeek = firstDayOfWeek.AddDays(6);
            placeholdersWeek[$"$W{i}"] = $"{firstDayOfWeek:dd/MM} - {lastDayOfWeek:dd/MM}";
        }
        return placeholdersWeek;
    }

    static Dictionary<string, string> GenerateFinalPlaceholders()
    {
        string[] greekDays = { "ΚΥΡ", "ΔΕΥ", "ΤΡΙ", "ΤΕΤ", "ΠΕΜ", "ΠΑΡ", "ΣΑΒ" };
        Dictionary<string, string> placeholdersFinal = new Dictionary<string, string>();

        var firstSaturday = GetFirstSaturday();
        var lastFriday = GetLastFriday();

        placeholdersFinal["$F0"] = $"{greekDays[(int)firstSaturday.DayOfWeek]} {firstSaturday:dd MMM yyyy}".ToUpper();
        placeholdersFinal["$F1"] = $"{greekDays[(int)lastFriday.DayOfWeek]} {lastFriday:dd MMM yyyy}".ToUpper();
        return placeholdersFinal;
    }
    static void ReplacePlaceholderInTextElement(DocumentFormat.OpenXml.Drawing.Text textElement, Dictionary<string, string> placeholders)
    {
        foreach (string placeholder in placeholders.Keys)
        {
            var words = textElement.Text.Split(new[] { ' ', '\t', '\n', '\r', '.', ',', '!', '?' }, StringSplitOptions.None);
            for (int i = 0; i < words.Length; i++)
            {
                if (words[i] == placeholder)
                {
                    words[i] = placeholders[placeholder];
                }
            }
            textElement.Text = string.Join(" ", words);
        }
    }

    static void ReplacePlaceholdersInPresentation(string presentationFilePath, Dictionary<string, string> placeholders)
    {
        using (PresentationDocument presentation = PresentationDocument.Open(presentationFilePath, true))
        {
            if (presentation.PresentationPart != null)
            {
                var slideParts = presentation.PresentationPart.SlideParts;

                foreach (var slidePart in slideParts)
                {
                    var textElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>();

                    foreach (var textElement in textElements)
                    {
                        ReplacePlaceholderInTextElement(textElement, placeholders);
                    }
                }
            }
        }
    }

    static void ProcessFile(string fileSource,string fileDestination, List<Dictionary<string, string>> placeholdersList)
    {
        // Ensure the destination directory exists
        Directory.CreateDirectory(Path.GetDirectoryName(fileDestination));

        // Copy the source file to the destination file before modifying
        File.Copy(fileSource, fileDestination, true);

        foreach (var placeholders in placeholdersList)
        {
            ReplacePlaceholdersInPresentation(fileDestination, placeholders);
        }
    }

    static void CopyOtherBureauPresentation(string otherBureausPresentationDestination)
    {
        string rootDirectory = Path.GetDirectoryName(otherBureausPresentationDestination);
        File.Copy(otherBureausPresentationDestination, rootDirectory+ "\\ΑΠΟΛΟΓΙΣΜΟΣ 3ο ΕΓ.pptx", true);
        File.Copy(otherBureausPresentationDestination, rootDirectory + "\\ΑΠΟΛΟΓΙΣΜΟΣ 4ο ΕΓ.pptx", true);

    }

    static void Main()
    {
        Console.OutputEncoding = Encoding.UTF8;

        string firstBureauPresentationTemplate = Path.Combine(AppContext.BaseDirectory, "C:\\Users\\Giorgos\\source\\repos\\Weeekly_Report_Automator\\template\\firstBureauTemplate.pptx");
        string otherBureausPresentationTemplate = Path.Combine(AppContext.BaseDirectory, "C:\\Users\\Giorgos\\source\\repos\\Weeekly_Report_Automator\\template\\otherBureausTemplate.pptx");
        string finalPresentationTemplate = Path.Combine(AppContext.BaseDirectory, "C:\\Users\\Giorgos\\source\\repos\\Weeekly_Report_Automator\\template\\finalPresentationTemplate.pptx");

        string firstBureauPresentationDestination = Path.Combine(AppContext.BaseDirectory, "C:\\Users\\Giorgos\\source\\repos\\Weeekly_Report_Automator\\output\\ΑΠΟΛΟΓΙΣΜΟΣ 1ο ΕΓ.pptx");
        string otherBureausPresentationDestination = Path.Combine(AppContext.BaseDirectory, "C:\\Users\\Giorgos\\source\\repos\\Weeekly_Report_Automator\\output\\ΑΠΟΛΟΓΙΣΜΟΣ 2ο ΕΓ.pptx");
        string finalPresentationDestination = Path.Combine(AppContext.BaseDirectory, "C:\\Users\\Giorgos\\source\\repos\\Weeekly_Report_Automator\\output\\ΑΠΟΛΟΓΙΣΜΟΣ ΤΕΛΙΚΟ.pptx");

        Dictionary<string, string> dayPlaceholders = GenerateDayPlaceholders();
        Dictionary<string, string> weekPlaceholders = GenerateWeekPlaceholders();
        Dictionary<string, string> finalPresentationPlaceholders = GenerateFinalPlaceholders();


        ProcessFile(firstBureauPresentationTemplate, firstBureauPresentationDestination, new List<Dictionary<string,string>> {dayPlaceholders, weekPlaceholders });

        ProcessFile(otherBureausPresentationTemplate, otherBureausPresentationDestination, new List<Dictionary<string, string>> { dayPlaceholders, weekPlaceholders });

        ProcessFile(finalPresentationTemplate, finalPresentationDestination, new List<Dictionary<string, string>>{finalPresentationPlaceholders});

        CopyOtherBureauPresentation(otherBureausPresentationDestination);
    }
}