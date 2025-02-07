using DocumentFormat.OpenXml.Packaging;
using System;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

class Program
{
    static DateTime getFirstSaturday()
    {
        return DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek - 1);
    }

    static DateTime getFirstWeek()
    {
        return getFirstSaturday().AddDays(-21);
    }

    static Dictionary<string, string> GenerateDayPlaceholders()
    {
        // Define Greek Day format
        string[] greekDays = { "ΚΥΡ", "ΔΕΥ", "ΤΡΙ", "ΤΕΤ", "ΠΕΜ", "ΠΑΡ", "ΣΑΒ" };

        Dictionary<string, string> placeholderDates = new Dictionary<string, string>();

        var firstSaturday = getFirstSaturday();
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

        var firstWeek = getFirstWeek();
        for (int i = 0; i < 6; i++)
        {
            DateTime firstDayOfWeek = firstWeek.AddDays(i * 7);
            DateTime lastDayOfWeek = firstDayOfWeek.AddDays(6);
            placeholdersWeek[$"$W{i}"] = $"{firstDayOfWeek:dd/MM} - {lastDayOfWeek:dd/MM}";
        }
        return placeholdersWeek;
    }

    static void Main()
    {
        Console.OutputEncoding = Encoding.UTF8;
        // Define Directories
        string sourceFilePath = Path.Combine(AppContext.BaseDirectory, "C:\\Users\\Giorgos\\source\\repos\\Weeekly_Report_Automator\\template\\template.pptx");
        string destinationFilePath = Path.Combine(AppContext.BaseDirectory, "C:\\Users\\Giorgos\\source\\repos\\Weeekly_Report_Automator\\output\\modified.pptx");

        // Ensure the destination directory exists
        Directory.CreateDirectory(Path.GetDirectoryName(destinationFilePath));

        // Copy the source file to the destination file before modifying
        File.Copy(sourceFilePath, destinationFilePath, true);

        Dictionary<string, string> dayPlaceholders = GenerateDayPlaceholders();
        Dictionary<string, string> weekPlaceholders = GenerateWeekPlaceholders();


        using (PresentationDocument presentation = PresentationDocument.Open(destinationFilePath, true))
        {
            if (presentation.PresentationPart != null)
            {
                var slideParts = presentation.PresentationPart.SlideParts;

                foreach (var slidePart in slideParts)
                {
                    var textElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>();

                    foreach (var textElement in textElements)
                    {
                        foreach (string placeholder in dayPlaceholders.Keys)
                        {
                            var words = textElement.Text.Split(new[] { ' ', '\t', '\n', '\r', '.', ',', '!', '?' }, StringSplitOptions.None);
                            for (int i = 0; i < words.Length; i++)
                            {
                                if (words[i] == placeholder)
                                {
                                    words[i] = dayPlaceholders[placeholder];
                                }
                            }
                            textElement.Text = string.Join(" ", words);
                        }
                    }
                }
            }
        }

        using (PresentationDocument presentation = PresentationDocument.Open(destinationFilePath, true))
        {
            if (presentation.PresentationPart != null)
            {
                var slideParts = presentation.PresentationPart.SlideParts;

                foreach (var slidePart in slideParts)
                {
                    var textElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>();

                    foreach (var textElement in textElements)
                    {
                        foreach (string placeholder in weekPlaceholders.Keys)
                        {
                            var words = textElement.Text.Split(new[] { ' ', '\t', '\n', '\r', '.', ',', '!', '?' }, StringSplitOptions.None);
                            for (int i = 0; i < words.Length; i++)
                            {
                                if (words[i] == placeholder)
                                {
                                    words[i] = weekPlaceholders[placeholder];
                                }
                            }
                            textElement.Text = string.Join(" ", words);
                        }
                    }
                }
            }
        }
    }
}