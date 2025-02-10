using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using System.Text;

public static class Weekly_Report_Automator
{
    private const string GreekDays = "ΚΥΡ,ΔΕΥ,ΤΡΙ,ΤΕΤ,ΠΕΜ,ΠΑΡ,ΣΑΒ";
    private static readonly string[] GreekDayArray = GreekDays.Split(',');
    private static CultureInfo GreekCulture = new("el-GR");

    public static DateTime GetFirstSaturday() => SystemTime.Today().AddDays(-(int)SystemTime.Today().DayOfWeek - 1);
    public static DateTime GetLastFriday() => GetFirstSaturday().AddDays(13);
    public static DateTime GetFirstWeek() => GetFirstSaturday().AddDays(-21);

    public static Dictionary<string, string> GenerateDayPlaceholders()
    {
        var placeholderDates = new Dictionary<string, string>();
        var firstSaturday = GetFirstSaturday();

        for (int i = 0; i < 14; i++)
        {
            var dayOfWeek = firstSaturday.AddDays(i);
            var greekDayOfWeek = dayOfWeek.ToString("dd MMM yyyy", GreekCulture);
            placeholderDates[$"$D{i}"] = $"{GreekDayArray[(int)dayOfWeek.DayOfWeek]}\n{greekDayOfWeek}".ToUpper();
        }

        return placeholderDates;
    }

    public static Dictionary<string, string> GenerateWeekPlaceholders()
    {
        var placeholdersWeek = new Dictionary<string, string>();
        var firstWeek = GetFirstWeek();

        for (int i = 0; i < 6; i++)
        {
            var firstDayOfWeek = firstWeek.AddDays(i * 7);
            var lastDayOfWeek = firstDayOfWeek.AddDays(6);

            var greekFirstDayOfWeek = firstDayOfWeek.ToString("dd/MM", GreekCulture);
            var greekLastDayOfWeek = lastDayOfWeek.ToString("dd/MM", GreekCulture);

            placeholdersWeek[$"$W{i}"] = $"{greekFirstDayOfWeek} - {greekLastDayOfWeek}";
        }

        return placeholdersWeek;
    }

    public static Dictionary<string, string> GenerateFinalPlaceholders()
    {
        var placeholdersFinal = new Dictionary<string, string>();
        var firstSaturday = GetFirstSaturday();
        var lastFriday = GetLastFriday();

        placeholdersFinal["$F0"] = $"{GreekDayArray[(int)firstSaturday.DayOfWeek]} {firstSaturday:dd MMM yyyy}".ToUpper();
        placeholdersFinal["$F1"] = $"{GreekDayArray[(int)lastFriday.DayOfWeek]} {lastFriday:dd MMM yyyy}".ToUpper();

        return placeholdersFinal;
    }

    public static void ReplacePlaceholderInTextElement(DocumentFormat.OpenXml.Drawing.Text textElement, Dictionary<string, string> placeholders)
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

    public static void ReplacePlaceholdersInPresentation(string presentationFilePath, Dictionary<string, string> placeholders)
    {
        try
        {
            using (var presentation = PresentationDocument.Open(presentationFilePath, true))
            {
                if (presentation.PresentationPart == null)
                {
                    Console.WriteLine($"Error processing presentation: {presentationFilePath} does not contain a presentation part.");
                    return;
                }
                foreach (var slidePart in presentation.PresentationPart.SlideParts)
                {
                    var textElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>();
                    foreach (var textElement in textElements)
                    {
                        ReplacePlaceholderInTextElement(textElement, placeholders);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing presentation: {ex.Message}");
        }
    }

    public static void ProcessFile(string fileSource, string fileDestination, List<Dictionary<string, string>> placeholdersList)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(fileDestination));
            File.Copy(fileSource, fileDestination, true);

            foreach (var placeholders in placeholdersList)
            {
                ReplacePlaceholdersInPresentation(fileDestination, placeholders);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing file: {ex.Message}");
        }
    }

    public static void CopyOtherBureauPresentation(string otherBureausPresentationDestination)
    {
        try
        {
            var rootDirectory = Path.GetDirectoryName(otherBureausPresentationDestination);
            File.Copy(otherBureausPresentationDestination, Path.Combine(rootDirectory, "ΑΠΟΛΟΓΙΣΜΟΣ 3ο ΕΓ.pptx"), true);
            File.Copy(otherBureausPresentationDestination, Path.Combine(rootDirectory, "ΑΠΟΛΟΓΙΣΜΟΣ 4ο ΕΓ.pptx"), true);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error copying bureau presentations: {ex.Message}");
        }
    }

    public static void Main()
    {
        Console.OutputEncoding = Encoding.UTF8;

        var baseDirectory = AppContext.BaseDirectory;
        var templateDirectory = Path.Combine(baseDirectory, "template");
        var outputDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "WeeklyReports");

        var firstBureauPresentationTemplate = Path.Combine(templateDirectory, "firstBureauTemplate.pptx");
        var otherBureausPresentationTemplate = Path.Combine(templateDirectory, "otherBureausTemplate.pptx");
        var finalPresentationTemplate = Path.Combine(templateDirectory, "finalPresentationTemplate.pptx");

        var firstBureauPresentationDestination = Path.Combine(outputDirectory, "ΑΠΟΛΟΓΙΣΜΟΣ 1ο ΕΓ.pptx");
        var otherBureausPresentationDestination = Path.Combine(outputDirectory, "ΑΠΟΛΟΓΙΣΜΟΣ 2ο ΕΓ.pptx");
        var finalPresentationDestination = Path.Combine(outputDirectory, "ΑΠΟΛΟΓΙΣΜΟΣ ΤΕΛΙΚΟ.pptx");

        var dayPlaceholders = GenerateDayPlaceholders();
        var weekPlaceholders = GenerateWeekPlaceholders();
        var finalPresentationPlaceholders = GenerateFinalPlaceholders();

        ProcessFile(firstBureauPresentationTemplate, firstBureauPresentationDestination, new List<Dictionary<string, string>> { dayPlaceholders, weekPlaceholders });
        ProcessFile(otherBureausPresentationTemplate, otherBureausPresentationDestination, new List<Dictionary<string, string>> { dayPlaceholders, weekPlaceholders });
        ProcessFile(finalPresentationTemplate, finalPresentationDestination, new List<Dictionary<string, string>> { finalPresentationPlaceholders });

        CopyOtherBureauPresentation(otherBureausPresentationDestination);
    }
}