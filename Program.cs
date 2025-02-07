using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Linq;

class Program
{
    static void Main()
    {
        // This is a placeholder
        string sourceFilePath = Path.Combine(AppContext.BaseDirectory, "C:\\Users\\Giorgos\\source\\repos\\Weeekly_Report_Automator\\template\\template.pptx");
        string destinationFilePath = Path.Combine(AppContext.BaseDirectory, "C:\\Users\\Giorgos\\source\\repos\\Weeekly_Report_Automator\\output\\modified.pptx");

        // Ensure the destination directory exists
        Directory.CreateDirectory(Path.GetDirectoryName(destinationFilePath));

        // Copy the source file to the destination file before modifying
        File.Copy(sourceFilePath, destinationFilePath, true);

        // Define Greek Day format
        string[] greekDays = { "ΚΥΡ", "ΔΕΥ", "ΤΡΙ", "ΤΕΤ", "ΠΕΜ", "ΠΑΡ", "ΣΑΒ" };
        CultureInfo greekCulture = new CultureInfo("el-GR");

        DateTime date = DateTime.UtcNow.Date;
        string formattedDate = $"{greekDays[(int)date.DayOfWeek]}\n{date:dd MMM yyyy}".ToUpper();
        string searchText = "{PLACEHOLDER}";

        // Define placeholders
        Dictionary<string, string> placeholderDates = new Dictionary<string, string>();

        // Get first day of retrospective
        var firstSaturday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek -1);
        for (int i = 0; i < 14; i++)
        {
            DateTime dayOfWeek = firstSaturday.AddDays(i);
            placeholderDates[$"{{PLACEHOLDER}}{i}"] = $"{greekDays[(int)dayOfWeek.DayOfWeek]}\n{dayOfWeek:dd MMM yyyy}".ToUpper();
            Console.WriteLine(placeholderDates[$"{{PLACEHOLDER}}{i}"]);
        }

        //using (PresentationDocument presentation = PresentationDocument.Open(destinationFilePath, true))
        //{
        //    var slideParts = presentation.PresentationPart.SlideParts;

        //    foreach (var slidePart in slideParts)
        //    {
        //        var textElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>();

        //        foreach (var textElement in textElements)
        //        {
        //            if (textElement.Text.Contains(searchText))
        //            {
        //                textElement.Text = textElement.Text.Replace(searchText, formattedDate);
        //            }
        //        }
        //    }
        //}
    }
}