using Xunit;
using System;
using System.Collections.Generic;

public class WeeklyReportAutomatorTests: TestsBase
{
    [Fact]
    public void GetFirstSaturday_ReturnsCorrectDate()
    {
        // Arrange
        var today = new DateTime(2023, 10, 10); // Tuesday, October 10, 2023
        var expectedFirstSaturday = new DateTime(2023, 10, 7); // Saturday, October 7, 2023

        // Act
        var firstSaturday = Weekly_Report_Automator.GetFirstSaturday();

        // Assert
        Assert.Equal(expectedFirstSaturday, firstSaturday);
    }

    [Fact]
    public void GetLastFriday_ReturnsCorrectDate()
    {
        // Arrange
        var firstSaturday = new DateTime(2023, 10, 7); // Saturday, October 7, 2023
        var expectedLastFriday = new DateTime(2023, 10, 20); // Friday, October 20, 2023

        // Act
        var lastFriday = Weekly_Report_Automator.GetLastFriday();

        // Assert
        Assert.Equal(expectedLastFriday, lastFriday);
    }

    [Fact]
    public void GenerateDayPlaceholders_ReturnsCorrectPlaceholders()
    {
        // Arrange
        var expectedPlaceholders = new Dictionary<string, string>
        {
            { "$D0", "ΣΑΒ\n07 ΟΚΤ 2023" },
            { "$D1", "ΚΥΡ\n08 ΟΚΤ 2023" },
            { "$D2", "ΔΕΥ\n09 ΟΚΤ 2023" },
            { "$D3", "ΤΡΙ\n10 ΟΚΤ 2023" },
            { "$D4", "ΤΕΤ\n11 ΟΚΤ 2023" },
            { "$D5", "ΠΕΜ\n12 ΟΚΤ 2023" },
            { "$D6", "ΠΑΡ\n13 ΟΚΤ 2023" },
            { "$D7", "ΣΑΒ\n14 ΟΚΤ 2023" },
            { "$D8", "ΚΥΡ\n15 ΟΚΤ 2023" },
            { "$D9", "ΔΕΥ\n16 ΟΚΤ 2023" },
            { "$D10", "ΤΡΙ\n17 ΟΚΤ 2023" },
            { "$D11", "ΤΕΤ\n18 ΟΚΤ 2023" },
            { "$D12", "ΠΕΜ\n19 ΟΚΤ 2023" },
            { "$D13", "ΠΑΡ\n20 ΟΚΤ 2023" }
        };

        // Act
        var dayPlaceholders = Weekly_Report_Automator.GenerateDayPlaceholders();

        // Assert
        Assert.Equal(expectedPlaceholders, dayPlaceholders);
    }

    [Fact]
    public void GenerateWeekPlaceholders_ReturnsCorrectPlaceholders()
    {
        // Arrange
        var expectedPlaceholders = new Dictionary<string, string>
        {
            { "$W0", "16/09 - 22/09" },
            { "$W1", "23/09 - 29/09" },
            { "$W2", "30/09 - 06/10" },
            { "$W3", "07/10 - 13/10" },
            { "$W4", "14/10 - 20/10" },
            { "$W5", "21/10 - 27/10" }
        };

        // Act
        var weekPlaceholders = Weekly_Report_Automator.GenerateWeekPlaceholders();

        // Assert
        Assert.Equal(expectedPlaceholders, weekPlaceholders);
    }

    [Fact]
    public void GenerateFinalPlaceholders_ReturnsCorrectPlaceholders()
    {
        // Arrange
        var expectedPlaceholders = new Dictionary<string, string>
        {
            { "$F0", "ΣΑΒ 07 ΟΚΤ 2023" },
            { "$F1", "ΠΑΡ 20 ΟΚΤ 2023" }
        };

        // Act
        var finalPlaceholders = Weekly_Report_Automator.GenerateFinalPlaceholders();

        // Assert
        Assert.Equal(expectedPlaceholders, finalPlaceholders);
    }
}