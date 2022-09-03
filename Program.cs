// See https://aka.ms/new-console-template for more information
using JournalGenerator.Utils;

Console.WriteLine("Hello, World!");
while (true)
{
    Console.Clear();

    Console.WriteLine("Applications");
    Console.WriteLine("  1: Generate Journal from Islandsbanki Foreign Payments");
    Console.WriteLine();
    Console.WriteLine("Enter the number of application to run or press 'x' to exit: ");


    string input = Console.ReadLine().ToLower();

    if (input == "x")
    {
        break;
    }

    int applicationNr = -1;
    if (!int.TryParse(input, out applicationNr))
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.Write("Only input numbers ");
        Console.ResetColor();
        Console.ReadKey();
    }
    else
    {
        switch (applicationNr)
        {
            case 1:
                GenerateJournalFromIslandsbankiForeignPayments();
                break;
            default:
                Console.Write("Invalid applicaiton number");
                Console.ResetColor();
                Console.ReadKey();
                break;
        }
    }
}

static void GenerateJournalFromIslandsbankiForeignPayments()
{
    string filePath;
    FileInfo fileInfo;
    while (true)
    {
        Console.WriteLine("Enter file path: ");
        filePath = Console.ReadLine();
        //filePath = @"C:\Users\egill\Desktop\Journal.xlsx"; // DEBUG
        if (!File.Exists(filePath))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Error: File not found");
            Console.ResetColor();
            continue;
        }

        fileInfo = new FileInfo(filePath);

        if (string.IsNullOrWhiteSpace(fileInfo.Extension) || !fileInfo.Extension.Equals(".xlsx", StringComparison.InvariantCultureIgnoreCase))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Error: File must be an Excel (.xlsx) file");
            Console.ResetColor();
        }
        else
        {
            break;
        }
    }

    try
    {
        byte[] fileBytes = File.ReadAllBytes(filePath);
        var outputFilePath = string.Format("{0}_ouput.xlsx", filePath.Replace(fileInfo.Extension, string.Empty));

        (byte[] outputBytes, int journalEntriesCount) = PaydayJournalUtils.GenerateJournalFromIslandsbankiForeignPayments(fileBytes);

        File.WriteAllBytes(outputFilePath, outputBytes);


        Console.WriteLine(string.Format("Successfully generated {0} journal entries to file: {1}", journalEntriesCount, outputFilePath));
        Console.ReadLine();
    }
    catch (Exception ex)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        string message = string.Format("{0} {1}", ex.Message, ex.InnerException != null ? ex.InnerException.Message : string.Empty);
        Console.WriteLine(string.Format("Error: {0}", message));
        Console.ResetColor();
        Console.ReadLine();
    }
}