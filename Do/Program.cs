using System.Data;
using System.Text;
using IronXL;

internal class Program
{
    /// <summary>
    /// Converts CardCastle !Simple! export to Moxfield import format.
    /// </summary>
    /// <test>
    /// Uncomment two below params and your collection should have 4 new cards.
    /// </test>
    /// <exception cref="Exception"></exception>
    private static void Main(string[] args)
    {
        // (!Insert csv file next to 'Do.csproj' file and force a copy action to bin folder. Then update name below.)
        var fileName = string.Empty; //"export_simple_1681933993.csv"; 

        // (!Moxfield has such date in export format. Just update current date.)
        var dateToUpdate = string.Empty; //"2023-04-19 20:00:00.000000";

        if (string.IsNullOrWhiteSpace(fileName)
            || string.IsNullOrWhiteSpace(dateToUpdate))
        {
            throw new Exception("Update above parameters!");
        }

        var list = new List<string>
        {
            "\"Count\",\"Tradelist Count\",\"Name\",\"Edition\",\"Condition\",\"Language\",\"Foil\",\"Tags\",\"Last Modified\",\"Collector Number\",\"Alter\",\"Proxy\",\"Purchase Price\""
        };
        foreach (var row in WorkBook.LoadCSV(fileName, ExcelFileFormat.XLS, ",").ToDataSet().Tables[0].Rows)
        {
            var count = (double?)((DataRow)row).ItemArray[0];
            var cardName = (string?)((DataRow)row).ItemArray[1];
            var setName = (string?)((DataRow)row).ItemArray[2];
            var collectorNumber = ((DataRow)row).ItemArray[3].ToString(); // Sometimes double => 211 | Sometimes string => 211a
            var foil = (bool?)((DataRow)row).ItemArray[4];

            if (count == null
                || string.IsNullOrWhiteSpace(cardName)
                || string.IsNullOrWhiteSpace(setName)
                || string.IsNullOrWhiteSpace(collectorNumber)
                || foil == null)
            {
                throw new Exception("Empty values not allowed! Check CardCastle export!");
            }

            var isFoil = (bool)foil ? "foil" : string.Empty;
            var stringBuilder = new StringBuilder();
            stringBuilder.Append($"\"{(int)count}\",");
            stringBuilder.Append($"\"{(int)count}\",");
            stringBuilder.Append($"\"{cardName}\",");
            stringBuilder.Append($"\"{setName}\",");
            stringBuilder.Append($"\"Near Mint\",");
            stringBuilder.Append($"\"English\",");
            stringBuilder.Append($"\"{isFoil}\",");
            stringBuilder.Append($"\"\",");
            stringBuilder.Append($"\"{dateToUpdate}\",");
            stringBuilder.Append($"\"{collectorNumber}\",");
            stringBuilder.Append($"\"False\",");
            stringBuilder.Append($"\"False\",");
            stringBuilder.Append($"\"\"");
            list.Add(stringBuilder.ToString());
        }
        File.WriteAllLines("Csv_For_Moxfield_Import.csv", list);
        Console.WriteLine("DONE!");
    }
}