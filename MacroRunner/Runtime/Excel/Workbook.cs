namespace MacroRunner.Runtime.Excel;

public class Workbook
{
    public string Path { get; protected set; }

    public void Close(bool saveChanges)
    {
    }

    public Worksheet Sheets(string name)
    {
        return null;
    }
}