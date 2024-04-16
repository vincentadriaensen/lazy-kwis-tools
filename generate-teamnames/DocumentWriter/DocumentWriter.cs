using Word = Microsoft.Office.Interop.Word;

namespace generate_teamnames.DocumentWriter;

internal class DocumentWriter
{
    private Word._Application oWord;
    private Word._Document oDoc;

    private object oPageBreak = Word.WdBreakType.wdSectionBreakNextPage; //.wdPageBreak;

    private object oMissing = System.Reflection.Missing.Value;
    private object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

    public DocumentWriter()
    {
        oWord = new Word.Application();
        oWord.Visible = false;
        oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);

        // set to landscape mode
        oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
    }

    internal void Run()
    {
        // todo: make the filename a parameter
        using StreamReader sr = new StreamReader("Example/input_teamnames.txt");
        string rawInput;

        // Loop over each line, insert into Word Document
        // Each line represents another teamname
        while ((rawInput = sr.ReadLine()) != null)
        {
            try
            {
                Console.WriteLine($"Adding {rawInput} to the list");

                var inputComponents = rawInput.Split(' ').ToList();
                var inputComponentsContainTeamNumber = int.TryParse(inputComponents.FirstOrDefault(), out var tn);

                string teamName = inputComponentsContainTeamNumber ? string.Join(' ', inputComponents.Skip(1).ToList()) : string.Join(' ', inputComponents.ToList());
                string? teamNumber = inputComponentsContainTeamNumber ? tn.ToString() : null;

                InsertNewPage(teamName, teamNumber);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"EXCEPTION: Could not generate page for {rawInput}: {ex.Message}");
            }
        }

        // Open the document now it's finished so it can be saved
        oWord.Visible = true;
    }

    void InsertNewPage(string teamName, string? teamNumber)
    {
        Word.Paragraph oPara1;
        oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);

        // Start with a new line on the page (puts the text more in center)
        oPara1.Range.Font.Size = 100; // use a fontsize of 30 for the initial new line

        if (teamNumber != null)
            oPara1.Range.Text = $"{teamNumber}\n{teamName}";
        else
            oPara1.Range.Text = teamName;

        oPara1.Range.Font.Bold = 1;
        oPara1.Format.SpaceBefore = 24;      // 24 pt spacing before paragraph.
        oPara1.Format.SpaceAfter = 24;       // 24 pt spacing after paragraph.
        oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; // Center align the text

        //oPara1.Range.Font.Size = 30; // use a fontsize of 30 for the initial new line

        // try to adjust the fontsize according to the lenght of the teamname
        oPara1.Range.Font.Size = teamName.Length <= 7 ? 125 : teamName.Length <= 9 ? 100 : teamName.Length <= 12 ? 75 : teamName.Length <= 19 ? 65 : teamName.Length <= 27 ? 60 : teamName.Length <= 32 ? 50 : 35;

        var wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
        wrdRng.InsertBreak(ref oPageBreak);
    }
}
