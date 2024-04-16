using Word = Microsoft.Office.Interop.Word;

namespace generate_teamnames.DocumentWriter;

internal class DocumentWriter
{
    private Word._Application oWord;
    private Word._Document oDoc;

    private object oPageBreak = Word.WdBreakType.wdPageBreak;

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
        string teamName;

        // Loop over each line, insert into Word Document
        // Each line represents another teamname
        while ((teamName = sr.ReadLine()) != null)
        {
            try
            {
                Console.WriteLine($"Adding {teamName} to the list");
                InsertNew(teamName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"EXCEPTION: Could not generate page for {teamName}: {ex.Message}");
            }
        }

        // Open the document now it's finished so it can be saved
        oWord.Visible = true;
    }

    void InsertNew(string teamName)
    {
        Word.Paragraph oPara1;
        oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
        oPara1.Range.Text = teamName;
        oPara1.Range.Font.Bold = 1;
        oPara1.Range.Font.Size = 72;        // Set the font size as big as possible. Adjust as needed.
        oPara1.Format.SpaceAfter = 24;      // 24 pt spacing after paragraph.
        oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; // Center align the text
        oPara1.Range.InsertParagraphAfter();

        var wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
        wrdRng.InsertBreak(ref oPageBreak);
    }

}
