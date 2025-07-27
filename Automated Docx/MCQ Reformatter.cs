using System;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

class Program
{
    static void Main()
    {
        var app = new Word.Application();
        var docSrc = app.Documents.Open(@"AP_-_MCQ_Sheet_-_Class_6_-_Chapter_1.1^J_1.2^J_1.3^J_1.4^J_1.5^J_1.6_-_স্বাভাবিক_সংখ্যা_ও_ভগ্নাংশ.docx");
        var docOut = app.Documents.Add();

        try
        {
            // Set two columns for output
            docOut.PageSetup.TextColumns.SetCount(2);

            int i = 1;
            foreach (Word.Paragraph para in docSrc.Paragraphs)
            {
                string text = para.Range.Text.Trim();
                if (string.IsNullOrEmpty(text)) continue;

                // Simple pattern for Bangla MCQ structure: adapt to your real content
                if (IsBanglaQuestion(text))
                {
                    AddPara(docOut, text);
                }
                else if (IsOption(text))
                {
                    AddPara(docOut, text, indent: 20);
                }
                else if (IsAnswer(text))
                {
                    AddPara(docOut, text, bold: true, color: 0x088565);
                }
                else
                {
                    AddPara(docOut, text); // For any reference or explanation line
                }

                // Copy over equations as objects
                foreach (Word.OMaths eq in para.Range.OMaths)
                {
                    eq.Range.Copy();
                    Word.Paragraph eqPara = docOut.Content.Paragraphs.Add();
                    eqPara.Range.Paste();
                }
            }

            docOut.SaveAs2(@"OUTPUT.docx");
        }
        finally
        {
            docSrc.Close();
            docOut.Close();
            app.Quit();
        }
    }

    static bool IsBanglaQuestion(string text)
    {
        // Bangla digit and dot
        return text.Length > 2 && char.GetUnicodeCategory(text[0]) == System.Globalization.UnicodeCategory.DecimalDigitNumber
            && text[1] == '.' && text[2] == ' ';
    }

    static bool IsOption(string text)
    {
        // Bangla options: ক. খ. গ. ঘ.
        return text.Length > 2 && "কখগঘ".Contains(text[0]) && text[1] == '.' && text[2] == ' ';
    }

    static bool IsAnswer(string text)
    {
        return text.StartsWith("উত্তর:") || text.StartsWith("Ans:");
    }

    static void AddPara(Word.Document doc, string text, bool bold = false, int indent = 0, int color = 0)
    {
        Word.Paragraph para = doc.Content.Paragraphs.Add();
        para.Range.Text = text;
        para.Range.Font.Name = "Tiro Bangla";
        para.Range.Font.Size = 11;
        para.Format.SpaceAfter = 0;
        para.Format.SpaceBefore = 0;
        para.Range.Font.Bold = bold ? 1 : 0;
        if (color != 0)
        {
            para.Range.Font.Color = (Word.WdColor)color;
        }
        if (indent > 0)
        {
            para.Format.LeftIndent = indent;
        }
        para.Range.InsertParagraphAfter();
    }
}
