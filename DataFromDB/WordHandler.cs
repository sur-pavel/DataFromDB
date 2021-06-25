using System;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Word;

namespace DataFromDB
{
    internal class WordHandler
    {
        private Document objDoc;
        private Application wordApp;
        private Paragraph paragraph;
        private string appPath;

        internal WordHandler()
        {
            wordApp = new Application();
            objDoc = wordApp.Documents.Add();
            paragraph = objDoc.Paragraphs.Add();
        }


        internal void AddString(String str)
        {            
            paragraph.Range.Text += str;
        }

        internal void SaveDoc()
        {
            String appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            wordApp.ActiveDocument.SaveAs(appPath + @"\InvNums.doc", WdSaveFormat.wdFormatDocument);
        }
        internal void Quit()
        {            
            objDoc.Close();
            wordApp.Quit();
        }
    }
}