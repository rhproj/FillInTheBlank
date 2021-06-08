using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;

namespace DocFiller
{
    class WordHelper
    {
        private FileInfo _fileInfo;

        public WordHelper(string fileName)
        {
            if (File.Exists(fileName))
            {
                _fileInfo = new FileInfo(fileName);
            }
            else
            {
                throw new ArgumentNullException("Шаблон документа не найден");
            }
        }

        internal void Replace(Dictionary<string, string> doc)
        {
            MSWord.Application wordApp = new MSWord.Application(); //wordApp.Visible = false;
            var file = _fileInfo.FullName;

            try
            {
                wordApp.Documents.Open(file, ReadOnly:true);

                foreach (var d in doc)
                {
                    MSWord.Find find = wordApp.Selection.Find; 
                    find.Text = d.Key;  //(<тэги>)
                    find.Replacement.Text = d.Value;

                    var wrap = MSWord.WdFindWrap.wdFindContinue;
                    var replace = MSWord.WdReplace.wdReplaceAll;

                    find.Execute(FindText: Type.Missing, 
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: Type.Missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: Type.Missing, Replace: replace);
                }

                var newFileName = Path.Combine(_fileInfo.DirectoryName, $"Документ от {DateTime.Now.ToString("dd.MM.yy")}.docx");

                wordApp.ActiveDocument.SaveAs2(newFileName);
                wordApp.ActiveDocument.Close();
                wordApp.Quit();
                wordApp = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (wordApp != null )
                {
                    wordApp.Quit();
                    wordApp = null;
                }
            }
        }
    }
}
