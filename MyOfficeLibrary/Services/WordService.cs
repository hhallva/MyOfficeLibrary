using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace MyOfficeLibrary.Services
{
    public class WordService : IOfficeService
    {
        private Application _wordApp;
        private Document _document;
        private bool _isOpen = false;

        public WordService(bool visible)
        {
            _wordApp = new Application { Visible = visible, DisplayAlerts = WdAlertLevel.wdAlertsNone };
        }

        #region Базовые функции
        public bool CreateFile(string? filePath = null)
        {
            try
            {
                _document = _wordApp.Documents.Add();
                _isOpen = true;

                if (filePath != null)
                {
                    _document.SaveAs(filePath);
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при создании Word файла: {ex.Message}");
                return false;
            }
        }

        public bool OpenFile(string filePath)
        {
            try
            {
                _document = _wordApp.Documents.Open(filePath);
                _isOpen = true;
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при открытии Word файла: {ex.Message}");
                return false;
            }
        }

        public bool SaveFile(string filePath)
        {
            if (!_isOpen || _document == null) return false;

            try
            {
                _document.SaveAs(filePath);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при сохранении Word файла: {ex.Message}");
                return false;
            }
        }

        public bool CloseFile()
        {
            if (_document != null && _isOpen)
            {
                _document.Close();
                _isOpen = false;
            }
            return true;
        }

        public bool ExportToPdf(string filePath, string subject, string body)
        {
            if (!_isOpen || _document == null)
                return false;

            try
            {
                CreateFile(filePath);
                AddHeading(subject);
                AddParagraph(body);

                //_document.SaveAs(filePath, WdExportFormat.wdExportFormatPDF);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при экспорте PDF файла: {ex.Message}");
                return false;
            }
        }

        public string? ReadAllText()
        {
            if (!_isOpen || _document == null)
                return null;
            return _document.Content.Text;
        }

        public void Dispose()
        {
            CloseFile();
            if (_wordApp != null)
            {
                _wordApp.Quit();
                Marshal.ReleaseComObject(_wordApp);
                _wordApp = null;
            }
        }
        #endregion

        #region Работа с текстом
        public bool ReplaceText(string searchText, string replaceText)
        {
            if (!_isOpen || _document == null)
                return false;

            var range = _document.Content;
            range.Find.Text = searchText;
            range.Find.Replacement.Text = replaceText;

            object replaceAll = WdReplace.wdReplaceAll;
            range.Find.Execute(Replace: ref replaceAll);

            return true;
        }

        public bool AddHeading(string text)
        {
            if (!_isOpen || _document == null)
                return false;

            try
            {
                Paragraph paragraph = _document.Content.Paragraphs.Add();
                paragraph.Range.Text = text;
                paragraph.Range.InsertParagraphAfter();

                var range = paragraph.Range;
                range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                range.Font.Size = 16;
                range.Font.Bold = 1;


                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка добавления заголовка: {ex.Message}");
                return false;
            }
        }

        public bool AddParagraph(string text)
        {
            if (!_isOpen || _document == null)
                return false;

            try
            {
                Paragraph paragraph = _document.Content.Paragraphs.Add();
                paragraph.Range.Text = text;
                paragraph.Range.InsertParagraphAfter();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при добавлении параграфа: {ex.Message}");
                return false;
            }
        }
        #endregion

        public void ProcessDocument(string filePath)
        {
            //try
            //{
                OpenFile(filePath);
                DocumentHelper.ProcessSections(_document);
                //SaveFile(_document.FullName);
            //}
            //finally
            //{
            //    Dispose();
            //}
        }

        public void MergeDocuments(string folderPath, string outputFile)
        {
            throw new NotImplementedException();
        }
    }
}