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

        public void ProcessDocument(string filePath)
        {
            try
            {
                OpenFile(filePath);
                DocumentHelper.ProcessSections(_document);
                SaveFile(_document.FullName);
            }
            finally
            {
                CloseFile();
            }
        }

        public void MergeDocuments(string folderPath, string outputFile)
        {
            throw new NotImplementedException();
        }
    }
}