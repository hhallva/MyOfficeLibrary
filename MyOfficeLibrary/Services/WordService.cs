using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace MyOfficeLibrary.Services
{
    public class WordService : IOfficeService
    {
        private Application _wordApp;
        private Document _document;
        private bool _isOpen = false;

        public WordService(bool visible)
        {
            _wordApp = new Application();
            _wordApp.Visible = visible;
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

        public bool ExportToPdf(string filePath)
        {
            if (!_isOpen || _document == null) return false;

            try
            {
                _document.ExportAsFixedFormat(filePath, WdExportFormat.wdExportFormatPDF);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при экспорте PDF файла: {ex.Message}");
                return false;
            }
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
    }
}
