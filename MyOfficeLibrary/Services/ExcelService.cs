using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace MyOfficeLibrary.Services
{
    public class ExcelService : IOfficeService
    {
        private Application _excelApp;
        private Workbook _workbook;
        private Worksheet _worksheet;
        private bool _isOpen = false;

        public ExcelService(bool visible)
        {
            _excelApp = new Application();
            _excelApp.Visible = visible;
        }

        public bool CreateFile(string? filePath = null)
        {
            throw new NotImplementedException();
        }

        public bool OpenFile(string filePath)
        {
            try
            {
                _workbook = _excelApp.Workbooks.Open(filePath);
                _worksheet = _workbook.ActiveSheet;
                _isOpen = true;
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при открытии Excel файла: {ex.Message}");
                return false;
            }
        }

        public bool SaveFile(string filePath)
        {
            if (!_isOpen || _workbook == null) return false;
            try
            {
                _workbook.SaveAs(filePath);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при сохранении Excel файла: {ex.Message}");
                return false;
            }
        }

        public bool ExportToPdf(string filePath)
        {
            if (!_isOpen || _workbook == null) return false;

            try
            {
                _workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, filePath);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при экспорте PDF файла: {ex.Message}");
                return false;
            }
        }

        public bool CloseFile()
        {
            if (_workbook != null && _isOpen)
            {
                _workbook.Close();
                _isOpen = false;
            }
            return true;
        }

        public void Dispose()
        {
            CloseFile();
            if (_excelApp != null)
            {
                _excelApp.Quit();
                Marshal.ReleaseComObject(_excelApp);
                _excelApp = null;
            }
        }

        public bool WriteCellValue(int row, int column, string value)
        {
            if (!_isOpen || _worksheet == null) return false;
            try
            {
                Excel.Range cell = _worksheet.Cells[row, column];
                cell.Value = value;
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка записи ячейки: {ex.Message}");
                return false;
            }
        }

        public string? ReadCellValue(int row, int column)
        {
            if (!_isOpen || _worksheet == null) return null;
            try
            {
                Excel.Range cell = _worksheet.Cells[row, column];
                return cell.Value;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка чтения ячейки: {ex.Message}");
                return null;
            }
        }

        public void ProcessDocument(string filePath)
        {
            throw new NotImplementedException();
        }

        public void MergeDocuments(string folderPath, string outputFile)
        {
            throw new NotImplementedException();
        }
    }
}
