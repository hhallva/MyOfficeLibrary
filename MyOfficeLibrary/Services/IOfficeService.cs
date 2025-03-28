using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyOfficeLibrary.Services
{
    public interface IOfficeService : IDisposable
    {
        bool CreateFile(string? filePath = null);
        bool OpenFile(string filePath);
        bool SaveFile(string filePath);
        bool ExportToPdf(string filePath);
        bool CloseFile();
    }
}
