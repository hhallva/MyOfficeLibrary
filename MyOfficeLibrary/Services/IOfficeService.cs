namespace MyOfficeLibrary.Services
{
    public interface IOfficeService : IDisposable
    {
        bool CreateFile(string? filePath = null);
        bool OpenFile(string filePath);
        bool SaveFile(string filePath);
        bool CloseFile();

        void ProcessDocument(string filePath);
        void MergeDocuments(string folderPath, string outputFile);
    }
}
