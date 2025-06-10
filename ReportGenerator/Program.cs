using MyOfficeLibrary.Services;
using System.Diagnostics;
using Timers = System.Timers;

Console.OutputEncoding = System.Text.Encoding.UTF8;
Console.WriteLine("⏳ Инициализация обработки лабораторных работ...");

var stopwatch = new Stopwatch();

try
{
    Console.Write("📁 Введите путь к папке с работами:");
    
    Console.Write("📁 Введите путь к папке с работами: ");
    var folderPath = Console.ReadLine();

    if (!Directory.Exists(folderPath))
    {
        Console.WriteLine("❌ Ошибка: Указанная папка не существует");
        return;
    }

    IOfficeService officeService = new WordService(false);
    stopwatch.Start();
    foreach (var file in Directory.GetFiles(folderPath, "Лабораторная работа *.docx"))
    {
        
        Console.WriteLine($"🔧 Обработка: {Path.GetFileName(file)}");
        officeService.ProcessDocument(file);
       
    }
    officeService.Dispose();

    stopwatch.Stop();
    Console.WriteLine($"\n🕒 Время затраченное на создание шаблонов: {stopwatch.Elapsed}");

    Console.WriteLine("\n🧩 Объединение документов...");
    var outputFile = Path.Combine(folderPath, "Общий_отчет_лабораторных_работ.docx");
    //officeService.MergeDocuments(folderPath, outputFile);
    Console.WriteLine($"\n✅ Готово! Итоговый отчёт сохранён как:\n{outputFile}");
}
catch (Exception ex)
{
    Console.WriteLine($"\n🚫 Критическая ошибка: {ex.Message}");
}