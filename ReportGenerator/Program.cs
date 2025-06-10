using MyOfficeLibrary.Services;

Console.OutputEncoding = System.Text.Encoding.UTF8;
Console.WriteLine("⏳ Инициализация обработки лабораторных работ...");

//try
//{
    Console.Write("📁 Введите путь к папке с работами: C:\\Temp\\Projects\\Тестирование\\LabWorks1 ");
    var folderPath = "C:\\Temp\\Projects\\Тестирование\\LabWorks1";

    //Console.Write("📁 Введите путь к папке с работами: ");
    //var folderPath = Console.ReadLine();


    if (!Directory.Exists(folderPath))
    {
        Console.WriteLine("❌ Ошибка: Указанная папка не существует");
        return;
    }

    IOfficeService officeService = new WordService(true);

    foreach (var file in Directory.GetFiles(folderPath, "Лабораторная работа *.docx"))
    {
        Console.WriteLine($"🔧 Обработка: {Path.GetFileName(file)}");
        officeService.ProcessDocument(file);
    }

    Console.WriteLine("\n🧩 Объединение документов...");
    var outputFile = Path.Combine(folderPath, "Общий_отчет_лабораторных_работ.docx");
    //officeService.MergeDocuments(folderPath, outputFile);
    Console.WriteLine($"\n✅ Готово! Итоговый отчёт сохранён как:\n{outputFile}");
//}
//catch (Exception ex)
//{
//    Console.WriteLine($"\n🚫 Критическая ошибка: {ex.Message}");
//}