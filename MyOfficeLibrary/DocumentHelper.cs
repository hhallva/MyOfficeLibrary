using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace MyOfficeLibrary
{
    public static class DocumentHelper
    {
        private static Document _doc { get; set; }

        private static readonly List<string> SectionsToRemove = new()
        {
            "Литература",
            "Подготовка к работе",
            "Основное оборудование",
            "Задание",
            "Порядок выполнения работы",
            "Содержание отчета"
        };

        public static void ProcessSections(Document doc)
        {
            _doc = doc;

            RemoveHeader();
            CreateChapters();
            CreateQuestions();
            CreateConclusion();
        }

        private static void RemoveHeader()
        {
            try
            {
                if (_doc.Sections.Count >= 1)
                {
                    Section firstSection = _doc.Sections[1];

                    if (firstSection.PageSetup.DifferentFirstPageHeaderFooter != 1)
                    {
                        HeaderFooter firstPageHeader = firstSection.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage];
                        if (firstPageHeader != null && firstPageHeader.Exists)
                        {
                            firstPageHeader.Range.Delete();
                            firstPageHeader.LinkToPrevious = false;
                        }
                    }
                    else
                    {
                        HeaderFooter primaryHeader = firstSection.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                        if (primaryHeader != null && primaryHeader.Exists)
                        {
                            primaryHeader.Range.Delete();
                            primaryHeader.LinkToPrevious = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при удалении первого колонтитула: {ex.Message}");
            }
        }
        

        private static void CreateChapters()
        {
            List<Paragraph> headerParagraphs = GetHeaders();

            for (int i = 0; i < headerParagraphs.Count; i++)
            {
                var headerPara = headerParagraphs[i];
                var headerText = headerPara.Range.Text.Replace("\r", "");

                if (headerText == "Контрольные вопросы")
                    ChangeHeaderText(headerPara, "Ответы на контрольные вопросы");
                else if (SectionsToRemove.Contains(headerText))
                    RemoveSectionContent(headerPara, headerParagraphs, i);
            }
        }

        private static void CreateQuestions()
        {
            var questionParagraphs = new List<Paragraph>();
            for (int i = 1; i <= _doc.Paragraphs.Count; i++)
            {
                var para = _doc.Paragraphs[i];
                if (IsQuestion(para))
                    questionParagraphs.Add(para);
            }

            for (int i = 0; i < questionParagraphs.Count; i++)
                AddAnswerToQuestion(questionParagraphs[i]);
        }

        private static void CreateConclusion()
        {
            try
            {
                List<Paragraph> headerParagraphs = GetHeaders();

                var start = headerParagraphs[0].Range.Start;
                var end = headerParagraphs[1].Range.Start;

                Word.Range chapter = _doc.Range(start, end);
                chapter.Copy();

                Word.Range endRange = _doc.Range(_doc.Content.End - 1, _doc.Content.End - 1);
                endRange.Paste();
                _doc.Range(_doc.Content.End - 1, _doc.Content.End).Delete();

                Word.Range newChapterRange = _doc.Range(endRange.Start, _doc.Content.End);

                ProcessCopiedChapter(newChapterRange);
                _doc.Range(_doc.Content.End - 1, _doc.Content.End).Delete();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка в CreateConclusion: {ex.Message}");
            }
        }


        private static List<Paragraph> GetHeaders()
        {
            var headerParagraphs = new List<Paragraph>();
            for (int i = 1; i <= _doc.Paragraphs.Count; i++)
            {
                var para = _doc.Paragraphs[i];
                if (IsHeader(para))
                    headerParagraphs.Add(para);
            }

            return headerParagraphs;
        }

        private static bool IsHeader(Paragraph paragraph)
        {
            try
            {
                return paragraph.get_Style().NameLocal == "Основная нумерация 1";
            }
            catch
            {
                return false;
            }
        }

        private static bool IsSubHeader(Paragraph paragraph)
        {
            try
            {
                return paragraph.get_Style().NameLocal == "Основная нумерация 2";
            }
            catch
            {
                return false;
            }
        }

        private static bool IsQuestion(Paragraph paragraph)
        {
            try
            {
                return paragraph.Range.Text.Contains("?");
            }
            catch
            {
                return false;
            }
        }


        private static void RemoveSectionContent(Paragraph header, List<Paragraph> allHeaders, int currentIndex)
        {
            try
            {
                int start = header.Range.Start;
                int end = _doc.Content.End;

                if (currentIndex < allHeaders.Count - 1)
                    end = allHeaders[currentIndex + 1].Range.Start;

                int length = end - start;
                if (length <= 0)
                    return;

                var rangeToDelete = _doc.Range(start, end);
                if (rangeToDelete != null)
                    rangeToDelete.Delete();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка удаления раздела: {ex.Message}");
            }
        }

        private static void ProcessCopiedChapter(Word.Range chapterRange)
        {

            foreach (Paragraph para in chapterRange.Paragraphs)
            {
                if (IsHeader(para))
                {
                    ChangeHeaderText(para, "Вывод");
                    break;
                }
            }

            var count = chapterRange.Paragraphs.Count;
            for (int i = 1; i <= count; i++)
            {
                var para = chapterRange.Paragraphs[i];
                if (IsSubHeader(para))
                {
                    string text = $"В ходе проделанной лабораторной работы {TransformVerbs(para.Range.Text.Trim('\r', '\a', ' '))}\r";
                    para.Range.Text = text;
                }
            }
        }

        private static void ChangeHeaderText(Paragraph headerPara, string newText)
        {
            Word.Range range = headerPara.Range;
            range.Text = $"{newText}\r";
            range.set_Style("Основная нумерация 1");
        }

        private static void AddAnswerToQuestion(Paragraph questionPara)
        {
            try
            {
                int endPos = questionPara.Range.End;
                Word.Range answerRange = _doc.Range();

                if (endPos < _doc.Content.End)
                {
                    answerRange.SetRange(endPos, endPos);
                }
                else
                {
                    answerRange.SetRange(endPos - 1, endPos - 1);
                    answerRange.InsertParagraphAfter();

                    answerRange.SetRange(endPos, endPos);
                    answerRange.Text = "Ответ:\n";
                    answerRange.set_Style("Без интервала");

                    answerRange.Font.Name = "Times New Roman";
                    answerRange.Font.Size = 14;

                    return;
                }

                answerRange.Text = "Ответ:\n\n";
                answerRange.Font.Name = "Times New Roman";
                answerRange.Font.Size = 14;
                answerRange.set_Style("Без интервала");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при добавлении ответа: {ex.Message}");
            }
        }

        private static string TransformVerbs(string objective)
        {
            return objective.Replace("Изучить", "изучил")
                            .Replace("Научиться", "научился")
                            .Replace("Закрепить", "закрепил")
                            .Replace("Получить", "получил")
                            .Replace("Освоить", "освоил")
                            .Replace("Исследовать", "исследовал")
                            .Replace("Разработать", "разработал")
                            .Replace("Создать", "создал");
        }
    }
}
