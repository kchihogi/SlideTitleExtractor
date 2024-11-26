using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

class Program
{
    // タイトルに相当する名前を複数の言語で定義
    static readonly string[] TitleKeywords = { "title", "タイトル", "titolo", "titre", "título", "überschrift" };

    static void Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.WriteLine("PowerPointファイルのパスを指定してください。");
            return;
        }

        string filePath = args[0];
        string outputFormat = args.Length > 1 ? args[1].ToLower() : "text";

        // ファイルの存在確認
        if (!File.Exists(filePath))
        {
            Console.WriteLine("指定されたファイルが見つかりません: " + filePath);
            return;
        }

        try
        {
            ExtractSlideTitles(filePath, outputFormat);
        }
        catch (Exception ex)
        {
            Console.WriteLine("エラーが発生しました: " + ex.Message);
        }
    }

    static void ExtractSlideTitles(string filePath, string outputFormat)
    {
        using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, false))
        {
            var presentationPart = presentationDocument.PresentationPart;

            if (presentationPart == null)
            {
                Console.WriteLine("プレゼンテーションデータが無効です。");
                return;
            }

            int slideIndex = 1;
            var slideTitles = new List<object>();

            var slideIdList = presentationPart.Presentation.SlideIdList;
            if (slideIdList == null)
            {
                Console.WriteLine("スライドが見つかりませんでした。");
                return;
            }

            var slideIds = slideIdList.ChildElements.OfType<SlideId>();
            foreach (var slideId in slideIds)
            {
                var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
                string title = GetSlideTitle(slidePart);
                slideTitles.Add(new { Slide = slideIndex, Title = title });
                slideIndex++;
            }

            // Output based on specified format
            if (outputFormat == "text")
            {
                string outputFilePath = Path.ChangeExtension(filePath, ".txt");
                using (StreamWriter writer = new StreamWriter(outputFilePath, false))
                {
                    foreach (var slide in slideTitles)
                    {
                        writer.WriteLine($"Slide {((dynamic)slide).Slide}: {((dynamic)slide).Title}");
                    }
                }
                Console.WriteLine("スライドタイトルをテキストファイルに出力しました: " + outputFilePath);
            }
            else if (outputFormat == "json")
            {
                string outputFilePath = Path.ChangeExtension(filePath, ".json");
                string jsonString = JsonSerializer.Serialize(slideTitles, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(outputFilePath, jsonString);
                Console.WriteLine("スライドタイトルをJSONファイルに出力しました: " + outputFilePath);
            }
            else
            {
                Console.WriteLine("不明な出力形式が指定されました。text または json を指定してください。");
            }
        }
    }

    static string GetSlideTitle(SlidePart slidePart)
    {
        // スライド内のタイトルテキストを探す
        var titleShape = slidePart.Slide.Descendants<Shape>()
            .FirstOrDefault(s =>
                TitleKeywords.Any(keyword =>
                    (s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.ToLower().Contains(keyword.ToLower()) ?? false)));

        if (titleShape != null)
        {
            return titleShape.TextBody?.InnerText ?? "（タイトルなし）";
        }

        return "（タイトルなし）";
    }
}
