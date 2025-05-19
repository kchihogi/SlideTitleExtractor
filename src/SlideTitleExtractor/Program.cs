using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Text.RegularExpressions;

class Config
{
    public bool IncludeSlideNumber { get; set; } = false;
    public int TitleMaxLength { get; set; } = 255;
    public bool UniqueTitles { get; set; } = true;

    public bool DeleteEmptyTitle { get; set; } = true;

    public bool OnlyBiggerLetters { get; set; } = true;

    public double DefaultFontSize { get; set; } = 6000;

    public static Config LoadConfig(string configPath)
    {
        if (File.Exists(configPath))
        {
            string json = File.ReadAllText(configPath);
            return JsonSerializer.Deserialize<Config>(json) ?? new Config();
        }
        return new Config();
    }
}

class SlideTitle
{
    public int? Slide { get; set; }
    public string Title { get; set; }
}

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

        // コンフィグの読み込み
        string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SlideTitleExtractor.config.json");
        Config config = Config.LoadConfig(configPath);

        // ファイルの存在確認
        if (!File.Exists(filePath))
        {
            Console.WriteLine("指定されたファイルが見つかりません: " + filePath);
            return;
        }

        try
        {
            ExtractSlideTitles(filePath, outputFormat, config);
        }
        catch (Exception ex)
        {
            Console.WriteLine("エラーが発生しました: " + ex.Message);
        }
    }

    static void ExtractSlideTitles(string filePath, string outputFormat, Config config)
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
            var slideTitles = new List<SlideTitle>();

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
                string title = GetSlideTitle(slidePart, config);
                if (title.Length > config.TitleMaxLength)
                {
                    title = title.Substring(0, config.TitleMaxLength);
                }
                slideTitles.Add(new SlideTitle { Slide = config.IncludeSlideNumber ? slideIndex : (int?)null, Title = title });
                slideIndex++;
            }

            // UniqueTitles
            if (config.UniqueTitles)
            {
                // if slide titles duplicate, add #1, #2, #3, ... to the end of the title
                var slideTitlesGrouped = slideTitles.GroupBy(s => s.Title);
                foreach (var group in slideTitlesGrouped)
                {                    
                    if (group.Count() > 1)
                    {
                        int index = 1;
                        foreach (var slide in group)
                        {
                            slide.Title += $" #{index}";
                            index++;
                        }
                    }
                }
            }

            // Output based on specified format
            if (outputFormat == "text")
            {
                string outputFilePath = Path.ChangeExtension(filePath, ".txt");
                using (StreamWriter writer = new StreamWriter(outputFilePath, false))
                {
                    foreach (var slide in slideTitles)
                    {
                        string slideInfo = config.IncludeSlideNumber ? $"Slide {((dynamic)slide).Slide}: " : "";
                        writer.WriteLine(slideInfo + ((dynamic)slide).Title);
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

    static string RemoveSpacesExceptBetweenAlphabets(string text)
    {
        // アルファベットの間のスペース: 残す（プレースホルダに置換）
        string temp = Regex.Replace(text, @"(?<=\p{IsBasicLatin}) (?=\p{IsBasicLatin})", "__SPACE__");

        // その他のスペース: 削除
        temp = temp.Replace(" ", "");

        // プレースホルダをスペースに戻す
        return temp.Replace("__SPACE__", " ");
    }

    static string GetSlideTitle(SlidePart slidePart, Config config)
    {
        // スライド内のタイトルテキストを探す
        var titleShape = slidePart.Slide.Descendants<Shape>()
            .FirstOrDefault(s =>
                TitleKeywords.Any(keyword =>
                    (s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.ToLower().Contains(keyword.ToLower()) ?? false)));

        if (titleShape != null)
        {
            var textBody = titleShape.TextBody;
            if (textBody == null)
            {
                return "（タイトルなし）";
            }

            // タイトルのテキストを取得
            var text = textBody.InnerText;
            if (string.IsNullOrWhiteSpace(text))
            {
                return "（タイトルなし）";
            }

            if (config.OnlyBiggerLetters)
            {
                text = "";
                double DeafaultFontSize = config.DefaultFontSize;

                var paragraphs = textBody.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>();
                var runs = textBody.Descendants<DocumentFormat.OpenXml.Drawing.Run>();
                //sort by font size
                // フォントサイズが取得できない場合は親要素（Paragraph, TextBody）から取得し、それでもなければデフォルト
                double GetFontSize(DocumentFormat.OpenXml.Drawing.Run run)
                {
                    // Run自身
                    var size = run.RunProperties?.FontSize?.Value;
                    if (size != null) return size.Value;

                    // Paragraph
                    var para = run.Ancestors<DocumentFormat.OpenXml.Drawing.Paragraph>().FirstOrDefault();
                    // ParagraphProperties の DefaultRunProperties から FontSize を取得
                    var paraProps = para?.ParagraphProperties;
                    if (paraProps != null)
                    {
                        var defaultRunProps = paraProps.GetFirstChild<DocumentFormat.OpenXml.Drawing.DefaultRunProperties>();
                        size = defaultRunProps?.FontSize?.Value;
                        if (size != null) return size.Value;
                    }
                    // TextBody
                    var textBody = run.Ancestors<DocumentFormat.OpenXml.Drawing.TextBody>().FirstOrDefault();
                    var paraProps2 = textBody?.Descendants<DocumentFormat.OpenXml.Drawing.ParagraphProperties>().FirstOrDefault();
                    if (paraProps2 != null)
                    {
                        var defaultRunProps2 = paraProps2.GetFirstChild<DocumentFormat.OpenXml.Drawing.DefaultRunProperties>();
                        size = defaultRunProps2?.FontSize?.Value;
                        if (size != null) return size.Value;
                    }

                    // デフォルト
                    return DeafaultFontSize;
                }

                var sortedRuns = runs.OrderByDescending(r => GetFontSize(r));
                // choose the biggest font size runs
                double maxFontSize = sortedRuns.FirstOrDefault()?.RunProperties?.FontSize?.Value ?? DeafaultFontSize;
                foreach (var run in sortedRuns)
                {
                    var fontSize = run.RunProperties?.FontSize?.Value ?? DeafaultFontSize;
                    if (fontSize == maxFontSize)
                    {
                        text += run.InnerText + " \n";
                    }
                }
            }

            // DeleteEmptyTitle
            if (config.DeleteEmptyTitle)
            {

                //タブと隣接するスペースを削除
                text = Regex.Replace(text, @"\s*\t\s*", "\t");
                //改行と隣接するスペースを削除
                text = Regex.Replace(text, @"\s*\r\n\s*", "\r\n");

                // タイトル内の半角・全角スペース・タブ・改行を削除
                text = Regex.Replace(text, @"[\t\r\n　]+", "");

                // アルファベット間以外のスペースを削除
                text = RemoveSpacesExceptBetweenAlphabets(text);

                text = text.Trim();
                if (string.IsNullOrWhiteSpace(text))
                {
                    return "（タイトルなし）";
                }
            }

            return text;
        }

        return "（タイトルなし）";
    }
}
