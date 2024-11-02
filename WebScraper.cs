using OfficeOpenXml; // EPPlus
using System.Drawing;
using AngleSharp;
using AngleSharp.Dom;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Net.Http;
using System.Xml.Linq;
using OfficeOpenXml.Style;

public class Post
{
    public string Title { get; set; }
    public string Link { get; set; }
    public DateTime ScrapedDate { get; set; } = DateTime.Now;
}

public class WebScraper
{
    private readonly IBrowsingContext _browser;
    public WebScraper()
    {
        var config = Configuration.Default.WithDefaultLoader();
        _browser = BrowsingContext.New(config);
    }

    public async Task<List<Post>> GetPosts1(string baseUrl)
    {
        var document = await _browser.OpenAsync(baseUrl);
        var element = document.QuerySelectorAll("summary");
        var posts = element.Select(e => new Post
        {
            Title = e.TextContent
        }).Where(e => e.Title != null).ToList();

        return posts;
    }

    public async Task<IEnumerable<Post>> GetPosts(
    string baseUrl,
    string nextPageSelector,
    int remainingPages
    )
    {
        var currentUrl = baseUrl;
        var url = new Url(currentUrl);
        var document = await _browser.OpenAsync(url);

        var element = document.QuerySelectorAll("h3 > a");
        var posts = element.Select(e => new Post
        {
            Title = e.TextContent.Trim(),
            Link = e.GetAttribute("href")
        }).Where(e => e != null);


        if (!string.IsNullOrEmpty(nextPageSelector))
        {
            var nextPageElement = document.QuerySelector(nextPageSelector);
            if (nextPageElement != null)
            {
                var nextPageUrl = nextPageElement.GetAttribute("href");
                baseUrl = nextPageUrl;
                document.Close();
            }
        }
        else
        {
            return posts;
        }


        remainingPages--;
        if (remainingPages == 0)
        {
            return posts;
        }

        var nextPosts = await GetPosts(baseUrl, nextPageSelector, remainingPages);
        await Task.Delay(2000);
        return posts.Concat(nextPosts);
    }
    public async Task ExportToExcel(
        string baseUrl,
        string nextPageSelector,
        int pagesToCrawl,
        string outputPath
    )
    {
        // 設定 EPPlus 授權
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // 獲取爬蟲資料
        var posts = await GetPosts(baseUrl, nextPageSelector, pagesToCrawl);

        using (var package = new ExcelPackage())
        {
            // 建立工作表
            var worksheet = package.Workbook.Worksheets.Add("Posts");

            // 設定標題列樣式
            var headerStyle = worksheet.Cells["A1:C1"].Style;
            headerStyle.Font.Bold = true;
            headerStyle.Fill.PatternType = ExcelFillStyle.Solid;
            headerStyle.Fill.BackgroundColor.SetColor(Color.LightBlue);
            headerStyle.Border.Bottom.Style = ExcelBorderStyle.Medium;

            // 添加標題
            worksheet.Cells[1, 1].Value = "標題";
            worksheet.Cells[1, 2].Value = "連結";
            worksheet.Cells[1, 3].Value = "爬蟲時間";

            // 添加資料
            int row = 2;
            foreach (var post in posts)
            {
                worksheet.Cells[row, 1].Value = post.Title;
                worksheet.Cells[row, 2].Value = post.Link;
                worksheet.Cells[row, 3].Value = post.ScrapedDate;

                // 設定超連結
                worksheet.Cells[row, 2].Hyperlink = new Uri(post.Link);
                worksheet.Cells[row, 2].Style.Font.Color.SetColor(Color.Blue);
                worksheet.Cells[row, 2].Style.Font.UnderLine = true;

                row++;
            }

            // 自動調整欄寬
            worksheet.Cells.AutoFitColumns();

            // 設定日期格式
            worksheet.Column(3).Style.Numberformat.Format = "yyyy-mm-dd hh:mm:ss";

            // 儲存檔案
            await package.SaveAsAsync(new FileInfo(outputPath));
        }

        Console.WriteLine($"Excel file has been created: {outputPath}");
        Console.WriteLine($"Total posts exported: {posts.Count()}");
    }
}



