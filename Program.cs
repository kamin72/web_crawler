using AngleSharp;
using AngleSharp.Dom;

public class Program
{

    public static async Task Main()
    {
        var scraper = new WebScraper();


        var baseUrl = "https://ithelp.ithome.com.tw/questions";
        //var posts =  await scraper.GetPosts1(baseUrl);
        //foreach (var post in posts)
        //{
        //    Console.WriteLine($"Title: {post.Title}");
        //}


        var nextPageSelector = ".page > li > a";
        var remainingPage = 2;
        var outputPath = @"D:\posts.xlsx";


        var posts = await scraper.GetPosts(baseUrl, nextPageSelector, remainingPage);

        Console.WriteLine($"顯示資料筆數: {posts.Count()}");
        foreach (var post in posts)
        {
            Console.WriteLine($"Title: {post.Title}");
            Console.WriteLine($"Link: {post.Link}");
        }
        await scraper.ExportToExcel(
           baseUrl,
           nextPageSelector,
           remainingPage,
           outputPath
       );
    }
}
