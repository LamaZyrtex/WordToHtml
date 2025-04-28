using Mammoth;
using Microsoft.AspNetCore.Mvc;
using HtmlAgilityPack;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;


namespace WordToHtml.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WordToHtml : ControllerBase
    {

        /// <summary>
        /// Clean a word document using Mammoth
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("clean-word")]
        [Consumes("multipart/form-data")]
        public  IActionResult CleanWord(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }

            using var stream = file.OpenReadStream();
            var styleMap = 
                @"" +
                "p[style-name='A01_Navn'] => p.A01_Navn\n" +
                "b => strong\n" +
                 "i => em\n" +
                 "u => u\n" +
                 "";

            var converter = new DocumentConverter().AddStyleMap(styleMap);

            var result = converter.ConvertToHtml(stream);

            return Content(result.Value, "text/html");
        }

        /// <summary>
        /// Clean word html using Html agility pack
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("clean-word-html")]
        [Consumes("multipart/form-data")]
        public async Task<IActionResult> CleanWordHtml(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }

            string htmlContent;

            using (var reader = new StreamReader(file.OpenReadStream(), Encoding.UTF8))
            {
                htmlContent = await reader.ReadToEndAsync();
            }

            var doc = new HtmlDocument();
            doc.LoadHtml(htmlContent);
            var htmlNode = doc.DocumentNode;

            // Remove inline styles
            RemoveInlineStyles(htmlNode);

            // Remove Mso classes
            RemoveMsoClasses(htmlNode);

            // Remove style tags
            RemoveStyleTags(htmlNode);

            // Remove word specific tags
            RemoveWordTags(htmlNode);

            // Clean attributes inside <html> tag
            CleanHtmlTagAttributes(htmlNode);

            // Remove all from <head> except <meta>
            RemoveFromHeadSection(htmlNode);

            // Remove all comments
            RemoveComments(htmlNode);

            // Remove empty nodes
            RemoveEmptyNodes(htmlNode);

            // Remove additional &nbsp;
            NormalizeNbsp(htmlNode);

            // Remove link attributes
            RemoveLinkAttributes(htmlNode);

            // Remove table attributes
            RemoveTableAttributes(htmlNode);

            // Ensure quoted attributes
            EnsureQuotedAttributes(htmlNode);

            var output = doc.DocumentNode.OuterHtml;

            return Content(output, "text/html; charset=utf-8");
        }

        private static void RemoveInlineStyles(HtmlNode node)
        {
            foreach (var singleNode in node.SelectNodes("//*[@style]") ?? Enumerable.Empty<HtmlNode>())
            {
                singleNode.Attributes["style"].Remove();
            }
        }

        private static void RemoveMsoClasses(HtmlNode node)
        {
            foreach (var singleNode in node.SelectNodes("//*[@class]") ?? Enumerable.Empty<HtmlNode>())
            {
                if (singleNode.GetAttributeValue("class", "").Contains("Mso"))
                {
                    singleNode.Attributes["class"].Remove();
                }
            }
        }

        private static void RemoveStyleTags(HtmlNode node)
        {
            foreach (var style in node.SelectNodes("//style") ?? Enumerable.Empty<HtmlNode>())
            {
                style.Remove();
            }
        }

        private static void RemoveWordTags(HtmlNode node)
        {
            foreach (var singleNode in node.Descendants().Where(n => n.Name.StartsWith("o:") || n.Name.StartsWith("v:") || n.Name == "xml" || n.Name == "w:worddocument").ToList())
            {
                singleNode.Remove();
            }
        }

        private static void CleanHtmlTagAttributes(HtmlNode node)
        {
            var htmlTag = node.SelectSingleNode("//html");
            if (htmlTag != null)
            {
                var allowedAttributes = new[] { "lang" };
                foreach (var attr in htmlTag.Attributes.Where(a => !allowedAttributes.Contains(a.Name)).ToList())
                {
                    htmlTag.Attributes.Remove(attr);
                }
            }
        }

        private static void RemoveFromHeadSection(HtmlNode node)
        {
            var head = node.SelectSingleNode("//head");
            if (head != null)
            {
                foreach (var child in head.ChildNodes.Where(n => n.Name != "meta").ToList())
                {
                    child.Remove();
                }
            }
        }

        private static void RemoveComments(HtmlNode node)
        {
            var comments = node.Descendants()
                .Where(n => n.NodeType == HtmlNodeType.Comment)
                .ToList();

            foreach (var comment in comments)
            {
                comment.Remove();
            }
        }

        private static void NormalizeNbsp(HtmlNode node)
        {
            foreach (var textNode in node.DescendantsAndSelf().Where(n => n.NodeType == HtmlNodeType.Text))
            {
                string updatedText = Regex.Replace(textNode.InnerHtml, @"(&nbsp;){2,}", "&nbsp;");
                textNode.InnerHtml = updatedText;
            }
        }

        private static void RemoveEmptyNodes(HtmlNode node)
        {
            // Lag en liste for å unngå endring i samling under iterasjon
            var children = node.ChildNodes.ToList();

            foreach (var child in children)
            {
                RemoveEmptyNodes(child); // Først rydde dypt

                // Sjekk om barnet nå er tomt
                if (IsNodeEmpty(child))
                {
                    child.Remove();
                }
            }
        }

        private static bool IsNodeEmpty(HtmlNode node)
        {
            // Ikke fjerne tekstnoder som kan ha innhold
            if (node.NodeType == HtmlNodeType.Text)
                return string.IsNullOrWhiteSpace(node.InnerText);

            // Fjern hvis ingen barnelementer og ingen ikke-whitespace-tekst
            return !node.HasChildNodes && string.IsNullOrWhiteSpace(node.InnerText);
        }

        private static void RemoveLinkAttributes(HtmlNode node)
        {
            foreach (var n in node.DescendantsAndSelf())
            {
                // Fjern 'link' og 'vlink' attributter hvis de finnes
                n.Attributes.Remove("link");
                n.Attributes.Remove("vlink");
            }
        }

        private static void RemoveTableAttributes(HtmlNode node)
        {
            foreach (var n in node.DescendantsAndSelf())
            {
                // Hvis det er en table: fjern width, border, cellspacing, cellpadding
                if (n.Name.Equals("table", StringComparison.OrdinalIgnoreCase))
                {
                    n.Attributes.Remove("width");
                    n.Attributes.Remove("border");
                    n.Attributes.Remove("cellspacing");
                    n.Attributes.Remove("cellpadding");
                }

                // Hvis det er en td: fjern width
                if (n.Name.Equals("td", StringComparison.OrdinalIgnoreCase))
                {
                    n.Attributes.Remove("width");
                }
            }
        }

        private static void EnsureQuotedAttributes(HtmlNode node)
        {
            foreach (var n in node.DescendantsAndSelf())
            {
                foreach (var attr in n.Attributes)
                {
                    // HtmlAgilityPack setter automatisk sitater når man skriver ut, men vi kan sikre verdiene ikke er null
                    if (string.IsNullOrEmpty(attr.QuoteType.ToString()))
                    {
                        attr.QuoteType = AttributeValueQuote.DoubleQuote;
                    }
                }
            }
        }
    }
}
