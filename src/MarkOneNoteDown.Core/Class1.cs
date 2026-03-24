using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using HtmlAgilityPack;

namespace MarkOneNoteDown.Core;

public sealed record NotebookRef(string Id, string Name);

public sealed record SectionRef(string Id, string Name, string NotebookId);

public sealed record PageRef(string Id, string Name, string SectionId);

public sealed record PageContent(string Id, string Title, string RawContent);

public sealed record AssetRef(string FileName, string SourceId);

public sealed record MarkdownPage(string Title, string Markdown, IReadOnlyList<AssetRef> Assets);

public sealed record ExportOptions(string OutputDirectory, bool IncludeAttachments, bool FlattenHierarchy);

public sealed record ExportProgress(string CurrentItem, int Completed, int Total);

public sealed record ExportResult(int ExportedPages, int ExportedAssets);

public interface IPageParser
{
    MarkdownPage Parse(PageContent content);
}

public interface IExportWriter
{
    Task WriteAsync(MarkdownPage page, ExportOptions options, CancellationToken cancellationToken);
}

public sealed class BasicPageParser : IPageParser
{
    public MarkdownPage Parse(PageContent content)
    {
        if (content is null)
        {
            throw new ArgumentNullException(nameof(content));
        }

        string title = string.IsNullOrWhiteSpace(content.Title) ? "Untitled" : content.Title.Trim();
        string body = ExtractPlainText(content.RawContent);

        if (string.IsNullOrWhiteSpace(body))
        {
            body = "_(Empty page or unsupported content.)_";
        }

        string markdown = $"# {title}\n\n{body}\n";
        return new MarkdownPage(title, markdown, Array.Empty<AssetRef>());
    }

    private static string ExtractPlainText(string rawContent)
    {
        if (string.IsNullOrWhiteSpace(rawContent))
        {
            return string.Empty;
        }

        try
        {
            XDocument doc = XDocument.Parse(rawContent);
            XNamespace ns = doc.Root?.Name.Namespace ?? string.Empty;

            var builder = new StringBuilder();
            foreach (XElement outline in doc.Descendants(ns + "Outline"))
            {
                foreach (XElement element in outline.Descendants(ns + "OE"))
                {
                    IEnumerable<string> texts = element.Descendants(ns + "T")
                        .Select(node => node.Value)
                        .Where(value => !string.IsNullOrWhiteSpace(value))
                        .Select(value => value.Trim());

                    string paragraph = string.Join(" ", texts);
                    if (string.IsNullOrWhiteSpace(paragraph))
                    {
                        continue;
                    }

                    if (builder.Length > 0)
                    {
                        builder.AppendLine();
                    }

                    builder.AppendLine(paragraph);
                }
            }

            return builder.ToString().TrimEnd();
        }
        catch
        {
            return string.Empty;
        }
    }
}

public sealed class HtmlPageParser : IPageParser
{
    public MarkdownPage Parse(PageContent content)
    {
        if (content is null)
        {
            throw new ArgumentNullException(nameof(content));
        }

        string title = string.IsNullOrWhiteSpace(content.Title) ? "Untitled" : content.Title.Trim();
        string body = ExtractHtmlText(content.RawContent);

        if (string.IsNullOrWhiteSpace(body))
        {
            body = "_(Empty page or unsupported content.)_";
        }

        string markdown = $"# {title}\n\n{body}\n";
        return new MarkdownPage(title, markdown, Array.Empty<AssetRef>());
    }

    private static string ExtractHtmlText(string html)
    {
        if (string.IsNullOrWhiteSpace(html))
        {
            return string.Empty;
        }

        var doc = new HtmlDocument();
        doc.LoadHtml(html);

        string title = doc.DocumentNode.SelectSingleNode("//title")?.InnerText?.Trim() ?? string.Empty;
        HtmlNode? bodyNode = doc.DocumentNode.SelectSingleNode("//body") ?? doc.DocumentNode;
        string text = NormalizeWhitespace(bodyNode.InnerText);

        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        if (!string.IsNullOrWhiteSpace(title) && text.StartsWith(title, StringComparison.OrdinalIgnoreCase))
        {
            text = text.Substring(title.Length).TrimStart();
        }

        return text;
    }

    private static string NormalizeWhitespace(string value)
    {
        var builder = new StringBuilder(value.Length);
        bool previousWasWhitespace = false;
        foreach (char ch in value)
        {
            if (char.IsWhiteSpace(ch))
            {
                if (!previousWasWhitespace)
                {
                    builder.Append(' ');
                    previousWasWhitespace = true;
                }
                continue;
            }

            previousWasWhitespace = false;
            builder.Append(ch);
        }

        return builder.ToString().Trim();
    }
}
