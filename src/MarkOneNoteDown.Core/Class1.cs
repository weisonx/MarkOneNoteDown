using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

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
            IEnumerable<string> texts = doc.Descendants(ns + "T")
                .Select(node => node.Value)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value.Trim());

            return string.Join(Environment.NewLine, texts);
        }
        catch
        {
            return string.Empty;
        }
    }
}
