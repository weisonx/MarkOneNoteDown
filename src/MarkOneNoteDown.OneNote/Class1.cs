using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using MarkOneNoteDown.Core;
using UglyToad.PdfPig;

namespace MarkOneNoteDown.OneNote;

public interface IOneNoteClient
{
    Task<IReadOnlyList<NotebookRef>> GetNotebooksAsync(CancellationToken cancellationToken);

    Task<IReadOnlyList<SectionRef>> GetSectionsAsync(string notebookId, CancellationToken cancellationToken);

    Task<IReadOnlyList<PageRef>> GetPagesAsync(string sectionId, CancellationToken cancellationToken);

    Task<PageContent> GetPageContentAsync(string pageId, CancellationToken cancellationToken);

    Task<OneNoteDiagnostics> DiagnoseAsync(CancellationToken cancellationToken);
}

public sealed record OneNoteDiagnostics(
    bool CanCreateCom,
    string? Version,
    string? HierarchySample,
    string? ErrorMessage,
    int? HResult);

public sealed class HtmlExportClient : IOneNoteClient
{
    public Task<IReadOnlyList<NotebookRef>> GetNotebooksAsync(CancellationToken cancellationToken)
        => Task.FromResult<IReadOnlyList<NotebookRef>>(Array.Empty<NotebookRef>());

    public Task<IReadOnlyList<SectionRef>> GetSectionsAsync(string notebookId, CancellationToken cancellationToken)
        => Task.FromResult<IReadOnlyList<SectionRef>>(Array.Empty<SectionRef>());

    public Task<IReadOnlyList<PageRef>> GetPagesAsync(string sectionId, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(sectionId) || !Directory.Exists(sectionId))
        {
            return Task.FromResult<IReadOnlyList<PageRef>>(Array.Empty<PageRef>());
        }

        var pages = Directory.EnumerateFiles(sectionId, "*.*", SearchOption.TopDirectoryOnly)
            .Where(path => path.EndsWith(".html", StringComparison.OrdinalIgnoreCase) ||
                           path.EndsWith(".htm", StringComparison.OrdinalIgnoreCase) ||
                           path.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .Select(path => new PageRef(path, Path.GetFileNameWithoutExtension(path), sectionId))
            .ToList();

        return Task.FromResult<IReadOnlyList<PageRef>>(pages);
    }

    public async Task<PageContent> GetPageContentAsync(string pageId, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(pageId) || !File.Exists(pageId))
        {
            return new PageContent(pageId ?? string.Empty, "Untitled", string.Empty, ContentKind.PlainText);
        }

        string extension = Path.GetExtension(pageId);
        string title = Path.GetFileNameWithoutExtension(pageId);

        if (extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            string text = await Task.Run(() => ExtractPdfText(pageId), cancellationToken);
            return new PageContent(pageId, title, text, ContentKind.PdfText);
        }

        string html = await File.ReadAllTextAsync(pageId, cancellationToken);
        return new PageContent(pageId, title, html, ContentKind.Html);
    }

    public Task<OneNoteDiagnostics> DiagnoseAsync(CancellationToken cancellationToken)
        => Task.FromResult(new OneNoteDiagnostics(true, "HTML/PDF Export", null, null, null));

    private static string ExtractPdfText(string path)
    {
        using PdfDocument document = PdfDocument.Open(path);
        var builder = new System.Text.StringBuilder();
        foreach (var page in document.GetPages())
        {
            string text = page.Text;
            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }

            if (builder.Length > 0)
            {
                builder.AppendLine();
            }

            builder.AppendLine(text.Trim());
        }

        return builder.ToString().Trim();
    }
}
