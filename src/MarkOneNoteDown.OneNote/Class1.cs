using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using MarkOneNoteDown.Core;

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
                           path.EndsWith(".htm", StringComparison.OrdinalIgnoreCase))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .Select(path => new PageRef(path, Path.GetFileNameWithoutExtension(path), sectionId))
            .ToList();

        return Task.FromResult<IReadOnlyList<PageRef>>(pages);
    }

    public async Task<PageContent> GetPageContentAsync(string pageId, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(pageId) || !File.Exists(pageId))
        {
            return new PageContent(pageId ?? string.Empty, "Untitled", string.Empty);
        }

        string html = await File.ReadAllTextAsync(pageId, cancellationToken);
        string title = Path.GetFileNameWithoutExtension(pageId);
        return new PageContent(pageId, title, html);
    }

    public Task<OneNoteDiagnostics> DiagnoseAsync(CancellationToken cancellationToken)
        => Task.FromResult(new OneNoteDiagnostics(true, "HTML Export", null, null, null));
}
