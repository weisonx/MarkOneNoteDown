using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

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
