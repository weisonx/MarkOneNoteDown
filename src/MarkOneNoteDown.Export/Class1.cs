using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using MarkOneNoteDown.Core;
using MarkOneNoteDown.OneNote;

namespace MarkOneNoteDown.Export;

public sealed class ExportPipeline
{
    private readonly IOneNoteClient oneNoteClient;
    private readonly IPageParser parser;
    private readonly IExportWriter writer;

    public ExportPipeline(IOneNoteClient oneNoteClient, IPageParser parser, IExportWriter writer)
    {
        this.oneNoteClient = oneNoteClient ?? throw new ArgumentNullException(nameof(oneNoteClient));
        this.parser = parser ?? throw new ArgumentNullException(nameof(parser));
        this.writer = writer ?? throw new ArgumentNullException(nameof(writer));
    }

    public async Task<ExportResult> ExportPagesAsync(
        IReadOnlyList<PageRef> pages,
        ExportOptions options,
        IProgress<ExportProgress>? progress,
        CancellationToken cancellationToken)
    {
        if (pages is null)
        {
            throw new ArgumentNullException(nameof(pages));
        }

        if (options is null)
        {
            throw new ArgumentNullException(nameof(options));
        }

        int completed = 0;
        int exportedAssets = 0;
        int total = pages.Count;

        foreach (PageRef page in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();

            PageContent content = await oneNoteClient.GetPageContentAsync(page.Id, cancellationToken);
            MarkdownPage markdown = parser.Parse(content);

            exportedAssets += markdown.Assets.Count;
            await writer.WriteAsync(markdown, options, cancellationToken);

            completed++;
            progress?.Report(new ExportProgress(page.Name, completed, total));
        }

        return new ExportResult(completed, exportedAssets);
    }
}
