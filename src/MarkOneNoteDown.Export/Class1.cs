using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
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

public sealed class FileSystemExportWriter : IExportWriter
{
    private readonly string assetsFolderName;

    public FileSystemExportWriter(string assetsFolderName = "_assets")
    {
        this.assetsFolderName = string.IsNullOrWhiteSpace(assetsFolderName) ? "_assets" : assetsFolderName.Trim();
    }

    public async Task WriteAsync(MarkdownPage page, ExportOptions options, CancellationToken cancellationToken)
    {
        if (page is null)
        {
            throw new ArgumentNullException(nameof(page));
        }

        if (options is null)
        {
            throw new ArgumentNullException(nameof(options));
        }

        string outputDirectory = options.OutputDirectory;
        if (string.IsNullOrWhiteSpace(outputDirectory))
        {
            throw new InvalidOperationException("Output directory is required.");
        }

        Directory.CreateDirectory(outputDirectory);

        string safeTitle = MakeSafeFileName(page.Title);
        string path = EnsureUniquePath(Path.Combine(outputDirectory, safeTitle + ".md"));

        await File.WriteAllTextAsync(path, page.Markdown, new UTF8Encoding(false), cancellationToken);
        await CopyAssetsAsync(page.Assets, outputDirectory, cancellationToken);
    }

    private async Task CopyAssetsAsync(IReadOnlyList<AssetRef> assets, string outputDirectory, CancellationToken cancellationToken)
    {
        if (assets.Count == 0)
        {
            return;
        }

        string assetsDirectory = Path.Combine(outputDirectory, assetsFolderName);
        Directory.CreateDirectory(assetsDirectory);

        foreach (AssetRef asset in assets)
        {
            if (string.IsNullOrWhiteSpace(asset.SourceId) || !File.Exists(asset.SourceId))
            {
                continue;
            }

            string destination = Path.Combine(assetsDirectory, asset.FileName);
            if (File.Exists(destination))
            {
                continue;
            }

            using FileStream sourceStream = File.OpenRead(asset.SourceId);
            await using FileStream destStream = File.Create(destination);
            await sourceStream.CopyToAsync(destStream, cancellationToken);
        }
    }

    private static string MakeSafeFileName(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return "Untitled";
        }

        char[] invalid = Path.GetInvalidFileNameChars();
        var builder = new StringBuilder(input.Length);
        foreach (char ch in input)
        {
            builder.Append(invalid.Contains(ch) ? '_' : ch);
        }

        string result = builder.ToString().Trim();
        return string.IsNullOrWhiteSpace(result) ? "Untitled" : result;
    }

    private static string EnsureUniquePath(string path)
    {
        if (!File.Exists(path))
        {
            return path;
        }

        string directory = Path.GetDirectoryName(path) ?? string.Empty;
        string fileName = Path.GetFileNameWithoutExtension(path);
        string extension = Path.GetExtension(path);

        for (int i = 1; i < 1000; i++)
        {
            string candidate = Path.Combine(directory, $"{fileName} ({i}){extension}");
            if (!File.Exists(candidate))
            {
                return candidate;
            }
        }

        return Path.Combine(directory, $"{fileName} ({Guid.NewGuid():N}){extension}");
    }
}
