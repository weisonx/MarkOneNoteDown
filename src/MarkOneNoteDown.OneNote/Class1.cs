using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using MarkOneNoteDown.Core;

namespace MarkOneNoteDown.OneNote;

public interface IOneNoteClient
{
    Task<IReadOnlyList<NotebookRef>> GetNotebooksAsync(CancellationToken cancellationToken);

    Task<IReadOnlyList<SectionRef>> GetSectionsAsync(string notebookId, CancellationToken cancellationToken);

    Task<IReadOnlyList<PageRef>> GetPagesAsync(string sectionId, CancellationToken cancellationToken);

    Task<PageContent> GetPageContentAsync(string pageId, CancellationToken cancellationToken);
}

public sealed class OneNoteClientStub : IOneNoteClient
{
    public Task<IReadOnlyList<NotebookRef>> GetNotebooksAsync(CancellationToken cancellationToken)
        => Task.FromResult<IReadOnlyList<NotebookRef>>(Array.Empty<NotebookRef>());

    public Task<IReadOnlyList<SectionRef>> GetSectionsAsync(string notebookId, CancellationToken cancellationToken)
        => Task.FromResult<IReadOnlyList<SectionRef>>(Array.Empty<SectionRef>());

    public Task<IReadOnlyList<PageRef>> GetPagesAsync(string sectionId, CancellationToken cancellationToken)
        => Task.FromResult<IReadOnlyList<PageRef>>(Array.Empty<PageRef>());

    public Task<PageContent> GetPageContentAsync(string pageId, CancellationToken cancellationToken)
        => Task.FromResult(new PageContent(pageId, "Untitled", string.Empty));
}

public sealed class OneNoteComClient : IOneNoteClient
{
    private const string OneNoteProgId = "OneNote.Application";

    public Task<IReadOnlyList<NotebookRef>> GetNotebooksAsync(CancellationToken cancellationToken)
        => RunStaAsync(() => GetNotebooksInternal(), cancellationToken);

    public Task<IReadOnlyList<SectionRef>> GetSectionsAsync(string notebookId, CancellationToken cancellationToken)
        => RunStaAsync(() => GetSectionsInternal(notebookId), cancellationToken);

    public Task<IReadOnlyList<PageRef>> GetPagesAsync(string sectionId, CancellationToken cancellationToken)
        => RunStaAsync(() => GetPagesInternal(sectionId), cancellationToken);

    public Task<PageContent> GetPageContentAsync(string pageId, CancellationToken cancellationToken)
        => RunStaAsync(() => GetPageContentInternal(pageId), cancellationToken);

    private static IReadOnlyList<NotebookRef> GetNotebooksInternal()
    {
        dynamic app = CreateOneNoteApplication();
        app.GetHierarchy(null, 0, out string xml);
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = doc.Root!.Name.Namespace;

        return doc.Descendants(ns + "Notebook")
            .Select(node => new NotebookRef(
                node.Attribute("ID")?.Value ?? string.Empty,
                node.Attribute("name")?.Value ?? "Untitled"))
            .ToList();
    }

    private static IReadOnlyList<SectionRef> GetSectionsInternal(string notebookId)
    {
        dynamic app = CreateOneNoteApplication();
        app.GetHierarchy(notebookId, 1, out string xml);
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = doc.Root!.Name.Namespace;

        return doc.Descendants(ns + "Section")
            .Select(node => new SectionRef(
                node.Attribute("ID")?.Value ?? string.Empty,
                node.Attribute("name")?.Value ?? "Untitled",
                notebookId))
            .ToList();
    }

    private static IReadOnlyList<PageRef> GetPagesInternal(string sectionId)
    {
        dynamic app = CreateOneNoteApplication();
        app.GetHierarchy(sectionId, 2, out string xml);
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = doc.Root!.Name.Namespace;

        return doc.Descendants(ns + "Page")
            .Select(node => new PageRef(
                node.Attribute("ID")?.Value ?? string.Empty,
                node.Attribute("name")?.Value ?? "Untitled",
                sectionId))
            .ToList();
    }

    private static PageContent GetPageContentInternal(string pageId)
    {
        dynamic app = CreateOneNoteApplication();
        app.GetPageContent(pageId, out string xml, 0);
        XDocument doc = XDocument.Parse(xml);
        XNamespace ns = doc.Root!.Name.Namespace;
        string title = doc.Descendants(ns + "Title").Descendants(ns + "Text").FirstOrDefault()?.Value ?? "Untitled";
        return new PageContent(pageId, title, xml);
    }

    private static dynamic CreateOneNoteApplication()
    {
        Type? type = Type.GetTypeFromProgID(OneNoteProgId, throwOnError: true);
        return Activator.CreateInstance(type!);
    }

    private static Task<T> RunStaAsync<T>(Func<T> action, CancellationToken cancellationToken)
    {
        if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
        {
            return Task.FromResult(action());
        }

        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        Thread thread = new Thread(() =>
        {
            try
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    tcs.TrySetCanceled(cancellationToken);
                    return;
                }

                T result = action();
                tcs.TrySetResult(result);
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return tcs.Task;
    }
}
