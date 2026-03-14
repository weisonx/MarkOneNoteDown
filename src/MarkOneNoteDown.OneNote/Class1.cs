using System;
using System.Collections.Generic;
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
