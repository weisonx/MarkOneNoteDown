using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using MarkOneNoteDown.Core;

namespace MarkOneNoteDown.OneNote
{
    public interface IOneNoteClient
    {
        Task<IReadOnlyList<NotebookRef>> GetNotebooksAsync(CancellationToken cancellationToken);
        Task<IReadOnlyList<SectionRef>> GetSectionsAsync(string notebookId, CancellationToken cancellationToken);
        Task<IReadOnlyList<PageRef>> GetPagesAsync(string sectionId, CancellationToken cancellationToken);
        Task<PageContent> GetPageContentAsync(string pageId, CancellationToken cancellationToken);
        Task UpdatePageContentAsync(string pageId, Func<XDocument, XDocument> modifyXml, CancellationToken cancellationToken);
        Task<OneNoteDiagnostics> DiagnoseAsync(CancellationToken cancellationToken);
    }

    public sealed record OneNoteDiagnostics(
        bool CanCreateCom,
        string? Version,
        string? HierarchySample,
        string? ErrorMessage,
        int? HResult);

    public sealed class OneNoteComClient : IOneNoteClient
    {
        private const string OneNoteProgId = "OneNote.Application";

        public Task<IReadOnlyList<NotebookRef>> GetNotebooksAsync(CancellationToken cancellationToken)
            => RunStaAsync<IReadOnlyList<NotebookRef>>(() => GetNotebooksInternal(), cancellationToken);

        public Task<IReadOnlyList<SectionRef>> GetSectionsAsync(string notebookId, CancellationToken cancellationToken)
            => RunStaAsync<IReadOnlyList<SectionRef>>(() => GetSectionsInternal(notebookId), cancellationToken);

        public Task<IReadOnlyList<PageRef>> GetPagesAsync(string sectionId, CancellationToken cancellationToken)
            => RunStaAsync<IReadOnlyList<PageRef>>(() => GetPagesInternal(sectionId), cancellationToken);

        public Task<PageContent> GetPageContentAsync(string pageId, CancellationToken cancellationToken)
            => RunStaAsync<PageContent>(() => GetPageContentInternal(pageId), cancellationToken);

        public Task UpdatePageContentAsync(string pageId, Func<XDocument, XDocument> modifyXml, CancellationToken cancellationToken)
            => RunStaAsync(() => UpdatePageContentInternal(pageId, modifyXml), cancellationToken);

        public Task<OneNoteDiagnostics> DiagnoseAsync(CancellationToken cancellationToken)
            => RunStaAsync<OneNoteDiagnostics>(() => DiagnoseInternal(), cancellationToken);

        #region Internal Methods

        private static IReadOnlyList<NotebookRef> GetNotebooksInternal()
        {
            dynamic app = CreateOneNoteApplication();
            try
            {
                app.GetHierarchy(null, 0, out string xml);
                XDocument doc = XDocument.Parse(xml);
                XNamespace ns = doc.Root!.Name.Namespace;

                return doc.Descendants(ns + "Notebook")
                    .Select(node => new NotebookRef(
                        node.Attribute("ID")?.Value ?? string.Empty,
                        node.Attribute("name")?.Value ?? "Untitled"))
                    .ToList();
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException($"OneNote COM failed (0x{ex.HResult:X}): {ex.Message}", ex);
            }
        }

        private static IReadOnlyList<SectionRef> GetSectionsInternal(string notebookId)
        {
            dynamic app = CreateOneNoteApplication();
            try
            {
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
            catch (COMException ex)
            {
                throw new InvalidOperationException($"OneNote COM failed (0x{ex.HResult:X}): {ex.Message}", ex);
            }
        }

        private static IReadOnlyList<PageRef> GetPagesInternal(string sectionId)
        {
            dynamic app = CreateOneNoteApplication();
            try
            {
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
            catch (COMException ex)
            {
                throw new InvalidOperationException($"OneNote COM failed (0x{ex.HResult:X}): {ex.Message}", ex);
            }
        }

        private static PageContent GetPageContentInternal(string pageId)
        {
            dynamic app = CreateOneNoteApplication();
            try
            {
                app.GetPageContent(pageId, out string xml, 0);
                XDocument doc = XDocument.Parse(xml);
                XNamespace ns = doc.Root!.Name.Namespace;

                string title = doc.Descendants(ns + "Title")
                                  .Descendants(ns + "Text")
                                  .FirstOrDefault()?.Value ?? "Untitled";

                return new PageContent(pageId, title, xml);
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException($"OneNote COM failed (0x{ex.HResult:X}): {ex.Message}", ex);
            }
        }

        private static void UpdatePageContentInternal(string pageId, Func<XDocument, XDocument> modifyXml)
        {
            dynamic app = CreateOneNoteApplication();
            try
            {
                app.GetPageContent(pageId, out string xml, 0);
                XDocument doc = XDocument.Parse(xml);

                XDocument newDoc = modifyXml(doc);

                if (newDoc.Root?.Name.NamespaceName == "")
                    newDoc.Root.Name = doc.Root.Name;

                string newXml = newDoc.ToString(SaveOptions.DisableFormatting);
                app.UpdatePageContent(newXml);
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException($"OneNote COM failed (0x{ex.HResult:X}): {ex.Message}", ex);
            }
        }

        private static OneNoteDiagnostics DiagnoseInternal()
        {
            try
            {
                dynamic app = CreateOneNoteApplication();
                string version = SafeGetVersion(app);
                app.GetHierarchy(null, 0, out string xml);
                string sample = xml.Length > 200 ? xml.Substring(0, 200) : xml;
                return new OneNoteDiagnostics(true, version, sample, null, null);
            }
            catch (COMException comEx)
            {
                return new OneNoteDiagnostics(false, null, null, comEx.Message, comEx.HResult);
            }
            catch (Exception ex)
            {
                return new OneNoteDiagnostics(false, null, null, ex.Message, ex.HResult);
            }
        }

        private static string SafeGetVersion(dynamic app)
        {
            try
            {
                return app.GetVersion();
            }
            catch
            {
                return "Unknown";
            }
        }

        #endregion

        #region COM Helpers

        private static dynamic CreateOneNoteApplication()
        {
            try
            {
                Type? type = Type.GetTypeFromProgID(OneNoteProgId, throwOnError: false);
                if (type is null)
                    throw new InvalidOperationException(BuildOneNoteNotFoundMessage());

                return Activator.CreateInstance(type) ?? throw new InvalidOperationException(BuildOneNoteNotFoundMessage());
            }
            catch (Exception ex) when (ex is not InvalidOperationException)
            {
                throw new InvalidOperationException(BuildOneNoteNotFoundMessage(), ex);
            }
        }

        private static string BuildOneNoteNotFoundMessage()
        {
            return "Failed to load OneNote COM automation. " +
                   "Please ensure OneNote desktop (Microsoft 365 or OneNote 2016) is installed " +
                   "and launched at least once. The Microsoft Store version (OneNote for Windows 10) does not provide COM automation.";
        }

        private static Task<T> RunStaAsync<T>(Func<T> action, CancellationToken cancellationToken)
        {
            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
                return Task.FromResult(action());

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

        // 非泛型版本，用于 Task 返回的 UpdatePageContentAsync
        private static Task RunStaAsync(Action action, CancellationToken cancellationToken)
        {
            return RunStaAsync<object>(() =>
            {
                action();
                return null!;
            }, cancellationToken);
        }

        #endregion
    }
}