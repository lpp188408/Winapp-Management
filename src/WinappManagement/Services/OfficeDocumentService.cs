using System.IO;
using System.Runtime.InteropServices;

namespace WinappManagement.Services;

public sealed class OfficeDocumentService
{
    public IReadOnlyList<OfficeDocumentInfo> Snapshot()
    {
        var documents = new List<OfficeDocumentInfo>();
        TryAddDocuments(documents, "Word.Application", "WINWORD", "Documents");
        TryAddDocuments(documents, "Excel.Application", "EXCEL", "Workbooks");
        TryAddDocuments(documents, "PowerPoint.Application", "POWERPNT", "Presentations");
        TryAddProtectedViewDocuments(documents, "Word.Application", "WINWORD", "Document");
        TryAddProtectedViewDocuments(documents, "Excel.Application", "EXCEL", "Workbook");
        TryAddProtectedViewDocuments(documents, "PowerPoint.Application", "POWERPNT", "Presentation");
        return documents;
    }

    private static void TryAddDocuments(List<OfficeDocumentInfo> documents, string progId, string processName, string collectionName)
    {
        object? app = null;
        object? collection = null;
        try
        {
            app = GetActiveObject(progId);
            if (app is null)
            {
                return;
            }

            collection = GetProperty(app, collectionName);
            var count = ToInt(GetProperty(collection, "Count"));
            for (var index = 1; index <= count; index++)
            {
                object? document = null;
                try
                {
                    document = Invoke(collection, "Item", index);
                    var name = ToString(GetProperty(document, "Name"));
                    var fullName = ToString(GetProperty(document, "FullName"));
                    var directoryPath = ToString(GetProperty(document, "Path"));

                    if (string.IsNullOrWhiteSpace(fullName) || !Path.IsPathRooted(fullName))
                    {
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(directoryPath))
                    {
                        directoryPath = Path.GetDirectoryName(fullName) ?? string.Empty;
                    }

                    if (string.IsNullOrWhiteSpace(name))
                    {
                        name = Path.GetFileName(fullName);
                    }

                    AddDocument(documents, processName, name, fullName, directoryPath);
                }
                catch
                {
                    // Some Office windows, such as Protected View or unsaved files, can reject automation.
                }
                finally
                {
                    ReleaseComObject(document);
                }
            }
        }
        catch
        {
            // Office may be absent, elevated, busy, or not exposing an automation object.
        }
        finally
        {
            ReleaseComObject(collection);
            ReleaseComObject(app);
        }
    }

    private static void TryAddProtectedViewDocuments(List<OfficeDocumentInfo> documents, string progId, string processName, string documentPropertyName)
    {
        object? app = null;
        object? collection = null;
        try
        {
            app = GetActiveObject(progId);
            if (app is null)
            {
                return;
            }

            collection = GetProperty(app, "ProtectedViewWindows");
            var count = ToInt(GetProperty(collection, "Count"));
            for (var index = 1; index <= count; index++)
            {
                object? protectedWindow = null;
                object? document = null;
                try
                {
                    protectedWindow = Invoke(collection, "Item", index);
                    document = GetProperty(protectedWindow, documentPropertyName);
                    var name = ToString(GetProperty(document, "Name"));
                    var fullName = ToString(GetProperty(document, "FullName"));
                    var directoryPath = ToString(GetProperty(document, "Path"));

                    if (string.IsNullOrWhiteSpace(fullName) || !Path.IsPathRooted(fullName))
                    {
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(directoryPath))
                    {
                        directoryPath = Path.GetDirectoryName(fullName) ?? string.Empty;
                    }

                    if (string.IsNullOrWhiteSpace(name))
                    {
                        name = Path.GetFileName(fullName);
                    }

                    AddDocument(documents, processName, name, fullName, directoryPath);
                }
                catch
                {
                    // Protected View can reject individual document access.
                }
                finally
                {
                    ReleaseComObject(document);
                    ReleaseComObject(protectedWindow);
                }
            }
        }
        catch
        {
            // Some Office versions may not expose ProtectedViewWindows.
        }
        finally
        {
            ReleaseComObject(collection);
            ReleaseComObject(app);
        }
    }

    private static void AddDocument(List<OfficeDocumentInfo> documents, string processName, string name, string fullName, string directoryPath)
    {
        if (documents.Any(document =>
            document.ProcessName.Equals(processName, StringComparison.OrdinalIgnoreCase)
            && document.FullName.Equals(fullName, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        documents.Add(new OfficeDocumentInfo(processName, name, fullName, directoryPath));
    }

    private static object? GetActiveObject(string progId)
    {
        var clsidResult = NativeMethods.CLSIDFromProgID(progId, out var clsid);
        if (clsidResult != 0)
        {
            return null;
        }

        var result = NativeMethods.GetActiveObject(ref clsid, nint.Zero, out var activeObject);
        return result == 0 ? activeObject : null;
    }

    private static object? GetProperty(object? target, string propertyName)
    {
        return target?.GetType().InvokeMember(
            propertyName,
            System.Reflection.BindingFlags.GetProperty,
            binder: null,
            target,
            args: null);
    }

    private static object? Invoke(object? target, string methodName, params object[] args)
    {
        return target?.GetType().InvokeMember(
            methodName,
            System.Reflection.BindingFlags.InvokeMethod,
            binder: null,
            target,
            args);
    }

    private static int ToInt(object? value)
    {
        try
        {
            return Convert.ToInt32(value);
        }
        catch
        {
            return 0;
        }
    }

    private static string ToString(object? value)
    {
        return value?.ToString() ?? string.Empty;
    }

    private static void ReleaseComObject(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            _ = Marshal.FinalReleaseComObject(value);
        }
    }
}
