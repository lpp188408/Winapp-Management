using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

namespace WinappManagement.Services;

public sealed class OfficeDocumentService
{
    public IReadOnlyList<OfficeDocumentInfo> Snapshot()
    {
        var documents = new List<OfficeDocumentInfo>();
        TryAddDocuments(documents, "Word.Application", "WINWORD", "Documents");
        TryAddDocuments(documents, "Excel.Application", "EXCEL", "Workbooks");
        TryAddExcelWindowDocuments(documents);
        TryAddExcelDocumentsFromPowerShell(documents);
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

    private static void TryAddExcelWindowDocuments(List<OfficeDocumentInfo> documents)
    {
        object? app = null;
        object? collection = null;
        try
        {
            app = GetActiveObject("Excel.Application");
            if (app is null)
            {
                return;
            }

            collection = GetProperty(app, "Windows");
            var count = ToInt(GetProperty(collection, "Count"));
            for (var index = 1; index <= count; index++)
            {
                object? window = null;
                object? sheet = null;
                object? workbook = null;
                try
                {
                    window = Invoke(collection, "Item", index);
                    var caption = ToString(GetProperty(window, "Caption"));
                    var hwnd = ToIntPtr(GetProperty(window, "Hwnd"));

                    sheet = GetProperty(window, "ActiveSheet");
                    workbook = GetProperty(sheet, "Parent");
                    var name = ToString(GetProperty(workbook, "Name"));
                    var fullName = ToString(GetProperty(workbook, "FullName"));
                    var directoryPath = ToString(GetProperty(workbook, "Path"));

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

                    AddDocument(documents, "EXCEL", name, fullName, directoryPath, caption, hwnd);
                }
                catch
                {
                    // Excel can expose chart windows, protected views, or add-in windows that do not map cleanly to a workbook.
                }
                finally
                {
                    ReleaseComObject(workbook);
                    ReleaseComObject(sheet);
                    ReleaseComObject(window);
                }
            }
        }
        catch
        {
            // Excel may be absent, elevated, busy, or not exposing the Windows collection.
        }
        finally
        {
            ReleaseComObject(collection);
            ReleaseComObject(app);
        }
    }

    private static void TryAddExcelDocumentsFromPowerShell(List<OfficeDocumentInfo> documents)
    {
        try
        {
            var script = """
                [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
                $items = @()
                try {
                    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
                    for ($i = 1; $i -le $excel.Windows.Count; $i++) {
                        try {
                            $window = $excel.Windows.Item($i)
                            $workbook = $window.ActiveSheet.Parent
                            if ($workbook.FullName -and [System.IO.Path]::IsPathRooted([string]$workbook.FullName)) {
                                $items += [PSCustomObject]@{
                                    Name = [string]$workbook.Name
                                    FullName = [string]$workbook.FullName
                                    DirectoryPath = [string]$workbook.Path
                                    WindowCaption = [string]$window.Caption
                                    WindowHandle = [string]$window.Hwnd
                                }
                            }
                        } catch {}
                    }
                } catch {}
                $items | ConvertTo-Json -Compress -Depth 3
                """;

            using var process = new Process();
            process.StartInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command {QuotePowerShellArgument(script)}",
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            process.Start();
            if (!process.WaitForExit(3000))
            {
                try
                {
                    process.Kill(entireProcessTree: true);
                }
                catch
                {
                    // Ignore cleanup failures.
                }

                return;
            }

            var output = process.StandardOutput.ReadToEnd();
            _ = process.StandardError.ReadToEnd();
            if (process.ExitCode != 0 || string.IsNullOrWhiteSpace(output))
            {
                return;
            }

            AddExcelDocumentsFromJson(documents, output);
        }
        catch
        {
            // This is a fallback for machines where direct Excel COM reflection fails.
        }
    }

    private static string QuotePowerShellArgument(string script)
    {
        return "\"" + script.Replace("\"", "\\\"").Replace("\r", " ").Replace("\n", " ") + "\"";
    }

    private static void AddExcelDocumentsFromJson(List<OfficeDocumentInfo> documents, string json)
    {
        using var payload = JsonDocument.Parse(json);
        if (payload.RootElement.ValueKind == JsonValueKind.Array)
        {
            foreach (var element in payload.RootElement.EnumerateArray())
            {
                AddExcelDocumentFromJson(documents, element);
            }
            return;
        }

        if (payload.RootElement.ValueKind == JsonValueKind.Object)
        {
            AddExcelDocumentFromJson(documents, payload.RootElement);
        }
    }

    private static void AddExcelDocumentFromJson(List<OfficeDocumentInfo> documents, JsonElement element)
    {
        var fullName = JsonString(element, "FullName");
        if (string.IsNullOrWhiteSpace(fullName) || !Path.IsPathRooted(fullName))
        {
            return;
        }

        var directoryPath = JsonString(element, "DirectoryPath");
        if (string.IsNullOrWhiteSpace(directoryPath))
        {
            directoryPath = Path.GetDirectoryName(fullName) ?? string.Empty;
        }

        var name = JsonString(element, "Name");
        if (string.IsNullOrWhiteSpace(name))
        {
            name = Path.GetFileName(fullName);
        }

        AddDocument(
            documents,
            "EXCEL",
            name,
            fullName,
            directoryPath,
            JsonString(element, "WindowCaption"),
            ToIntPtr(JsonString(element, "WindowHandle")));
    }

    private static string JsonString(JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property) && property.ValueKind != JsonValueKind.Null
            ? property.ToString()
            : string.Empty;
    }

    private static void AddDocument(List<OfficeDocumentInfo> documents, string processName, string name, string fullName, string directoryPath, string windowCaption = "", nint windowHandle = 0)
    {
        var existingIndex = documents.FindIndex(document =>
            document.ProcessName.Equals(processName, StringComparison.OrdinalIgnoreCase)
            && document.FullName.Equals(fullName, StringComparison.OrdinalIgnoreCase));
        if (existingIndex >= 0)
        {
            var existing = documents[existingIndex];
            if ((existing.WindowHandle == 0 && windowHandle != 0)
                || (string.IsNullOrWhiteSpace(existing.WindowCaption) && !string.IsNullOrWhiteSpace(windowCaption)))
            {
                documents[existingIndex] = existing with
                {
                    WindowCaption = string.IsNullOrWhiteSpace(existing.WindowCaption) ? windowCaption : existing.WindowCaption,
                    WindowHandle = existing.WindowHandle == 0 ? windowHandle : existing.WindowHandle
                };
            }

            return;
        }

        documents.Add(new OfficeDocumentInfo(processName, name, fullName, directoryPath, windowCaption, windowHandle));
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

    private static nint ToIntPtr(object? value)
    {
        try
        {
            return (nint)Convert.ToInt64(value);
        }
        catch
        {
            return 0;
        }
    }

    private static void ReleaseComObject(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            _ = Marshal.FinalReleaseComObject(value);
        }
    }
}
