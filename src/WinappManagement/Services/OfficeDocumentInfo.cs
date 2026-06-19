using System.IO;

namespace WinappManagement.Services;

public sealed record OfficeDocumentInfo(
    string ProcessName,
    string Name,
    string FullName,
    string DirectoryPath)
{
    public string FileName => Path.GetFileName(FullName);
}
