using System.IO;

namespace WinappManagement.Services;

public sealed record OfficeDocumentInfo(
    string ProcessName,
    string Name,
    string FullName,
    string DirectoryPath,
    string WindowCaption = "",
    nint WindowHandle = 0)
{
    public string FileName => Path.GetFileName(FullName);
}
