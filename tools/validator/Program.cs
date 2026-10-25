// tools/validator/Program.cs  -- O1. `dotnet run -- file.pptx [Microsoft365]` prints the errors as JSON.
// Exit codes: 0 valid, 1 validation errors, 2 the package could not be opened.
// Not yet run: the machine this was written on has no .NET SDK. The harness compares the error count of a saved
// file with the original's, because some PowerPoint-written files have validator errors of their own.
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var path = args[0];
var version = args.Length > 1 ? Enum.Parse<FileFormatVersions>(args[1]) : FileFormatVersions.Microsoft365;
try
{
    using var doc = PresentationDocument.Open(path, false);
    // The default cap is 1000 errors; 0 removes it so error counts compare exactly.
    var validator = new OpenXmlValidator(version) { MaxNumberOfErrors = 0 };
    var errors = validator.Validate(doc)
        .Select(e => new { e.Id, e.Description, Part = e.Part?.Uri.ToString(), Path = e.Path?.XPath, Type = e.ErrorType.ToString() })
        .ToList();
    Console.WriteLine(JsonSerializer.Serialize(errors));
    return errors.Count == 0 ? 0 : 1;
}
catch (Exception e) when (e is OpenXmlPackageException or InvalidDataException or IOException)
{
    Console.WriteLine(JsonSerializer.Serialize(new { error = e.Message }));
    return 2;
}
