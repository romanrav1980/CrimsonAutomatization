using System.Diagnostics;
using System.Text.Json;

namespace Automation.Web.Services;

public sealed class MailOperatorActionService
{
    private readonly IConfiguration _configuration;
    private readonly IWebHostEnvironment _environment;

    public MailOperatorActionService(IConfiguration configuration, IWebHostEnvironment environment)
    {
        _configuration = configuration;
        _environment = environment;
    }

    public async Task<MailOperatorActionResult> ApplyActionAsync(
        string messageKey,
        string action,
        string owner,
        string notes,
        string actor,
        CancellationToken cancellationToken = default)
    {
        var repoRoot = Path.GetFullPath(Path.Combine(_environment.ContentRootPath, "..", ".."));
        var pythonExe = _configuration["MailMvp:PythonExe"] ?? "python";
        var scriptPath = Path.Combine(repoRoot, "scripts", "apply_mail_action.py");
        var dbPath = ResolveFromRepoRoot(_configuration["MailMvp:DbPath"] ?? "data/automation.db", repoRoot);
        var readModelPath = ResolveFromRepoRoot(_configuration["MailMvp:ReadModelPath"] ?? "data/mail_triage_read_model.json", repoRoot);

        var startInfo = new ProcessStartInfo
        {
            FileName = pythonExe,
            WorkingDirectory = repoRoot,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
        };
        startInfo.ArgumentList.Add(scriptPath);
        startInfo.ArgumentList.Add("--message-key");
        startInfo.ArgumentList.Add(messageKey);
        startInfo.ArgumentList.Add("--action");
        startInfo.ArgumentList.Add(action);
        startInfo.ArgumentList.Add("--owner");
        startInfo.ArgumentList.Add(owner ?? "");
        startInfo.ArgumentList.Add("--notes");
        startInfo.ArgumentList.Add(notes ?? "");
        startInfo.ArgumentList.Add("--actor");
        startInfo.ArgumentList.Add(actor ?? "");
        startInfo.ArgumentList.Add("--db-path");
        startInfo.ArgumentList.Add(dbPath);
        startInfo.ArgumentList.Add("--read-model-path");
        startInfo.ArgumentList.Add(readModelPath);

        using var process = new Process { StartInfo = startInfo };
        process.Start();

        var stdoutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var stderrTask = process.StandardError.ReadToEndAsync(cancellationToken);
        await process.WaitForExitAsync(cancellationToken);

        var stdout = await stdoutTask;
        var stderr = await stderrTask;

        if (process.ExitCode != 0)
        {
            return new MailOperatorActionResult
            {
                Succeeded = false,
                Message = string.IsNullOrWhiteSpace(stderr) ? "Mail action failed." : stderr.Trim(),
            };
        }

        try
        {
            using var document = JsonDocument.Parse(stdout);
            var root = document.RootElement;
            return new MailOperatorActionResult
            {
                Succeeded = true,
                Message = BuildSuccessMessage(
                    GetString(root, "action"),
                    GetString(root, "messageKey"),
                    GetString(root, "owner"),
                    GetString(root, "status")
                ),
            };
        }
        catch
        {
            return new MailOperatorActionResult
            {
                Succeeded = true,
                Message = "Mail action was applied successfully.",
            };
        }
    }

    private static string BuildSuccessMessage(string action, string messageKey, string owner, string status)
    {
        var normalizedAction = action switch
        {
            "approve" => "Suggestion approved",
            "archive" => "Mail item archived",
            "manual" => "Item marked for manual review",
            "assign_owner" => $"Owner assigned to {owner}",
            _ => "Mail action applied",
        };

        var statusSuffix = string.IsNullOrWhiteSpace(status) ? "" : $" Status: {status}.";
        return $"{normalizedAction} for {messageKey}.{statusSuffix}";
    }

    private static string GetString(JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property) ? property.GetString() ?? "" : "";
    }

    private static string ResolveFromRepoRoot(string configuredPath, string repoRoot)
    {
        return Path.IsPathRooted(configuredPath)
            ? configuredPath
            : Path.GetFullPath(Path.Combine(repoRoot, configuredPath));
    }
}

public sealed class MailOperatorActionResult
{
    public bool Succeeded { get; set; }
    public string Message { get; set; } = "";
}
