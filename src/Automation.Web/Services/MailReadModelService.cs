using System.Text.Json;
using Automation.Web.Models;

namespace Automation.Web.Services;

public sealed class MailReadModelService
{
    private readonly IConfiguration _configuration;
    private readonly IWebHostEnvironment _environment;

    public MailReadModelService(IConfiguration configuration, IWebHostEnvironment environment)
    {
        _configuration = configuration;
        _environment = environment;
    }

    public MailDashboardViewModel LoadDashboard()
    {
        var configuredPath = _configuration["MailMvp:ReadModelPath"] ?? "..\\..\\data\\mail_triage_read_model.json";
        var configuredIngestStatusPath = _configuration["MailMvp:IngestStatusPath"] ?? "..\\..\\data\\mail_ingest_status.json";
        var fullPath = Path.GetFullPath(Path.Combine(_environment.ContentRootPath, configuredPath));
        var fullIngestStatusPath = Path.GetFullPath(Path.Combine(_environment.ContentRootPath, configuredIngestStatusPath));

        var viewModel = new MailDashboardViewModel
        {
            DataPath = fullPath,
            IngestStatusPath = fullIngestStatusPath,
            HasData = false,
        };

        LoadHistoricalBacklog(viewModel, fullIngestStatusPath);

        if (!File.Exists(fullPath))
        {
            return viewModel;
        }

        try
        {
            var json = File.ReadAllText(fullPath);
            using var document = JsonDocument.Parse(json);
            var root = document.RootElement;

            viewModel.GeneratedUtc = root.TryGetProperty("generatedUtc", out var generatedUtc) ? generatedUtc.GetString() ?? "" : "";
            viewModel.HasData = true;

            if (root.TryGetProperty("summary", out var summary))
            {
                viewModel.Summary = new MailDashboardSummary
                {
                    TotalMessages = GetInt(summary, "totalMessages"),
                    NeedsDecision = GetInt(summary, "needsDecision"),
                    AutoReady = GetInt(summary, "autoReady"),
                    ManualReview = GetInt(summary, "manualReview"),
                    HighUrgency = GetInt(summary, "highUrgency"),
                };
            }

            if (root.TryGetProperty("needsDecision", out var items) && items.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in items.EnumerateArray())
                {
                    viewModel.NeedsDecision.Add(new MailDecisionItemViewModel
                    {
                        MessageKey = GetString(item, "messageKey"),
                        Subject = GetString(item, "subject"),
                        Sender = GetString(item, "sender"),
                        SenderDomain = GetString(item, "senderDomain"),
                        SourceFolderName = GetString(item, "sourceFolderName"),
                        SourceFolderPath = GetString(item, "sourceFolderPath"),
                        SourceStoreName = GetString(item, "sourceStoreName"),
                        ReceivedUtc = GetString(item, "receivedUtc"),
                        ProcessType = GetString(item, "processType"),
                        Confidence = GetDouble(item, "confidence"),
                        Urgency = GetString(item, "urgency"),
                        DecisionMode = GetString(item, "decisionMode"),
                        RecommendedAction = GetString(item, "recommendedAction"),
                        DecisionReason = GetString(item, "decisionReason"),
                        Status = GetString(item, "status"),
                        ServiceLevelState = GetString(item, "serviceLevelState"),
                        Owner = GetString(item, "owner"),
                        AttachmentCount = GetInt(item, "attachmentCount"),
                        AnalyzedAttachmentCount = GetInt(item, "analyzedAttachmentCount"),
                        AttachmentSummary = GetString(item, "attachmentSummary"),
                        AttachmentAnalysisPath = GetString(item, "attachmentAnalysisPath"),
                        AttachmentKinds = GetStringArray(item, "attachmentKinds"),
                        AttachmentAnalysis = GetAttachmentAnalysis(item, "attachmentAnalysis"),
                        Labels = GetStringArray(item, "labels"),
                        BodyPath = GetString(item, "bodyPath"),
                        SourcePath = GetString(item, "sourcePath"),
                        RawMessagePath = GetString(item, "rawMessagePath"),
                        WebLink = GetString(item, "webLink"),
                    });
                }
            }

            return viewModel;
        }
        catch
        {
            viewModel.HasData = false;
            return viewModel;
        }
    }

    private static void LoadHistoricalBacklog(MailDashboardViewModel viewModel, string fullIngestStatusPath)
    {
        if (!File.Exists(fullIngestStatusPath))
        {
            return;
        }

        try
        {
            var json = File.ReadAllText(fullIngestStatusPath);
            using var document = JsonDocument.Parse(json);
            var root = document.RootElement;

            if (!root.TryGetProperty("historicalBacklog", out var historicalBacklog))
            {
                return;
            }

            viewModel.HasHistoricalBacklog = true;
            viewModel.HistoricalBacklog = new HistoricalMailBacklogViewModel
            {
                GeneratedUtc = GetString(root, "generatedUtc"),
                Mailbox = GetString(root, "mailbox"),
                Folder = GetString(root, "folder"),
                CurrentYear = GetInt(root, "currentYear"),
                TotalHistoricalMessages = GetInt(historicalBacklog, "totalHistoricalMessages"),
                HistoricalUnread = GetInt(historicalBacklog, "historicalUnread"),
                HistoricalNotInRaw = GetInt(historicalBacklog, "historicalNotInRaw"),
                HistoricalUnreadNotInRaw = GetInt(historicalBacklog, "historicalUnreadNotInRaw"),
                HistoricalUnreadWithAttachments = GetInt(historicalBacklog, "historicalUnreadWithAttachments"),
                HistoricalUnreadAttachmentCount = GetInt(historicalBacklog, "historicalUnreadAttachmentCount"),
                HistoricalUnreadNotInRawWithAttachments = GetInt(historicalBacklog, "historicalUnreadNotInRawWithAttachments"),
                HistoricalUnreadNotInRawAttachmentCount = GetInt(historicalBacklog, "historicalUnreadNotInRawAttachmentCount"),
                OldestYear = GetInt(historicalBacklog, "oldestYear"),
            };

            if (root.TryGetProperty("historicalCatchup", out var catchup))
            {
                viewModel.HistoricalBacklog.Catchup = new HistoricalMailCatchupViewModel
                {
                    Enabled = GetBool(catchup, "enabled"),
                    LastRunUtc = GetString(catchup, "last_run_utc"),
                    LastProcessedCount = GetInt(catchup, "last_processed_count"),
                    LastSkippedExisting = GetInt(catchup, "last_skipped_existing"),
                    LastAttachmentCount = GetInt(catchup, "last_attachment_count"),
                    BatchSize = GetInt(catchup, "batch_size"),
                };
            }

            if (historicalBacklog.TryGetProperty("years", out var years) && years.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in years.EnumerateArray())
                {
                    viewModel.HistoricalBacklog.Years.Add(new HistoricalMailBacklogYearViewModel
                    {
                        Year = GetInt(item, "year"),
                        TotalMessages = GetInt(item, "totalMessages"),
                        UnreadMessages = GetInt(item, "unreadMessages"),
                        NotInRaw = GetInt(item, "notInRaw"),
                        UnreadNotInRaw = GetInt(item, "unreadNotInRaw"),
                    });
                }
            }
        }
        catch
        {
            viewModel.HasHistoricalBacklog = false;
        }
    }

    private static string GetString(JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property) ? property.GetString() ?? "" : "";
    }

    private static int GetInt(JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property) && property.TryGetInt32(out var value) ? value : 0;
    }

    private static double GetDouble(JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property) && property.TryGetDouble(out var value) ? value : 0;
    }

    private static bool GetBool(JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property)
            && property.ValueKind is JsonValueKind.True or JsonValueKind.False
            && property.GetBoolean();
    }

    private static List<string> GetStringArray(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var property) || property.ValueKind != JsonValueKind.Array)
        {
            return [];
        }

        var result = new List<string>();
        foreach (var item in property.EnumerateArray())
        {
            var value = item.GetString();
            if (!string.IsNullOrWhiteSpace(value))
            {
                result.Add(value);
            }
        }

        return result;
    }

    private static List<MailAttachmentInsightViewModel> GetAttachmentAnalysis(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var property) || property.ValueKind != JsonValueKind.Array)
        {
            return [];
        }

        var result = new List<MailAttachmentInsightViewModel>();
        foreach (var item in property.EnumerateArray())
        {
            result.Add(new MailAttachmentInsightViewModel
            {
                Name = GetString(item, "name"),
                StoredPath = GetString(item, "stored_path"),
                Kind = GetString(item, "kind"),
                ContentType = GetString(item, "content_type"),
                Extension = GetString(item, "extension"),
                Size = GetInt(item, "size"),
                Sha256 = GetString(item, "sha256"),
                Saved = GetBool(item, "saved"),
                AnalysisStatus = GetString(item, "analysis_status"),
                Summary = GetString(item, "summary"),
                Signals = GetStringArray(item, "signals"),
                WorksheetCount = GetInt(item, "worksheet_count"),
                PageCount = GetInt(item, "page_count"),
                ImageWidth = GetInt(item, "image_width"),
                ImageHeight = GetInt(item, "image_height"),
                TextExcerpt = GetString(item, "text_excerpt"),
            });
        }

        return result;
    }
}
