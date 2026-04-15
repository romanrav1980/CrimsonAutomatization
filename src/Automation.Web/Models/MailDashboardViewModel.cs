namespace Automation.Web.Models;

public sealed class MailDashboardViewModel
{
    public string GeneratedUtc { get; set; } = "";
    public string DataPath { get; set; } = "";
    public string IngestStatusPath { get; set; } = "";
    public bool HasData { get; set; }
    public bool HasHistoricalBacklog { get; set; }
    public string FlashMessage { get; set; } = "";
    public string FlashKind { get; set; } = "";
    public MailDashboardSummary Summary { get; set; } = new();
    public HistoricalMailBacklogViewModel HistoricalBacklog { get; set; } = new();
    public List<MailDecisionItemViewModel> NeedsDecision { get; set; } = [];
}

public sealed class MailDashboardSummary
{
    public int TotalMessages { get; set; }
    public int NeedsDecision { get; set; }
    public int AutoReady { get; set; }
    public int ManualReview { get; set; }
    public int HighUrgency { get; set; }
}

public sealed class MailDecisionItemViewModel
{
    public string MessageKey { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Sender { get; set; } = "";
    public string SenderDomain { get; set; } = "";
    public string SourceFolderName { get; set; } = "";
    public string SourceFolderPath { get; set; } = "";
    public string SourceStoreName { get; set; } = "";
    public string ReceivedUtc { get; set; } = "";
    public string ProcessType { get; set; } = "";
    public double Confidence { get; set; }
    public string Urgency { get; set; } = "";
    public string DecisionMode { get; set; } = "";
    public string RecommendedAction { get; set; } = "";
    public string DecisionReason { get; set; } = "";
    public string Status { get; set; } = "";
    public string ServiceLevelState { get; set; } = "";
    public string Owner { get; set; } = "";
    public int AttachmentCount { get; set; }
    public int AnalyzedAttachmentCount { get; set; }
    public string AttachmentSummary { get; set; } = "";
    public string AttachmentAnalysisPath { get; set; } = "";
    public List<string> AttachmentKinds { get; set; } = [];
    public List<MailAttachmentInsightViewModel> AttachmentAnalysis { get; set; } = [];
    public List<string> Labels { get; set; } = [];
    public List<MailAuditEventViewModel> RecentAuditEvents { get; set; } = [];
    public string BodyPath { get; set; } = "";
    public string SourcePath { get; set; } = "";
    public string RawMessagePath { get; set; } = "";
    public string WebLink { get; set; } = "";
}

public sealed class MailAuditEventViewModel
{
    public string EventType { get; set; } = "";
    public string CreatedUtc { get; set; } = "";
    public string Actor { get; set; } = "";
    public string Action { get; set; } = "";
    public string Notes { get; set; } = "";
    public string Summary { get; set; } = "";
}

public sealed class MailAttachmentInsightViewModel
{
    public string Name { get; set; } = "";
    public string StoredPath { get; set; } = "";
    public string Kind { get; set; } = "";
    public string ContentType { get; set; } = "";
    public string Extension { get; set; } = "";
    public int Size { get; set; }
    public string Sha256 { get; set; } = "";
    public bool Saved { get; set; }
    public string AnalysisStatus { get; set; } = "";
    public string Summary { get; set; } = "";
    public List<string> Signals { get; set; } = [];
    public int WorksheetCount { get; set; }
    public int PageCount { get; set; }
    public int ImageWidth { get; set; }
    public int ImageHeight { get; set; }
    public string TextExcerpt { get; set; } = "";
}

public sealed class HistoricalMailBacklogViewModel
{
    public string GeneratedUtc { get; set; } = "";
    public string Mailbox { get; set; } = "";
    public string Folder { get; set; } = "";
    public int CurrentYear { get; set; }
    public int TotalHistoricalMessages { get; set; }
    public int HistoricalUnread { get; set; }
    public int HistoricalNotInRaw { get; set; }
    public int HistoricalUnreadNotInRaw { get; set; }
    public int HistoricalUnreadWithAttachments { get; set; }
    public int HistoricalUnreadAttachmentCount { get; set; }
    public int HistoricalUnreadNotInRawWithAttachments { get; set; }
    public int HistoricalUnreadNotInRawAttachmentCount { get; set; }
    public int OldestYear { get; set; }
    public HistoricalMailCatchupViewModel Catchup { get; set; } = new();
    public List<HistoricalMailBacklogYearViewModel> Years { get; set; } = [];
}

public sealed class HistoricalMailBacklogYearViewModel
{
    public int Year { get; set; }
    public int TotalMessages { get; set; }
    public int UnreadMessages { get; set; }
    public int NotInRaw { get; set; }
    public int UnreadNotInRaw { get; set; }
}

public sealed class HistoricalMailCatchupViewModel
{
    public bool Enabled { get; set; }
    public string LastRunUtc { get; set; } = "";
    public int LastProcessedCount { get; set; }
    public int LastSkippedExisting { get; set; }
    public int LastAttachmentCount { get; set; }
    public int BatchSize { get; set; }
}
