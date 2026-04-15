using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using Automation.Web.Models;
using Automation.Web.Services;

namespace Automation.Web.Controllers;

public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;
    private readonly MailReadModelService _mailReadModelService;
    private readonly MailOperatorActionService _mailOperatorActionService;

    public HomeController(
        ILogger<HomeController> logger,
        MailReadModelService mailReadModelService,
        MailOperatorActionService mailOperatorActionService)
    {
        _logger = logger;
        _mailReadModelService = mailReadModelService;
        _mailOperatorActionService = mailOperatorActionService;
    }

    public IActionResult Index()
    {
        return View(AttachFlashState(_mailReadModelService.LoadDashboard()));
    }

    public IActionResult NeedsDecision()
    {
        return View(AttachFlashState(_mailReadModelService.LoadDashboard()));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ApplyMailAction(
        string messageKey,
        string action,
        string? owner,
        string? notes,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(messageKey) || string.IsNullOrWhiteSpace(action))
        {
            TempData["MailActionKind"] = "error";
            TempData["MailActionMessage"] = "Mail action request was incomplete.";
            return RedirectToAction(nameof(NeedsDecision));
        }

        var actor = Environment.UserName;
        var result = await _mailOperatorActionService.ApplyActionAsync(
            messageKey,
            action,
            owner ?? "",
            notes ?? "",
            actor,
            cancellationToken);

        TempData["MailActionKind"] = result.Succeeded ? "success" : "error";
        TempData["MailActionMessage"] = result.Message;
        return RedirectToAction(nameof(NeedsDecision));
    }

    public IActionResult Privacy()
    {
        return View();
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }

    private MailDashboardViewModel AttachFlashState(MailDashboardViewModel viewModel)
    {
        viewModel.FlashKind = TempData["MailActionKind"] as string ?? "";
        viewModel.FlashMessage = TempData["MailActionMessage"] as string ?? "";
        return viewModel;
    }
}
