using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Diagnostics;
using TaskManagementSystem.DataContext;
using TaskManagementSystem.Models;
using TaskManagementSystem.Services;

namespace TaskManagementSystem.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly ApplicationDbContext _context;
        private readonly ActivityLogger _actilogger;

        public HomeController(ILogger<HomeController> logger, ApplicationDbContext context, ActivityLogger actilogger)
        {
            _logger = logger;
            _context = context;
            _actilogger = actilogger;
        }

        public async Task<IActionResult> Index()
        {
            var name = HttpContext.Session.GetString("UserName");
            ViewBag.UserName = name;
            var recentLogs = await _context.ActivityLogs
                .OrderByDescending(x => x.CreatedAt)
                .Take(5)
                .ToListAsync();
            ViewBag.RecentLogs = recentLogs;
            return View();
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
    }
}
