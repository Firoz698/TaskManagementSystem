using Microsoft.AspNetCore.Mvc;
using TaskManagementSystem.DataContext;
using TaskManagementSystem.Models;
using TaskManagementSystem.Services;

namespace TaskManagementSystem.Controllers
{
    public class GameController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly ActivityLogger _logger;

        public GameController(ApplicationDbContext context, ActivityLogger logger)
        {
            _context = context;
            _logger = logger;
        }

        // 🎮 Game Home Page
        public IActionResult Index()
        {
            ViewData["Title"] = "Game Zone";
            return View();
        }

        // 🐍 Snake Game
        public IActionResult Snake()
        {
            ViewData["Title"] = "Snake Game";
            return View();
        }

        // ❌ Tic Tac Toe ⭕
        public IActionResult TicTacToe()
        {
            ViewData["Title"] = "Tic-Tac-Toe Game";
            return View();
        }

        // ✈️ Plane Shooter Game
        public IActionResult PlaneShooter()
        {
            ViewData["Title"] = "Plane Shooter";
            return View();
        }

        // 👧 Cute Doll Runner
        public IActionResult CuteDollRunner()
        {
            ViewData["Title"] = "Cute Doll Runner";
            return View();
        }

        // -------------------------------------------------------
        // 💾 Save Game Session — called via AJAX from the browser
        // -------------------------------------------------------
        [HttpPost]
        public async Task<IActionResult> SaveSession([FromBody] SaveSessionRequest request)
        {
            var userId = HttpContext.Session.GetInt32("UserId");
            var userName = HttpContext.Session.GetString("UserName") ?? "Guest";

            var session = new GameSession
            {
                UserId = userId,
                GameName = request.GameName,
                Score = request.Score,
                DurationSeconds = request.DurationSeconds,
                Result = request.Result,
                StartedAt = request.StartedAt,
                EndedAt = DateTime.Now,
                ExtraData = request.ExtraData,
                CreatedAt = DateTime.Now,
                IsActive = true,
                IsDeleted = false
            };

            _context.GameSessions.Add(session);
            await _context.SaveChangesAsync();

            await _logger.LogAsync(userName, "GameSession", $"'{userName}' played {request.GameName} — Score: {request.Score}, Result: {request.Result}");

            return Json(new { success = true, sessionId = session.Id });
        }

        // 📊 Leaderboard for a specific game
        [HttpGet]
        public IActionResult Leaderboard(string gameName = "Snake")
        {
            var top = _context.GameSessions
                .Where(g => !g.IsDeleted && g.GameName == gameName)
                .OrderByDescending(g => g.Score)
                .Take(10)
                .Select(g => new
                {
                    g.Score,
                    g.DurationSeconds,
                    g.Result,
                    g.StartedAt,
                    UserName = g.User != null ? g.User.UserName : "Guest"
                })
                .ToList();

            return Json(top);
        }
    }

    // DTO for the AJAX body
    public class SaveSessionRequest
    {
        public string GameName { get; set; } = "";
        public int Score { get; set; }
        public int DurationSeconds { get; set; }
        public string Result { get; set; } = "GameOver";
        public DateTime StartedAt { get; set; }
        public string? ExtraData { get; set; }
    }
}