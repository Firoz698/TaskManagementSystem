using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using TaskManagementSystem.DataContext;
using TaskManagementSystem.Models;
using TaskManagementSystem.Services;

namespace TaskManagementSystem.Controllers
{
    public class GameController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly ActivityLogger _logger;
        private readonly ILogger<GameController> _log;

        public GameController(ApplicationDbContext context, ActivityLogger logger, ILogger<GameController> log)
        {
            _context = context;
            _logger = logger;
            _log = log;
        }

        // ─────────────────────────────────────────────────────────
        // GET /Game/Index  —  Game Zone home with live stats
        // ─────────────────────────────────────────────────────────

        public IActionResult Index()
        {
            ViewData["Title"] = "Game Zone";

            try
            {
                var sessions = _context.GameSessions
                    .Where(g => !g.IsDeleted)
                    .ToList(); // single DB call, then filter in memory

                // ── Total plays & unique players ──
                ViewBag.TotalPlays = sessions.Count;
                ViewBag.UniquePlayers = sessions
                    .Where(g => g.UserId != null)
                    .Select(g => g.UserId)
                    .Distinct()
                    .Count();

                // ── Per-game top score + play count ──
                var gameNames = new[] { "Snake", "TicTacToe", "PlaneShooter", "CuteDollRunner" };

                var gameStats = gameNames.ToDictionary(
                    name => name,
                    name =>
                    {
                        var gameSessions = sessions.Where(g => g.GameName == name).ToList();
                        return new
                        {
                            Plays = gameSessions.Count,
                            TopScore = gameSessions.Any() ? gameSessions.Max(g => g.Score) : 0,
                            TopUser = gameSessions
                                .OrderByDescending(g => g.Score)
                                .Select(g => g.User?.UserName ?? "Guest")
                                .FirstOrDefault() ?? "-"
                        };
                    });

                ViewBag.GameStats = gameStats;

                // ── Recent activity (last 5 sessions) ──
                ViewBag.RecentSessions = _context.GameSessions
                    .Where(g => !g.IsDeleted)
                    .OrderByDescending(g => g.CreatedAt)
                    .Take(5)
                    .Select(g => new
                    {
                        g.GameName,
                        g.Score,
                        g.Result,
                        g.CreatedAt,
                        UserName = g.User != null ? g.User.UserName : "Guest"
                    })
                    .ToList();
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "GameController.Index: failed to load stats");

                // Safe defaults so the view never crashes
                ViewBag.TotalPlays = 0;
                ViewBag.UniquePlayers = 0;
                ViewBag.GameStats = new Dictionary<string, object>();
                ViewBag.RecentSessions = new List<object>();
            }

            return View();
        }

        // ─────────────────────────────────────────────────────────
        // Game Pages
        // ─────────────────────────────────────────────────────────

        public IActionResult Snake() => SafeView("Snake Game");
        public IActionResult TicTacToe() => SafeView("Tic-Tac-Toe Game");
        public IActionResult PlaneShooter() => SafeView("Plane Shooter");
        public IActionResult CuteDollRunner() => SafeView("Cute Doll Runner");

        private IActionResult SafeView(string title)
        {
            try
            {
                ViewData["Title"] = title;
                return View();
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Failed to render view: {Title}", title);
                TempData["ErrorMessage"] = "Page could not be loaded. Please try again.";
                return RedirectToAction(nameof(Index));
            }
        }

        // ─────────────────────────────────────────────────────────
        // POST /Game/SaveSession
        // ─────────────────────────────────────────────────────────

        [HttpPost]
        public async Task<IActionResult> SaveSession([FromBody] SaveSessionRequest? request)
        {
            try
            {
                if (request == null)
                {
                    _log.LogWarning("SaveSession: null request body");
                    return Json(new { success = false, error = "Request body is missing." });
                }

                if (string.IsNullOrWhiteSpace(request.GameName))
                    return Json(new { success = false, error = "GameName is required." });

                if (request.Score < 0)
                    return Json(new { success = false, error = "Score cannot be negative." });

                if (request.DurationSeconds < 0)
                    return Json(new { success = false, error = "DurationSeconds cannot be negative." });

                int? userId = null;
                string userName = "Guest";
                try
                {
                    userId = HttpContext.Session.GetInt32("UserId");
                    userName = HttpContext.Session.GetString("UserName") ?? "Guest";
                }
                catch (Exception sessionEx)
                {
                    _log.LogWarning(sessionEx, "SaveSession: could not read session — defaulting to Guest");
                }

                var gameName = Truncate(request.GameName.Trim(), 100);
                var result = Truncate(string.IsNullOrWhiteSpace(request.Result) ? "GameOver" : request.Result.Trim(), 50);
                var extraData = request.ExtraData;
                var startedAt = request.StartedAt == default ? DateTime.Now : request.StartedAt;
                if (startedAt.Year < 1 || startedAt.Year > 9999) startedAt = DateTime.Now;

                if (userId.HasValue)
                {
                    bool userExists = await _context.Users.AnyAsync(u => u.Id == userId.Value);
                    if (!userExists) { _log.LogWarning("SaveSession: UserId {UserId} not found", userId); userId = null; }
                }

                var now = DateTime.Now;
                var session = new GameSession
                {
                    UserId = userId,
                    GameName = gameName,
                    Score = request.Score,
                    DurationSeconds = request.DurationSeconds,
                    Result = result,
                    StartedAt = startedAt,
                    EndedAt = now,
                    ExtraData = extraData,
                    CreatedAt = now,
                    UpdatedAt = null,
                    IsActive = true,
                    IsDeleted = false
                };

                _context.GameSessions.Add(session);
                await _context.SaveChangesAsync();

                try
                {
                    await _logger.LogAsync(userName, "GameSession",
                        $"'{userName}' played {session.GameName} — Score: {session.Score}, Result: {session.Result}");
                }
                catch (Exception logEx)
                {
                    _log.LogWarning(logEx, "SaveSession: activity log failed for session {Id}", session.Id);
                }

                return Json(new { success = true, sessionId = session.Id });
            }
            catch (DbUpdateException dbEx)
            {
                var sqlEx = dbEx.InnerException as SqlException;
                var sqlMsg = sqlEx != null ? $"SQL {sqlEx.Number}: {sqlEx.Message}" : dbEx.InnerException?.Message ?? dbEx.Message;
                _log.LogError(dbEx, "SaveSession: DbUpdateException — {SqlMsg}", sqlMsg);

                var clientMsg = sqlEx?.Number switch
                {
                    2627 or 2601 => "Duplicate entry detected.",
                    547 => "Data integrity error.",
                    8152 => "A text value is too long.",
                    515 => "A required field is missing.",
                    _ => "Database error — session could not be saved."
                };
                return Json(new { success = false, error = clientMsg, sqlError = sqlMsg });
            }
            catch (OperationCanceledException)
            {
                _log.LogWarning("SaveSession: request cancelled");
                return Json(new { success = false, error = "Request was cancelled." });
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "SaveSession: unexpected error");
                return Json(new { success = false, error = "An unexpected error occurred." });
            }
        }

        // ─────────────────────────────────────────────────────────
        // GET /Game/Leaderboard?gameName=Snake
        // ─────────────────────────────────────────────────────────

        [HttpGet]
        public IActionResult Leaderboard(string gameName = "Snake")
        {
            try
            {
                if (string.IsNullOrWhiteSpace(gameName)) return Json(new List<object>());

                var top = _context.GameSessions
                    .Where(g => !g.IsDeleted && g.GameName == gameName.Trim())
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
            catch (Exception ex)
            {
                _log.LogError(ex, "Leaderboard: error for '{GameName}'", gameName);
                return Json(new List<object>());
            }
        }

        // ─────────────────────────────────────────────────────────
        private static string? Truncate(string? value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return value;
            return value.Length <= maxLength ? value : value[..maxLength];
        }
    }

    // ── DTO ──────────────────────────────────────────────────────
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