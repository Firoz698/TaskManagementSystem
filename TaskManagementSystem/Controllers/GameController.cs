using Microsoft.AspNetCore.Mvc;

namespace TaskManagementSystem.Controllers
{
    public class GameController : Controller
    {
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

        // ❌ Tic Tac Toe Game ⭕
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
    }
}
