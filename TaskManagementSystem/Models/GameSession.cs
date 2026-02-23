using TaskManagementSystem.Common;

namespace TaskManagementSystem.Models
{
    /// <summary>
    /// Tracks each game session a user plays
    /// </summary>
    public class GameSession : Base
    {
        // Which user played
        public int? UserId { get; set; }
        public User? User { get; set; }

        // Which game (e.g. "Snake", "TicTacToe", "PlaneShooter", "CuteDollRunner")
        public string? GameName { get; set; }

        // Score achieved in this session
        public int Score { get; set; }

        // How long they played (in seconds)
        public int DurationSeconds { get; set; }

        // Session outcome: "Win", "GameOver", "Quit"
        public string? Result { get; set; }

        // When the game started
        public DateTime StartedAt { get; set; }

        // When the game ended
        public DateTime? EndedAt { get; set; }

        // Extra info (e.g. level reached, kills, etc.)
        public string? ExtraData { get; set; }
    }
}