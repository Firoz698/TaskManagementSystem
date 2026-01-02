namespace TaskManagementSystem.Models
{
    public class ActivityLog
    {
        public int Id { get; set; }
        public string? UserName { get; set; }
        public string? ActionType { get; set; } // e.g. Create, Update, Delete
        public string? Description { get; set; }
        public string? IpAddress { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.Now;
    }
}
