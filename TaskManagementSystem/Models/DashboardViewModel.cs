namespace TaskManagementSystem.Models
{
    public class DashboardViewModel
    {
        public int TotalUsers { get; set; }
        public int ActiveSessions { get; set; }
        public int ReportsToday { get; set; }
        public int PendingTasks { get; set; }

        public List<ActivityViewModel> RecentActivities { get; set; } = new List<ActivityViewModel>();
    }
}
