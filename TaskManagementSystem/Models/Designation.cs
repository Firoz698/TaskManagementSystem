namespace TaskManagementSystem.Models
{
    public class Designation
    {
        public int Id { get; set; }
        public string Name { get; set; } = null!;
        public int DepartmentId { get; set; }
        public Department Department { get; set; } = null!;
        public bool IsActive { get; set; }
    }
}
