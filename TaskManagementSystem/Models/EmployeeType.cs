namespace TaskManagementSystem.Models
{
    public class EmployeeType
    {
        public int Id { get; set; }
        public string Name { get; set; } = null!;
        public bool IsActive { get; set; }
        public List<Employee>? Employees { get; set; }
    }
}
