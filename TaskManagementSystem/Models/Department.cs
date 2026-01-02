namespace TaskManagementSystem.Models
{
    public class Department
    {
        public int Id { get; set; }
        public string Name { get; set; } = null!;
        public bool IsActive { get; set; }
        public ICollection<Designation>? Designations { get; set; } = new List<Designation>();
        public List<Employee>? Employees { get; set; }
    }
}
