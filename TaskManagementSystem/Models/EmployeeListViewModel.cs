namespace TaskManagementSystem.Models
{
    public class EmployeeListViewModel
    {
        public List<Employee>? Employees { get; set; }
        public int TotalPages { get; set; }
        public string? SearchTerm { get; set; }
        public int? SelectedDepartmentId { get; set; }
        public int? SelectedEmployeeTypeId { get; set; }
        public int PageSize { get; set; }
        public int PageNumber { get; set; }
        public int Total { get; set; }
        public List<Department>? Departments { get; set; }
        public List<EmployeeType>? EmployeeTypes { get; set; }
    }
}
