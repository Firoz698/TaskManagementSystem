    using Microsoft.EntityFrameworkCore;
using TaskManagementSystem.DataContext;
using TaskManagementSystem.Models;

namespace TaskManagementSystem.Services
{
    public class EmployeeService
    {
        private readonly ApplicationDbContext _context;
        public EmployeeService(ApplicationDbContext context)
        {
            _context = context;
        }
        public async Task<(List<Employee> Employees, int TotalCount)> GetEmployees(
            string? searchTerm,
            int? departmentId,
            int? employeeTypeId,
            int pageNumber,
            int pageSize)
        {
            var query = _context.Employees
                .Include(e => e.Department)
                .Include(e => e.Designation)
                .Include(e => e.EmployeeType)
                .AsQueryable();
            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                query = query.Where(e => e.FullName.Contains(searchTerm));
            }
            if (departmentId.HasValue && departmentId.Value > 0)
            {
                query = query.Where(e => e.DepartmentId == departmentId.Value);
            }
            if (employeeTypeId.HasValue && employeeTypeId.Value > 0)
            {
                query = query.Where(e => e.EmployeeTypeId == employeeTypeId.Value);
            }
            var totalCount = await query.CountAsync();
            var employees = await query
                .Skip((pageNumber - 1) * pageSize)
                .Take(pageSize)
                .AsNoTracking()
                .ToListAsync();
            return (employees, totalCount);
        }
        public async Task<Employee?> GetEmployeeByIdAsync(int id)
        {
            return await _context.Employees
                .Include(e => e.Department)
                .Include(e => e.Designation)
                .Include(e => e.EmployeeType)
                .AsNoTracking()
                .FirstOrDefaultAsync(e => e.Id == id);
        }
        public async Task CreateEmployeeAsync(Employee employee)
        {
            _context.Employees.Add(employee);
            await _context.SaveChangesAsync();
        }
        public async Task UpdateEmployeeAsync(Employee employee)
        {
            _context.Entry(employee).State = EntityState.Modified;
            await _context.SaveChangesAsync();
        }
        public async Task DeleteEmployeeAsync(int id)
        {
            var employee = await GetEmployeeByIdAsync(id);
            if (employee != null)
            {
                _context.Employees.Remove(employee);
                await _context.SaveChangesAsync();
            }
        }
        // Master tables retrieval
        public async Task<List<Department>> GetDepartmentsAsync()
        {
            return await _context.Departments.AsNoTracking().ToListAsync();
        }
        public async Task<List<EmployeeType>> GetEmployeeTypesAsync()
        {
            return await _context.EmployeeTypes.AsNoTracking().ToListAsync();
        }
        public async Task<List<Designation>> GetDesignationsByDepartmentAsync(int departmentId)
        {
            return await _context.Designations
                .Where(d => d.DepartmentId == departmentId)
                .AsNoTracking()
                .ToListAsync();
        }
    }
}
