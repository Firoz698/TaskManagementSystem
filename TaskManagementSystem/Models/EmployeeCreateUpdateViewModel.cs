using System.ComponentModel.DataAnnotations;

namespace TaskManagementSystem.Models
{
    public class EmployeeCreateUpdateViewModel
    {
        public int? Id { get; set; }
        [Display(Name = "Full Name")]
        [Required(ErrorMessage = "Full Name is required")]
        [StringLength(100)]
        public string FullName { get; set; } = null!;
        [Required(ErrorMessage = "Email is required")]
        [EmailAddress]
        public string Email { get; set; } = null!;
        [Display(Name = "Department")]
        [Required(ErrorMessage = "Department is required")]
        public int DepartmentId { get; set; }
        [Display(Name = "Designation")]
        [Required(ErrorMessage = "Designation is required")]
        public int DesignationId { get; set; }
        [Display(Name = "Hire Date")]
        [Required(ErrorMessage = "Hire Date is required")]
        [DataType(DataType.Date)]
        [CustomValidation(typeof(EmployeeCreateUpdateViewModel), nameof(ValidateHireDate))]
        public DateTime HireDate { get; set; } = DateTime.Today;
        [Display(Name = "Date of Birth")]
        [Required(ErrorMessage = "Date of Birth is required")]
        [DataType(DataType.Date)]
        [CustomValidation(typeof(EmployeeCreateUpdateViewModel), nameof(ValidateDateOfBirth))]
        public DateTime DateOfBirth { get; set; } = DateTime.Today.AddYears(-60);
        [Display(Name = "Employee Type")]
        [Required(ErrorMessage = "Employee Type is required")]
        public int EmployeeTypeId { get; set; }
        [Required(ErrorMessage = "Gender is required")]
        public string Gender { get; set; } = null!;
        [Required(ErrorMessage = "Salary is required")]
        [Range(0, double.MaxValue)]
        public decimal Salary { get; set; }
        // For dropdown lists
        public List<Department>? Departments { get; set; }
        public List<Designation>? Designations { get; set; }
        public List<EmployeeType>? EmployeeTypes { get; set; }
        // Validation methods:
        public static ValidationResult? ValidateHireDate(DateTime? hireDate, ValidationContext context)
        {
            if (!hireDate.HasValue)
                return new ValidationResult("Hire Date is required.");
            if (hireDate.Value.Date > DateTime.Today)
                return new ValidationResult("Hire Date cannot be in the future.");
            return ValidationResult.Success;
        }
        public static ValidationResult? ValidateDateOfBirth(DateTime? dob, ValidationContext context)
        {
            if (!dob.HasValue)
                return new ValidationResult("Date of Birth is required.");
            var minDob = DateTime.Today.AddYears(-60);  // Max age 60 years (adjust as needed)
            var maxDob = DateTime.Today.AddYears(-18);  // Min working age 18 years
            if (dob.Value.Date < minDob || dob.Value.Date > maxDob)
                return new ValidationResult($"Date of Birth must be between {minDob:yyyy-MM-dd} and {maxDob:yyyy-MM-dd}.");
            return ValidationResult.Success;
        }
    }
}
