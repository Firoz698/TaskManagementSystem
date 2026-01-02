using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using TaskManagementSystem.Common;

namespace TaskManagementSystem.Models
{
    public class Expense : Base
    {

        [DataType(DataType.Date)]
        public DateTime? Date { get; set; } = DateTime.Now;

        [Required]
        [StringLength(200)]
        public string Description { get; set; }

        [Required]
        [DataType(DataType.Currency)]
        public decimal Amount { get; set; }

        // ✅ Just store the user ID who created the expense
        [Required]
        public int CreatedByUserId { get; set; }

        // (Optional) Display only username in report view
        public string? CreatedByUserName { get; set; }
    }
}
