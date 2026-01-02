using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using TaskManagementSystem.Common;

namespace TaskManagementSystem.Models
{
    public class Transaction : Base
    {
        // 🔹 Person Relation
        [ForeignKey("Person")]
        [Required(ErrorMessage = "Please select a person")]
        public int PersonId { get; set; } // Foreign Key
        public Person? Person { get; set; } // Navigation property

        // 🔹 ঐচ্ছিকভাবে পুরোনো PersonName রাখা হয়েছে compatibility এর জন্য
        [StringLength(100)]
        public string? PersonName { get; set; }

        [Required]
        [Column(TypeName = "decimal(18,2)")]
        public decimal Amount { get; set; } // টাকার পরিমাণ

        [Required]
        [StringLength(20)]
        public string? Type { get; set; } // "Given" বা "Received" বা "Loan"

        public string? Description { get; set; } // বিস্তারিত

        [Required]
        [DataType(DataType.Date)]
        public DateTime Date { get; set; } = DateTime.Now;

        // 🔹 User Relation (যে user add করেছে)
        [ForeignKey("User")]
        public int UserId { get; set; }
        public User? User { get; set; }

        [StringLength(100)]
        public string? CreatedByUserName { get; set; }
    }
}
     