using System.ComponentModel.DataAnnotations;
using TaskManagementSystem.Common;

namespace TaskManagementSystem.Models
{
    public class Person : Base
    {
        [Required]
        [StringLength(100)]
        public string? Name { get; set; }

        [StringLength(200)]
        public string? Address { get; set; }

        [StringLength(15)]
        public string? Phone { get; set; }

        // user relation (যে user add করেছে)
        [StringLength(100)]
        public string? CreatedByUserName { get; set; }

        // navigation property
        public ICollection<Transaction>? Transactions { get; set; }
    }
}
