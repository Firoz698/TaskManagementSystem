using System.ComponentModel.DataAnnotations;
using TaskManagementSystem.Common;

namespace TaskManagementSystem.Models
{
    public class TaskItem
    {
        public int Id { get; set; }

        [Required(ErrorMessage = "Title is required")]
        [StringLength(200)]
        public string Title { get; set; }

        [Required(ErrorMessage = "Description is required")]
        [StringLength(1000)]
        public string Description { get; set; }

        // NEW: Add Status property
        [Required(ErrorMessage = "Status is required")]
        [StringLength(50)]
        public string Status { get; set; } = "Pending"; // Default value

        public string? ImagePath { get; set; }

        public int UserId { get; set; }

        public DateTime CreatedAt { get; set; }

        public DateTime? UpdatedAt { get; set; }

        public bool IsDeleted { get; set; } = false;
    }
}
