using System.ComponentModel.DataAnnotations;
using TaskManagementSystem.Common;

namespace TaskManagementSystem.Models
{
    public class TaskItem : Base
    {
        public int UserId { get; set; }

        [Required(ErrorMessage = "Please enter a title.")]
        [StringLength(100, ErrorMessage = "Title cannot be longer than 100 characters.")]
        public string Title { get; set; }

        [Required(ErrorMessage = "Please enter a description.")]
        [StringLength(1000, ErrorMessage = "Description cannot be longer than 1000 characters.")]
        public string Description { get; set; }

        public string? ImagePath { get; set; } // uploaded image path
    }
}
