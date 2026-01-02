using TaskManagementSystem.Common;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace TaskManagementSystem.Models.Permission
{
    public class Menu : Base
    {
        public string Title { get; set; } = string.Empty;
        public string ControllerName { get; set; } = string.Empty;
        public string ActionName { get; set; } = string.Empty;
        public int? ParentId { get; set; }
    }
}
