using TaskManagementSystem.Common;
using System.ComponentModel.DataAnnotations.Schema;

namespace TaskManagementSystem.Models.Permission
{
    public class RoleMenuPermission : Base
    {
        public bool CanView { get; set; } = false;
        public bool CanCreate { get; set; } = false;
        public bool CanEdit { get; set; } = false;
        public bool CanDelete { get; set; } = false;

        public int RoleId { get; set; }        // FK
        public Role? Role { get; set; }        // Navigation

        public int MenuId { get; set; }        // FK
        public Menu? Menu { get; set; }        // Navigation
    }

}
