using TaskManagementSystem.Common;
using System.Collections.Generic;

namespace TaskManagementSystem.Models.Permission
{
    public class Role : Base
    {
        public string Name { get; set; } = string.Empty;
        public string? Description { get; set; }
        public ICollection<RoleMenuPermission>? RoleMenuPermissions { get; set; }

    }
}
