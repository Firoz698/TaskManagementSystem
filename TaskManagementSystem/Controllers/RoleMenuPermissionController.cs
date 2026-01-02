using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using TaskManagementSystem.Models.Permission;
using TaskManagementSystem.DataContext;
using TaskManagementSystem.Models;
using TaskManagementSystem.Services;

namespace TaskManagementSystem.Controllers
{
    public class RoleMenuPermissionController : Controller
    {

        private readonly ActivityLogger _activityLogger;
        private readonly ApplicationDbContext _context;

        public RoleMenuPermissionController(ApplicationDbContext context, ActivityLogger activityLogger)
        {
            _context = context;
            _activityLogger = activityLogger;
        }
        public async Task<IActionResult> Index(int? userId)
        {
            var currentUser = HttpContext.Session.GetString("UserName") ?? "Unknown";
            ViewBag.Users = await _context.Users
                .Where(u => !u.IsDeleted && u.IsActive)
                .Select(u => new { u.Id, u.UserName, u.RoleId })
                .ToListAsync();

            if (userId == null)
            {
                ViewBag.SelectedUserId = null;
                ViewBag.SelectedUserRole = null;

                await _activityLogger.LogAsync(
                    currentUser,
                    "View RoleMenuPermission",
                    $"User '{currentUser}' viewed RoleMenuPermission page without selecting a user."
                );

                return View(new List<RoleMenuPermission>());
            }

            var user = await _context.Users
                .Include(u => u.Role)
                .FirstOrDefaultAsync(u => u.Id == userId && !u.IsDeleted && u.IsActive);

            if (user == null)
            {
                ViewBag.SelectedUserId = userId;
                ViewBag.SelectedUserRole = "Not Found";
                await _activityLogger.LogAsync(
                    currentUser,
                    "View RoleMenuPermission",
                    $"User '{currentUser}' tried to view permissions for non-existing user ID: {userId}."
                );

                return View(new List<RoleMenuPermission>());
            }

            var permissions = await _context.RoleMenuPermissions
                .Include(rmp => rmp.Menu)
                .Include(rmp => rmp.Role)
                .Where(rmp => rmp.RoleId == user.RoleId && !rmp.IsDeleted && rmp.IsActive)
                .OrderBy(rmp => rmp.Menu.Title)
                .ToListAsync();

            ViewBag.SelectedUserId = userId;
            ViewBag.SelectedUserRole = user.Role?.Name ?? "No Role";

            await _activityLogger.LogAsync(
                currentUser,
                "View RoleMenuPermission",
                $"User '{currentUser}' viewed permissions for user '{user.UserName}' (Role: '{user.Role?.Name ?? "No Role"}')."
            );

            return View(permissions);
        }

        [HttpPost]
        public async Task<IActionResult> UpdatePermissions(int userId, List<RoleMenuPermission> Permissions)
        {
            var currentUser = HttpContext.Session.GetString("UserName") ?? "Unknown";
            var user = await _context.Users
                .Include(u => u.Role)
                .FirstOrDefaultAsync(u => u.Id == userId && !u.IsDeleted && u.IsActive);

            if (user == null)
            {
                TempData["Error"] = "User not found!";
                await _activityLogger.LogAsync(
                    currentUser,
                    "Update RoleMenuPermission",
                    $"User '{currentUser}' attempted to update permissions for non-existing user ID: {userId}."
                );
                return RedirectToAction("Index");
            }

            foreach (var perm in Permissions)
            {
                var dbPerm = await _context.RoleMenuPermissions
                    .FirstOrDefaultAsync(p => p.Id == perm.Id);

                if (dbPerm != null)
                {
                    dbPerm.CanView = perm.CanView;
                    dbPerm.CanCreate = perm.CanCreate;
                    dbPerm.CanEdit = perm.CanEdit;
                    dbPerm.CanDelete = perm.CanDelete;
                    dbPerm.UpdatedAt = DateTime.Now;
                }
            }
            await _context.SaveChangesAsync();
            await _activityLogger.LogAsync(
                currentUser,
                "Update RoleMenuPermission",
                $"User '{currentUser}' updated permissions for user '{user.UserName}' (Role: '{user.Role?.Name ?? "No Role"}')."
            );

            TempData["Success"] = "Permissions updated successfully!";
            return RedirectToAction("Index", new { userId = userId });
        }

        public IActionResult Create()
        {
            LoadDropdowns();
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(RoleMenuPermission permission)
        {
            if (ModelState.IsValid)
            {
                var exists = await _context.RoleMenuPermissions
                    .AnyAsync(rmp => rmp.RoleId == permission.RoleId && rmp.MenuId == permission.MenuId);

                if (exists)
                {
                    ModelState.AddModelError("", "Permission already exists for this role and menu.");
                    LoadDropdowns(permission.RoleId, permission.MenuId);
                    return View(permission);
                }

                permission.IsActive = true;
                permission.IsDeleted = false;
                permission.CreatedAt = DateTime.Now;
                _context.Add(permission);
                await _context.SaveChangesAsync();
                var currentUser = HttpContext.Session.GetString("UserName") ?? "Unknown";
                var role = await _context.Roles.FindAsync(permission.RoleId);
                var menu = await _context.Menus.FindAsync(permission.MenuId);

                string actionType = "Create RoleMenuPermission";
                string logMessage = $"User '{currentUser}' created a new permission for Role '{role?.Name}' and Menu '{menu?.Title}'.";

                await _activityLogger.LogAsync(currentUser, actionType, logMessage);

                return RedirectToAction(nameof(Index));
            }

            LoadDropdowns(permission.RoleId, permission.MenuId);
            return View(permission);
        }

        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null) return NotFound();

            var permission = await _context.RoleMenuPermissions.FindAsync(id);
            if (permission == null) return NotFound();

            LoadDropdowns(permission.RoleId, permission.MenuId);
            return View(permission);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, RoleMenuPermission permission)
        {
            if (id != permission.Id)
                return NotFound();

            if (ModelState.IsValid)
            {
                permission.UpdatedAt = DateTime.Now;
                _context.Update(permission);
                await _context.SaveChangesAsync();
                var currentUser = HttpContext.Session.GetString("UserName") ?? "Unknown";
                var role = await _context.Roles.FindAsync(permission.RoleId);
                var menu = await _context.Menus.FindAsync(permission.MenuId);

                string actionType = "Edit RoleMenuPermission";
                string logMessage = $"User '{currentUser}' edited permission for Role '{role?.Name}' and Menu '{menu?.Title}'.";
                await _activityLogger.LogAsync(currentUser, actionType, logMessage);

                return RedirectToAction(nameof(Index));
            }

            LoadDropdowns(permission.RoleId, permission.MenuId);
            return View(permission);
        }

        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null) return NotFound();

            var permission = await _context.RoleMenuPermissions
                .Include(rmp => rmp.Role)
                .Include(rmp => rmp.Menu)
                .FirstOrDefaultAsync(rmp => rmp.Id == id);

            if (permission == null) return NotFound();

            return View(permission);
        }

        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var permission = await _context.RoleMenuPermissions.FindAsync(id);
            if (permission != null)
            {
                _context.RoleMenuPermissions.Remove(permission);
                await _context.SaveChangesAsync();
                var currentUser = HttpContext.Session.GetString("UserName") ?? "Unknown";

                var role = await _context.Roles.FindAsync(permission.RoleId);
                var menu = await _context.Menus.FindAsync(permission.MenuId);

                string actionType = "Delete RoleMenuPermission";
                string logMessage = $"User '{currentUser}' deleted permission for Role '{role?.Name}' and Menu '{menu?.Title}'.";

                await _activityLogger.LogAsync(currentUser, actionType, logMessage);
            }

            return RedirectToAction(nameof(Index));
        }

        private void LoadDropdowns(int? selectedRoleId = null, int? selectedMenuId = null)
        {
            ViewBag.Roles = new SelectList(_context.Roles.Where(r => r.IsActive), "Id", "Name", selectedRoleId);
            ViewBag.Menus = new SelectList(_context.Menus.Where(m => m.IsActive), "Id", "Title", selectedMenuId);
        }
    }
}
