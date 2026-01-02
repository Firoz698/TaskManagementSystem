using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using TaskManagementSystem.Models;
using System.Linq;
using TaskManagementSystem.DataContext;
using System;
using TaskManagementSystem.Services;
using Microsoft.Extensions.Logging;
using iTextSharp.text.pdf;
using iTextSharp.text;
using Microsoft.Extensions.Hosting;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace TaskManagementSystem.Controllers
{
    public class AccountController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly ActivityLogger _logger;

        private readonly IWebHostEnvironment _environment;
        public AccountController(ApplicationDbContext context, ActivityLogger logger, IWebHostEnvironment environment)
        {
            _context = context;
            _logger = logger;
            _environment = environment;
        }

        public IActionResult Login()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Login(string username, string password)
        {
            if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
            {
                ViewBag.Error = "Username and Password are required!";
                return View();
            }

            var user = _context.Users
                .Where(u => !u.IsDeleted)
                .FirstOrDefault(u => (u.UserName ?? "").Equals(username) && (u.Password ?? "").Equals(password));

            if (user != null)
            {
                HttpContext.Session.SetString("UserName", user.UserName ?? "");
                HttpContext.Session.SetInt32("UserId", user.Id);

                await _logger.LogAsync(user.UserName, "Login", $"User '{user.UserName}' logged in.");

                return RedirectToAction("Index", "Home");

            }

            ViewBag.Error = "Invalid credentials!";
            await _logger.LogAsync(username ?? "Unknown", "FailedLogin", $"Failed login attempt for '{username}'.");

            return View();
        }

        public async Task<IActionResult> Logout()
        {
            var userName = HttpContext.Session.GetString("UserName") ?? "Unknown";
            await _logger.LogAsync(userName, "Logout", $"User '{userName}' logged out.");

            HttpContext.Session.Clear();
            return RedirectToAction("Login");
        }

        public IActionResult Dashboard()
        {
            var userName = HttpContext.Session.GetString("UserName");
            if (string.IsNullOrEmpty(userName))
            {
                return RedirectToAction("Login");
            }

            ViewBag.UserName = userName;
            return View();
        }


        public IActionResult Index(string searchText = "", int page = 1, int pageSize = 9)
        {
            if (string.IsNullOrEmpty(HttpContext.Session.GetString("UserName")))
                return RedirectToAction("Login");

            var query = _context.Users
                                .Include(u => u.Role) // ✅ Include Role
                                .Where(u => !u.IsDeleted)
                                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(searchText))
            {
                query = query.Where(u =>
                    u.UserName.Contains(searchText) ||
                    u.Email.Contains(searchText) ||
                    u.Address.Contains(searchText) ||
                    u.Contact.Contains(searchText) ||
                    u.About.Contains(searchText) ||
                    (u.Role != null && u.Role.Name.Contains(searchText)) // ✅ search by role
                );
            }

            var totalUsers = query.Count();
            var totalPages = (int)Math.Ceiling(totalUsers / (double)pageSize);

            var pagedUsers = query
                             .OrderBy(u => u.Id)
                             .Skip((page - 1) * pageSize)
                             .Take(pageSize)
                             .ToList();

            ViewBag.CurrentPage = page;
            ViewBag.TotalPages = totalPages;
            ViewBag.SearchText = searchText;

            if (Request.Headers["X-Requested-With"] == "XMLHttpRequest")
            {
                return PartialView("_UserTable", pagedUsers);
            }

            return View(pagedUsers);
        }

        public IActionResult Create()
        {
            ViewBag.Roles = new SelectList(_context.Roles.Where(r => !r.IsDeleted && r.IsActive), "Id", "Name");
            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(User user, IFormFile? Photo)
        {
            if (ModelState.IsValid)
            {
                if (Photo != null && Photo.Length > 0)
                {
                    var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/images/users");
                    if (!Directory.Exists(uploadsFolder))
                        Directory.CreateDirectory(uploadsFolder);

                    var fileName = Guid.NewGuid().ToString() + Path.GetExtension(Photo.FileName);
                    var filePath = Path.Combine(uploadsFolder, fileName);

                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await Photo.CopyToAsync(stream);
                    }

                    user.PhotoPath = "/images/users/" + fileName;
                }

                user.CreatedAt = DateTime.Now;
                user.IsActive = true;
                user.IsDeleted = false;

                _context.Users.Add(user);
                await _context.SaveChangesAsync();

                // ✅ Log creation
                await _logger.LogAsync(HttpContext.Session.GetString("UserName") ?? "System",
                                       "CreateUser", $"User '{user.UserName}' created.");

                return RedirectToAction("Index");
            }

            ViewBag.Roles = new SelectList(_context.Roles
                                           .Where(r => !r.IsDeleted && r.IsActive)
                                           .ToList(), "Id", "Name", user.RoleId);

            return View(user);
        }

        public IActionResult Edit(int id)
        {
            var user = _context.Users.FirstOrDefault(u => u.Id == id && !u.IsDeleted);
            if (user == null) return NotFound();

            // ✅ Populate Role dropdown
            ViewBag.Roles = new SelectList(_context.Roles
                                           .Where(r => !r.IsDeleted && r.IsActive)
                                           .ToList(), "Id", "Name", user.RoleId);

            return View(user);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(User user, IFormFile? Photo)
        {
            var existingUser = _context.Users.Find(user.Id);
            if (existingUser == null) return NotFound();

            if (Photo != null && Photo.Length > 0)
            {
                var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/images/users");
                if (!Directory.Exists(uploadsFolder))
                    Directory.CreateDirectory(uploadsFolder);

                var fileName = Guid.NewGuid().ToString() + Path.GetExtension(Photo.FileName);
                var filePath = Path.Combine(uploadsFolder, fileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await Photo.CopyToAsync(stream);
                }

                existingUser.PhotoPath = "/images/users/" + fileName;
            }

            existingUser.UserName = user.UserName;
            existingUser.Email = user.Email;
            if (!string.IsNullOrEmpty(user.Password))
                existingUser.Password = user.Password;

            existingUser.Address = user.Address;
            existingUser.Contact = user.Contact;
            existingUser.About = user.About;
            existingUser.RoleId = user.RoleId;
            existingUser.UpdatedAt = DateTime.Now;

            _context.Users.Update(existingUser);
            await _context.SaveChangesAsync();

            // ✅ Log update
            await _logger.LogAsync(HttpContext.Session.GetString("UserName") ?? "System",
                                   "EditUser", $"User '{existingUser.UserName}' updated.");

            return RedirectToAction("Index");
        }

        public async Task<IActionResult> Delete(int id)
        {
            var user = _context.Users.FirstOrDefault(u => u.Id == id && !u.IsDeleted);
            if (user == null) return NotFound();

            user.IsDeleted = true;
            user.UpdatedAt = DateTime.Now;
            user.UpdatedBy = HttpContext.Session.GetString("UserName") ?? "System";

            _context.Users.Update(user);
            _context.SaveChanges();

            // ✅ Log delete
            await _logger.LogAsync(HttpContext.Session.GetString("UserName") ?? "System",
                                   "DeleteUser", $"User '{user.UserName}' deleted.");

            return RedirectToAction("Index");
        }


        public IActionResult Details(int id)
        {
            var user = _context.Users
                               .Include(u => u.Role)  // <-- Include Role
                               .FirstOrDefault(u => u.Id == id && !u.IsDeleted);

            if (user == null) return NotFound();

            return View(user);
        }


        [HttpGet]
        public async Task<IActionResult> UserReport(string? search)
        {
            var query = _context.Users.Include(u => u.Role).Where(u => !u.IsDeleted);

            if (!string.IsNullOrEmpty(search))
            {
                query = query.Where(u =>
                    (u.UserName ?? "").Contains(search) ||
                    (u.Email ?? "").Contains(search) ||
                    (u.Address ?? "").Contains(search) ||
                    (u.Contact ?? "").Contains(search) ||
                    (u.About ?? "").Contains(search) ||
                    (u.Role != null && u.Role.Name.Contains(search))
                );
            }

            var users = await query.OrderBy(u => u.Id).ToListAsync();

            if (!users.Any())
            {
                TempData["Error"] = "No user data found for report.";
                return RedirectToAction(nameof(Dashboard));
            }

            await _logger.LogAsync(HttpContext.Session.GetString("UserName") ?? "System",
                                   "UserReport", $"Generated user report PDF. Total Users: {users.Count}");

            using var ms = new MemoryStream();
            var doc = new Document(PageSize.A4, 20f, 20f, 40f, 40f);

            // Font
            var fontPath = Path.Combine(_environment.WebRootPath, "fonts", "arial.ttf");
            if (!System.IO.File.Exists(fontPath))
                fontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf");

            var baseFont = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            var normalFont = new Font(baseFont, 10);
            var boldFont = new Font(baseFont, 12, Font.BOLD);
            var headerFont = new Font(baseFont, 18, Font.BOLD);

            var writer = PdfWriter.GetInstance(doc, ms);

            // Footer with page numbers and creator
            writer.PageEvent = new PdfPageEvents(baseFont);

            doc.Open();

            // Header
            var logoPath = Path.Combine(_environment.WebRootPath, "images", "fire_logo.png");
            var headerTable = new PdfPTable(2) { WidthPercentage = 100, SpacingAfter = 10f };
            headerTable.SetWidths(new float[] { 15f, 85f });

            if (System.IO.File.Exists(logoPath))
            {
                var logo = iTextSharp.text.Image.GetInstance(logoPath);
                logo.ScaleAbsolute(50f, 50f);
                headerTable.AddCell(new PdfPCell(logo) { Border = Rectangle.NO_BORDER, HorizontalAlignment = Element.ALIGN_RIGHT });
            }
            else
            {
                headerTable.AddCell(new PdfPCell(new Phrase("")) { Border = Rectangle.NO_BORDER });
            }

            headerTable.AddCell(new PdfPCell(new Phrase("Task Management System User Report", headerFont))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = Element.ALIGN_CENTER,
                VerticalAlignment = Element.ALIGN_MIDDLE
            });

            doc.Add(headerTable);

            // Subtitle
            var subTitle = new Paragraph($"User Report (Total: {users.Count})", normalFont)
            {
                Alignment = Element.ALIGN_CENTER,
                SpacingAfter = 15f
            };
            doc.Add(subTitle);

            // Table
            var table = new PdfPTable(7) { WidthPercentage = 100f }; // ✅ Added Role column
            table.SetWidths(new float[] { 5f, 18f, 22f, 15f, 15f, 15f, 15f });

            string[] headers = { "SL", "User Name", "Email", "Role", "Contact", "Address", "About" };

            foreach (var h in headers)
            {
                table.AddCell(new PdfPCell(new Phrase(h, boldFont))
                {
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    BackgroundColor = BaseColor.LightGray,
                    Padding = 5f
                });
            }

            int sl = 1;
            foreach (var user in users)
            {
                table.AddCell(new PdfPCell(new Phrase(sl.ToString(), normalFont)) { Padding = 5f });
                table.AddCell(new PdfPCell(new Phrase(user.UserName ?? "", normalFont)) { Padding = 5f });
                table.AddCell(new PdfPCell(new Phrase(user.Email ?? "", normalFont)) { Padding = 5f });
                table.AddCell(new PdfPCell(new Phrase(user.Role?.Name ?? "N/A", normalFont)) { Padding = 5f }); // ✅ Role
                table.AddCell(new PdfPCell(new Phrase(user.Contact ?? "", normalFont)) { Padding = 5f });
                table.AddCell(new PdfPCell(new Phrase(user.Address ?? "", normalFont)) { Padding = 5f });
                table.AddCell(new PdfPCell(new Phrase(user.About ?? "", normalFont)) { Padding = 5f });
                sl++;
            }

            doc.Add(table);
            doc.Close();

            return File(ms.ToArray(), "application/pdf", "UserReport.pdf");
        }


        // ✅ Page Event Handler for Footer
        public class PdfPageEvents : PdfPageEventHelper
        {
            private BaseFont _baseFont;
            public PdfPageEvents(BaseFont baseFont)
            {
                _baseFont = baseFont;
            }

            public override void OnEndPage(PdfWriter writer, Document document)
            {
                PdfPTable footerTable = new PdfPTable(2) { TotalWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin };
                footerTable.SetWidths(new float[] { 50f, 50f });

                footerTable.AddCell(new PdfPCell(new Phrase("Report created by Md Firoz", new Font(_baseFont, 9)))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_LEFT
                });

                footerTable.AddCell(new PdfPCell(new Phrase($"Page {writer.PageNumber}", new Font(_baseFont, 9)))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_RIGHT
                });

                footerTable.WriteSelectedRows(0, -1, document.LeftMargin, document.BottomMargin - 5, writer.DirectContent);
            }
        }


    }
}
