using iTextSharp.text.pdf;
using iTextSharp.text;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using TaskManagementSystem.DataContext;
using TaskManagementSystem.Models;

namespace TaskManagementSystem.Controllers
{
    public class TaskController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly IWebHostEnvironment _environment;
        public TaskController(ApplicationDbContext context, IWebHostEnvironment environment)
        {
            _context = context;
            _environment = environment;
        }

        public IActionResult Index(string search, string status)
        {
            var userId = HttpContext.Session.GetInt32("UserId");
            if (userId == null)
                return RedirectToAction("Login", "Account");

            var tasksQuery = _context.Tasks
                .Where(t => t.UserId == userId && !t.IsDeleted);

            if (!string.IsNullOrEmpty(search))
            {
                tasksQuery = tasksQuery.Where(t =>
                    t.Title.Contains(search) ||
                    t.Description.Contains(search) ||
                    t.CreatedAt.ToString().Contains(search));
            }

            if (!string.IsNullOrEmpty(status))
            {
                tasksQuery = tasksQuery.Where(t => t.Status == status);
            }

            var tasks = tasksQuery
                .OrderByDescending(t => t.CreatedAt)
                .ToList();

            // Calculate statistics
            var totalTasks = tasks.Count;
            var pendingCount = tasks.Count(t => t.Status == "Pending");
            var inProgressCount = tasks.Count(t => t.Status == "In Progress");
            var completedCount = tasks.Count(t => t.Status == "Completed");

            var pendingPercent = totalTasks > 0 ? (pendingCount * 100.0 / totalTasks) : 0;
            var inProgressPercent = totalTasks > 0 ? (inProgressCount * 100.0 / totalTasks) : 0;
            var completedPercent = totalTasks > 0 ? (completedCount * 100.0 / totalTasks) : 0;

            ViewData["Search"] = search;
            ViewData["Status"] = status;
            ViewData["TotalTasks"] = totalTasks;
            ViewData["PendingCount"] = pendingCount;
            ViewData["InProgressCount"] = inProgressCount;
            ViewData["CompletedCount"] = completedCount;
            ViewData["PendingPercent"] = Math.Round(pendingPercent, 1);
            ViewData["InProgressPercent"] = Math.Round(inProgressPercent, 1);
            ViewData["CompletedPercent"] = Math.Round(completedPercent, 1);

            return View(tasks);
        }

        public IActionResult Create()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Create(TaskItem task, IFormFile? ImageFile)
        {
            var userId = HttpContext.Session.GetInt32("UserId");
            if (userId == null)
                return RedirectToAction("Login", "Account");

            if (ModelState.IsValid)
            {
                if (ImageFile != null && ImageFile.Length > 0)
                {
                    string uploadFolder = Path.Combine(_environment.WebRootPath, "uploads");
                    if (!Directory.Exists(uploadFolder))
                        Directory.CreateDirectory(uploadFolder);

                    string fileName = Guid.NewGuid().ToString() + Path.GetExtension(ImageFile.FileName);
                    string filePath = Path.Combine(uploadFolder, fileName);

                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        ImageFile.CopyTo(stream);
                    }

                    task.ImagePath = "/uploads/" + fileName;
                }

                task.UserId = userId.Value;
                task.CreatedAt = DateTime.Now;
                task.Status = string.IsNullOrEmpty(task.Status) ? "Pending" : task.Status;

                _context.Tasks.Add(task);
                _context.SaveChanges();

                TempData["Success"] = "Task created successfully!";
                return RedirectToAction("Index");
            }

            return View(task);
        }

        public IActionResult Edit(int id)
        {
            var userId = HttpContext.Session.GetInt32("UserId");
            if (userId == null)
                return RedirectToAction("Login", "Account");

            var task = _context.Tasks.FirstOrDefault(t => t.Id == id && t.UserId == userId && !t.IsDeleted);
            if (task == null) return NotFound();

            return View(task);
        }

        [HttpPost]
        public IActionResult Edit(TaskItem task, IFormFile? ImageFile)
        {
            var userId = HttpContext.Session.GetInt32("UserId");
            if (userId == null)
                return RedirectToAction("Login", "Account");

            var existing = _context.Tasks.FirstOrDefault(t => t.Id == task.Id && t.UserId == userId && !t.IsDeleted);
            if (existing == null) return NotFound();

            if (ModelState.IsValid)
            {
                existing.Title = task.Title;
                existing.Description = task.Description;
                existing.Status = task.Status;
                existing.UpdatedAt = DateTime.Now;

                if (ImageFile != null && ImageFile.Length > 0)
                {
                    string uploadFolder = Path.Combine(_environment.WebRootPath, "uploads");
                    if (!Directory.Exists(uploadFolder))
                        Directory.CreateDirectory(uploadFolder);

                    string fileName = Guid.NewGuid().ToString() + Path.GetExtension(ImageFile.FileName);
                    string filePath = Path.Combine(uploadFolder, fileName);

                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        ImageFile.CopyTo(stream);
                    }

                    existing.ImagePath = "/uploads/" + fileName;
                }

                _context.Tasks.Update(existing);
                _context.SaveChanges();

                TempData["Success"] = "Task updated successfully!";
                return RedirectToAction("Index");
            }

            return View(task);
        }

        public IActionResult Delete(int id)
        {
            var userId = HttpContext.Session.GetInt32("UserId");
            if (userId == null)
                return RedirectToAction("Login", "Account");

            var task = _context.Tasks.FirstOrDefault(t => t.Id == id && t.UserId == userId && !t.IsDeleted);
            if (task == null) return NotFound();

            task.IsDeleted = true;
            task.UpdatedAt = DateTime.Now;

            _context.Tasks.Update(task);
            _context.SaveChanges();

            TempData["Success"] = "Task deleted successfully!";
            return RedirectToAction("Index");
        }

        public IActionResult Details(int id)
        {
            var userId = HttpContext.Session.GetInt32("UserId");
            if (userId == null)
                return RedirectToAction("Login", "Account");

            var task = _context.Tasks.FirstOrDefault(t => t.Id == id && t.UserId == userId && !t.IsDeleted);
            if (task == null) return NotFound();

            return View(task);
        }

        [HttpGet]
        public async Task<IActionResult> TaskReport(string? search, DateTime? from, DateTime? to, string? status)
        {
            var userId = HttpContext.Session.GetInt32("UserId");
            if (userId == null)
                return RedirectToAction("Login", "Account");

            var query = _context.Tasks.AsQueryable();
            query = query.Where(t => t.UserId == userId && !t.IsDeleted);

            if (!string.IsNullOrEmpty(search))
                query = query.Where(t =>
                    (t.Title ?? "").Contains(search) ||
                    (t.Description ?? "").Contains(search) ||
                    t.CreatedAt.ToString().Contains(search));

            if (from.HasValue && to.HasValue)
            {
                var toDate = to.Value.AddDays(1);
                query = query.Where(t => t.CreatedAt >= from && t.CreatedAt < toDate);
            }

            if (!string.IsNullOrEmpty(status))
            {
                query = query.Where(t => t.Status == status);
            }

            var tasks = await query.OrderByDescending(t => t.CreatedAt).ToListAsync();

            if (!tasks.Any())
            {
                TempData["Error"] = "No tasks found for the selected filter.";
                return RedirectToAction(nameof(Index));
            }

            // Calculate statistics
            var totalTasks = tasks.Count;
            var pendingCount = tasks.Count(t => t.Status == "Pending");
            var inProgressCount = tasks.Count(t => t.Status == "In Progress");
            var completedCount = tasks.Count(t => t.Status == "Completed");

            var pendingPercent = totalTasks > 0 ? Math.Round(pendingCount * 100.0 / totalTasks, 1) : 0;
            var inProgressPercent = totalTasks > 0 ? Math.Round(inProgressCount * 100.0 / totalTasks, 1) : 0;
            var completedPercent = totalTasks > 0 ? Math.Round(completedCount * 100.0 / totalTasks, 1) : 0;

            // ── PDF Setup ──
            using var ms = new MemoryStream();
            var doc = new Document(new Rectangle(842f, 595f), 20f, 20f, 40f, 40f); // A4 Landscape
            PdfWriter.GetInstance(doc, ms);
            doc.Open();

            // Font
            var fontPath = Path.Combine(_environment.WebRootPath, "fonts", "arial.ttf");
            if (!System.IO.File.Exists(fontPath))
                fontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf");

            var baseFont = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            var normalFont = new Font(baseFont, 10);
            var boldFont = new Font(baseFont, 12, Font.BOLD);
            var headerFont = new Font(baseFont, 20, Font.BOLD);
            var smallFont = new Font(baseFont, 9);

            // Header
            var logoPath = Path.Combine(_environment.WebRootPath, "images", "fire_logo.png");
            var headerTable = new PdfPTable(2) { WidthPercentage = 100, SpacingAfter = 5f };
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

            headerTable.AddCell(new PdfPCell(new Phrase("Task Management System - Task Report", headerFont))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = Element.ALIGN_CENTER,
                VerticalAlignment = Element.ALIGN_MIDDLE
            });

            doc.Add(headerTable);
            doc.Add(new Paragraph("\n"));

            // Statistics Section
            var statsTable = new PdfPTable(4) { WidthPercentage = 100, SpacingAfter = 15f };
            statsTable.SetWidths(new float[] { 25f, 25f, 25f, 25f });

            // Total Tasks
            statsTable.AddCell(new PdfPCell(new Phrase($"Total Tasks: {totalTasks}", boldFont))
            {
                BackgroundColor = new BaseColor(99, 102, 241),
                HorizontalAlignment = Element.ALIGN_CENTER,
                Padding = 10f,
                BorderColor = BaseColor.White,
                BorderWidth = 2f
            });

            // Pending
            statsTable.AddCell(new PdfPCell(new Phrase($"Pending: {pendingCount} ({pendingPercent}%)", boldFont))
            {
                BackgroundColor = new BaseColor(239, 68, 68),
                HorizontalAlignment = Element.ALIGN_CENTER,
                Padding = 10f,
                BorderColor = BaseColor.White,
                BorderWidth = 2f
            });

            // In Progress
            statsTable.AddCell(new PdfPCell(new Phrase($"In Progress: {inProgressCount} ({inProgressPercent}%)", boldFont))
            {
                BackgroundColor = new BaseColor(245, 158, 11),
                HorizontalAlignment = Element.ALIGN_CENTER,
                Padding = 10f,
                BorderColor = BaseColor.White,
                BorderWidth = 2f
            });

            // Completed
            statsTable.AddCell(new PdfPCell(new Phrase($"Completed: {completedCount} ({completedPercent}%)", boldFont))
            {
                BackgroundColor = new BaseColor(16, 185, 129),
                HorizontalAlignment = Element.ALIGN_CENTER,
                Padding = 10f,
                BorderColor = BaseColor.White,
                BorderWidth = 2f
            });

            doc.Add(statsTable);

            // Filter Information
            string filterText = string.Empty;
            var titleParts = new List<string>();
            if (!string.IsNullOrEmpty(search)) titleParts.Add($"Search: {search}");
            if (from.HasValue && to.HasValue) titleParts.Add($"Date: {from:dd-MMM-yyyy} to {to:dd-MMM-yyyy}");
            if (!string.IsNullOrEmpty(status)) titleParts.Add($"Status: {status}");
            if (titleParts.Count > 0) filterText = "Filters Applied: " + string.Join(", ", titleParts);

            if (!string.IsNullOrEmpty(filterText))
            {
                var filterPara = new Paragraph(filterText, smallFont)
                {
                    Alignment = Element.ALIGN_CENTER,
                    SpacingAfter = 10f
                };
                doc.Add(filterPara);
            }

            doc.Add(new Paragraph("\n"));

            // Table
            var table = new PdfPTable(6) { WidthPercentage = 100f };
            table.SetWidths(new float[] { 5f, 20f, 30f, 12f, 12f, 15f });
            string[] headers = { "SL", "Title", "Description", "Status", "Image", "Created At" };
            
            foreach (var h in headers)
                table.AddCell(new PdfPCell(new Phrase(h, boldFont))
                {
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    BackgroundColor = new BaseColor(31, 41, 55),
                    Padding = 8f,
                    BorderColor = BaseColor.White,
                    BorderWidth = 1f
                });

            int sl = 1;
            foreach (var task in tasks)
            {
                // SL
                table.AddCell(new PdfPCell(new Phrase(sl.ToString(), normalFont)) 
                { 
                    Padding = 8f, 
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    VerticalAlignment = Element.ALIGN_MIDDLE 
                });

                // Title
                table.AddCell(new PdfPCell(new Phrase(task.Title ?? "", normalFont)) 
                { 
                    Padding = 8f,
                    VerticalAlignment = Element.ALIGN_MIDDLE 
                });

                // Description
                table.AddCell(new PdfPCell(new Phrase(task.Description ?? "", normalFont)) 
                { 
                    Padding = 8f,
                    VerticalAlignment = Element.ALIGN_MIDDLE 
                });

                // Status with color
                BaseColor statusColor = task.Status switch
                {
                    "Completed" => new BaseColor(16, 185, 129),
                    "In Progress" => new BaseColor(245, 158, 11),
                    _ => new BaseColor(239, 68, 68)
                };

                table.AddCell(new PdfPCell(new Phrase(task.Status ?? "Pending", boldFont))
                {
                    Padding = 8f,
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    BackgroundColor = statusColor
                });

                // Image
                if (!string.IsNullOrEmpty(task.ImagePath))
                {
                    try
                    {
                        var imgPath = Path.Combine(_environment.WebRootPath, task.ImagePath.TrimStart('/'));
                        if (System.IO.File.Exists(imgPath))
                        {
                            var img = iTextSharp.text.Image.GetInstance(imgPath);
                            img.ScaleAbsolute(50f, 50f);
                            table.AddCell(new PdfPCell(img) 
                            { 
                                Padding = 5f, 
                                HorizontalAlignment = Element.ALIGN_CENTER,
                                VerticalAlignment = Element.ALIGN_MIDDLE 
                            });
                        }
                        else
                        {
                            table.AddCell(new PdfPCell(new Phrase("No Image", smallFont)) 
                            { 
                                Padding = 8f, 
                                HorizontalAlignment = Element.ALIGN_CENTER,
                                VerticalAlignment = Element.ALIGN_MIDDLE 
                            });
                        }
                    }
                    catch
                    {
                        table.AddCell(new PdfPCell(new Phrase("Error", smallFont)) 
                        { 
                            Padding = 8f, 
                            HorizontalAlignment = Element.ALIGN_CENTER,
                            VerticalAlignment = Element.ALIGN_MIDDLE 
                        });
                    }
                }
                else
                {
                    table.AddCell(new PdfPCell(new Phrase("No Image", smallFont)) 
                    { 
                        Padding = 8f, 
                        HorizontalAlignment = Element.ALIGN_CENTER,
                        VerticalAlignment = Element.ALIGN_MIDDLE 
                    });
                }

                // Created At
                table.AddCell(new PdfPCell(new Phrase(task.CreatedAt.ToString("dd-MMM-yyyy\nhh:mm tt"), normalFont)) 
                { 
                    Padding = 8f,
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    VerticalAlignment = Element.ALIGN_MIDDLE 
                });
                
                sl++;
            }

            doc.Add(table);

            // Footer
            doc.Add(new Paragraph("\n"));
            var footer = new Paragraph($"Generated on: {DateTime.Now:dd-MMM-yyyy hh:mm tt}", smallFont)
            {
                Alignment = Element.ALIGN_RIGHT
            };
            doc.Add(footer);

            doc.Close();

            return File(ms.ToArray(), "application/pdf", $"TasksReport_{DateTime.Now:yyyyMMdd_HHmmss}.pdf");
        }
    }
}