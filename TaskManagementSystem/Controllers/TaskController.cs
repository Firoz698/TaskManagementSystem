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

        public IActionResult Index(string search)
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

            var tasks = tasksQuery
                .OrderByDescending(t => t.CreatedAt)
                .ToList();

            ViewData["Search"] = search; 

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

                _context.Tasks.Add(task);
                _context.SaveChanges();

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
        public async Task<IActionResult> TaskReport(string? search, DateTime? from, DateTime? to)
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

            var tasks = await query.OrderByDescending(t => t.CreatedAt).ToListAsync();

            if (!tasks.Any())
            {
                TempData["Error"] = "No tasks found for the selected filter.";
                return RedirectToAction(nameof(Index));
            }

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

            headerTable.AddCell(new PdfPCell(new Phrase("Task Management System Task Report", headerFont))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = Element.ALIGN_CENTER,
                VerticalAlignment = Element.ALIGN_MIDDLE
            });

            doc.Add(headerTable);
            doc.Add(new Paragraph("\n"));

            // Subtitle
            string filterText = string.Empty;
            var titleParts = new List<string>();
            if (!string.IsNullOrEmpty(search)) titleParts.Add($"Search: {search}");
            if (from.HasValue && to.HasValue) titleParts.Add($"Date: {from:dd-MMM-yyyy} to {to:dd-MMM-yyyy}");
            if (titleParts.Count > 0) filterText = " (Filters – " + string.Join(", ", titleParts) + ")";

            var subTitle = new Paragraph($"Tasks Report (Total: {tasks.Count}){filterText}", normalFont)
            {
                Alignment = Element.ALIGN_CENTER,
                SpacingAfter = 15f
            };
            doc.Add(subTitle);

            // Table
            var table = new PdfPTable(5) { WidthPercentage = 100f };
            table.SetWidths(new float[] { 5f, 20f, 40f, 15f, 20f });
            string[] headers = { "SL", "Title", "Description", "Image", "Created At" };
            foreach (var h in headers)
                table.AddCell(new PdfPCell(new Phrase(h, boldFont))
                {
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    BackgroundColor = BaseColor.LightGray,
                    Padding = 5f
                });

            int sl = 1;
            foreach (var task in tasks)
            {
                table.AddCell(new PdfPCell(new Phrase(sl.ToString(), normalFont)) { Padding = 5f });
                table.AddCell(new PdfPCell(new Phrase(task.Title ?? "", normalFont)) { Padding = 5f });
                table.AddCell(new PdfPCell(new Phrase(task.Description ?? "", normalFont)) { Padding = 5f });

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
                            table.AddCell(new PdfPCell(img) { Padding = 5f, HorizontalAlignment = Element.ALIGN_CENTER });
                        }
                        else
                        {
                            table.AddCell(new PdfPCell(new Phrase("No Image", normalFont)) { Padding = 5f, HorizontalAlignment = Element.ALIGN_CENTER });
                        }
                    }
                    catch
                    {
                        table.AddCell(new PdfPCell(new Phrase("Error", normalFont)) { Padding = 5f, HorizontalAlignment = Element.ALIGN_CENTER });
                    }
                }
                else
                {
                    table.AddCell(new PdfPCell(new Phrase("No Image", normalFont)) { Padding = 5f, HorizontalAlignment = Element.ALIGN_CENTER });
                }

                table.AddCell(new PdfPCell(new Phrase(task.CreatedAt.ToString("dd-MMM-yyyy hh:mm tt"), normalFont)) { Padding = 5f });
                sl++;
            }

            doc.Add(table);
            doc.Close();

            return File(ms.ToArray(), "application/pdf", "TasksReport.pdf");
        }
    }
}
