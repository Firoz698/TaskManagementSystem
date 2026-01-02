using Microsoft.AspNetCore.Mvc;
using TaskManagementSystem.DataContext;
using TaskManagementSystem.Models;
using System.Linq;
using iTextSharp.text.pdf;
using iTextSharp.text;
using Microsoft.EntityFrameworkCore;
using TaskManagementSystem.Services;

namespace TaskManagementSystem.Controllers
{
    public class ExpenseController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly IWebHostEnvironment _environment;
        private readonly ActivityLogger _logger;

        public ExpenseController(ApplicationDbContext context, IWebHostEnvironment environment, ActivityLogger logger)
        {
            _context = context;
            _environment = environment;
            _logger = logger;
        }

        [HttpGet]
        public async Task<IActionResult> Index(string? month, DateTime? from, DateTime? to, int page = 1, int pageSize = 10)
        {
            var expenses = _context.Expenses.Where(x => !x.IsDeleted).AsQueryable();

            if (from.HasValue || to.HasValue)
            {
                if (from.HasValue)
                    expenses = expenses.Where(x => x.Date.HasValue && x.Date.Value >= from.Value);
                if (to.HasValue)
                    expenses = expenses.Where(x => x.Date.HasValue && x.Date.Value <= to.Value);

                ViewBag.FromDate = from?.ToString("yyyy-MM-dd");
                ViewBag.ToDate = to?.ToString("yyyy-MM-dd");
                ViewBag.SelectedMonth = null;
            }
            else if (!string.IsNullOrEmpty(month))
            {
                if (DateTime.TryParse($"{month}-01", out var selectedMonth))
                {
                    var monthStart = new DateTime(selectedMonth.Year, selectedMonth.Month, 1);
                    var monthEnd = monthStart.AddMonths(1).AddDays(-1);
                    expenses = expenses.Where(x => x.Date.HasValue && x.Date.Value >= monthStart && x.Date.Value <= monthEnd);
                    ViewBag.SelectedMonth = month;
                }
            }
            else
            {
                var currentMonth = DateTime.Now.Month;
                var currentYear = DateTime.Now.Year;
                expenses = expenses.Where(x => x.Date.HasValue &&
                                               x.Date.Value.Month == currentMonth &&
                                               x.Date.Value.Year == currentYear);
                ViewBag.SelectedMonth = DateTime.Now.ToString("yyyy-MM");
                ViewBag.FromDate = null;
                ViewBag.ToDate = null;
            }

            var totalRecords = expenses.Count();
            var totalPages = Math.Max(1, (int)Math.Ceiling((double)totalRecords / pageSize)); // Always at least 1

            page = Math.Max(1, Math.Min(page, totalPages)); // Ensure page in range

            var expensesList = expenses
                .OrderByDescending(x => x.Date)
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .ToList();

            ViewBag.TotalAmount = expenses.Any() ? expenses.Sum(x => x.Amount) : 0;
            ViewBag.CurrentPage = page;
            ViewBag.PageSize = pageSize;
            ViewBag.TotalRecords = totalRecords;
            ViewBag.TotalPages = totalPages;

            return View(expensesList);
        }

        [HttpGet]
        public IActionResult Create()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Create(Expense model)
        {
            if (ModelState.IsValid)
            {
                var userId = HttpContext.Session.GetInt32("UserId") ?? 0;
                var userName = HttpContext.Session.GetString("UserName") ?? "Unknown";

                if (!model.Date.HasValue)
                    model.Date = DateTime.Now;

                model.CreatedByUserId = userId;
                model.CreatedByUserName = userName;

                _context.Expenses.Add(model);
                _context.SaveChanges();

                return RedirectToAction("Index");
            }

            return View(model);
        }

        [HttpGet]
        public IActionResult Edit(int id)
        {
            var expense = _context.Expenses.Find(id);
            if (expense == null) return NotFound();
            return View(expense);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Edit(Expense model)
        {
            if (ModelState.IsValid)
            {
                var expense = _context.Expenses.Find(model.Id);
                if (expense == null) return NotFound();

                expense.Description = model.Description;
                expense.Amount = model.Amount;
                expense.Date = model.Date ?? DateTime.Now;

                _context.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Delete(int id)
        {
            var expense = _context.Expenses.Find(id);
            if (expense == null) return NotFound();
            expense.IsDeleted = true;
            _context.SaveChanges();

            return RedirectToAction("Index");
        }


        [HttpGet]
        public async Task<IActionResult> ExpenseReport(string? month, DateTime? from, DateTime? to)
        {
            var userId = HttpContext.Session.GetInt32("UserId");
            var userName = HttpContext.Session.GetString("UserName") ?? "Unknown User";

            var expenses = _context.Expenses.Where(x => !x.IsDeleted).AsQueryable();

            // Filter by From/To date (priority)
            if (from.HasValue || to.HasValue)
            {
                if (from.HasValue)
                    expenses = expenses.Where(x => x.Date.HasValue && x.Date.Value >= from.Value);
                if (to.HasValue)
                    expenses = expenses.Where(x => x.Date.HasValue && x.Date.Value <= to.Value);
            }
            else if (!string.IsNullOrEmpty(month)) // Filter by month if no date range
            {
                if (DateTime.TryParse($"{month}-01", out var selectedMonth))
                {
                    var monthStart = new DateTime(selectedMonth.Year, selectedMonth.Month, 1);
                    var monthEnd = monthStart.AddMonths(1).AddDays(-1);
                    expenses = expenses.Where(x => x.Date.HasValue && x.Date.Value >= monthStart && x.Date.Value <= monthEnd);
                }
            }
            else // Default: current month
            {
                var currentMonth = DateTime.Now.Month;
                var currentYear = DateTime.Now.Year;
                expenses = expenses.Where(x => x.Date.HasValue &&
                                               x.Date.Value.Month == currentMonth &&
                                               x.Date.Value.Year == currentYear);
            }

            var expenseList = await expenses.OrderByDescending(x => x.Date).ToListAsync();
            if (!expenseList.Any())
            {
                TempData["Error"] = "No expenses found for the selected filter.";
                return RedirectToAction(nameof(Index));
            }

            var totalAmount = expenseList.Sum(x => x.Amount);

            // ── PDF Setup ──
            using var ms = new MemoryStream();
            var doc = new Document(PageSize.A4.Rotate(), 20f, 20f, 40f, 60f); // extra bottom margin for footer
            var writer = PdfWriter.GetInstance(doc, ms);

            // Footer event
            writer.PageEvent = new PdfFooter(userName);

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
            var logoPath = Path.Combine(_environment.WebRootPath, "Logo", "logo.jpg");
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

            headerTable.AddCell(new PdfPCell(new Phrase("Expense Report", headerFont))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = Element.ALIGN_CENTER,
                VerticalAlignment = Element.ALIGN_MIDDLE
            });

            doc.Add(headerTable);
            doc.Add(new Paragraph("\n"));

            // Subtitle / Filters
            var filterParts = new List<string>();
            if (!string.IsNullOrEmpty(month)) filterParts.Add($"Month: {month}");
            if (from.HasValue && to.HasValue) filterParts.Add($"Date: {from:dd-MMM-yyyy} to {to:dd-MMM-yyyy}");
            var subtitleText = $"Total Expenses: {expenseList.Count}" + (filterParts.Count > 0 ? $" (Filters – {string.Join(", ", filterParts)})" : "");
            var subtitle = new Paragraph(subtitleText, normalFont) { Alignment = Element.ALIGN_CENTER, SpacingAfter = 10f };
            doc.Add(subtitle);

            // Table
            var table = new PdfPTable(5) { WidthPercentage = 100f };
            table.SetWidths(new float[] { 5f, 25f, 40f, 15f, 15f });
            string[] headers = { "SL", "Date", "Description", "Amount", "Added By" };
            foreach (var h in headers)
                table.AddCell(new PdfPCell(new Phrase(h, boldFont)) { HorizontalAlignment = Element.ALIGN_CENTER, BackgroundColor = BaseColor.LightGray, Padding = 5f });

            int sl = 1;
            foreach (var exp in expenseList)
            {
                table.AddCell(new PdfPCell(new Phrase(sl.ToString(), normalFont)) { Padding = 5f });
                table.AddCell(new PdfPCell(new Phrase(exp.Date?.ToString("dd-MMM-yyyy") ?? "-", normalFont)) { Padding = 5f });
                table.AddCell(new PdfPCell(new Phrase(exp.Description, normalFont)) { Padding = 5f });
                table.AddCell(new PdfPCell(new Phrase(exp.Amount.ToString("0.00"), normalFont)) { Padding = 5f, HorizontalAlignment = Element.ALIGN_RIGHT });
                table.AddCell(new PdfPCell(new Phrase(exp.CreatedByUserName ?? "-", normalFont)) { Padding = 5f });
                sl++;
            }

            // Total row
            table.AddCell(new PdfPCell(new Phrase("Total", boldFont)) { Colspan = 3, HorizontalAlignment = Element.ALIGN_RIGHT, Padding = 5f });
            table.AddCell(new PdfPCell(new Phrase(totalAmount.ToString("0.00"), boldFont)) { HorizontalAlignment = Element.ALIGN_RIGHT, Padding = 5f });
            table.AddCell(new PdfPCell(new Phrase("")) { Padding = 5f }); // empty for "Added By"

            doc.Add(table);
            doc.Close();

            return File(ms.ToArray(), "application/pdf", "ExpenseReport.pdf");
        }

        public class PdfFooter : PdfPageEventHelper
        {
            private readonly string _userName;
            public PdfFooter(string userName)
            {
                _userName = userName;
            }

            public override void OnEndPage(PdfWriter writer, Document document)
            {
                var font = FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.NORMAL, BaseColor.Gray);
                var cb = writer.DirectContent;
                float yPos = document.PageSize.GetBottom(30);

                // Left
                ColumnText.ShowTextAligned(cb, Element.ALIGN_LEFT, new Phrase("Developed by Md Firoz Ali", font),
                    document.LeftMargin, yPos, 0);

                // Center
                ColumnText.ShowTextAligned(cb, Element.ALIGN_CENTER, new Phrase($"Report generated by {_userName}", font),
                    document.PageSize.Width / 2, yPos, 0);

                // Right (Page Number)
                ColumnText.ShowTextAligned(cb, Element.ALIGN_RIGHT, new Phrase($"Page {writer.PageNumber}", font),
                    document.PageSize.Width - document.RightMargin, yPos, 0);
            }
        }


    }
}
