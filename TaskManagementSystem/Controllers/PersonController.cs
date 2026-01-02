using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using TaskManagementSystem.DataContext;
using TaskManagementSystem.Models;
using TaskManagementSystem.Services;

namespace TaskManagementSystem.Controllers
{
    public class PersonController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly ActivityLogger _activityLogger;

        public PersonController(ApplicationDbContext context, ActivityLogger activityLogger)
        {
            _context = context;
            _activityLogger = activityLogger;
        }

        public async Task<IActionResult> Index(string searchString = "", int page = 1, int pageSize = 5)
        {
            var userName = HttpContext.Session.GetString("UserName") ?? "Unknown";
            var query = _context.Persons.Where(p => !p.IsDeleted && p.CreatedByUserName == userName);

            if (!string.IsNullOrWhiteSpace(searchString))
                query = query.Where(p => p.Name.Contains(searchString));

            var totalItems = await query.CountAsync();
            var persons = await query
                            .OrderByDescending(p => p.CreatedAt)
                            .Skip((page - 1) * pageSize)
                            .Take(pageSize)
                            .ToListAsync();

            ViewBag.CurrentPage = page;
            ViewBag.TotalPages = (int)Math.Ceiling(totalItems / (double)pageSize);
            ViewBag.SearchString = searchString;

            if (Request.Headers["X-Requested-With"] == "XMLHttpRequest")
                return PartialView("_PersonTablePartial", persons);

            return View(persons);
        }

        public IActionResult Create()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(Person model)
        {
            var userName = HttpContext.Session.GetString("UserName") ?? "Unknown";
            await _activityLogger.LogAsync(userName, "Add Client", $"Client Add this User : '{userName}'. ");
            if (ModelState.IsValid)
            {
                model.CreatedAt = DateTime.Now;
                model.IsActive = true;
                model.CreatedByUserName = userName;

                _context.Persons.Add(model);
                await _context.SaveChangesAsync();

                return RedirectToAction("Index");
            }

            return View(model);
        }

        public async Task<IActionResult> Details(int? id)
        {
            var userName = HttpContext.Session.GetString("UserName") ?? "Unknown";
            await _activityLogger.LogAsync(userName, "View Client Detils", $"Client view Detils this User : '{userName}'. ");
            if (id == null)
                return NotFound();

            var person = await _context.Persons
                .FirstOrDefaultAsync(m => m.Id == id && !m.IsDeleted);

            if (person == null)
                return NotFound();

            return View(person);
        }

        public async Task<IActionResult> Edit(int? id)
        {
            var userName = HttpContext.Session.GetString("UserName") ?? "Unknown";
            await _activityLogger.LogAsync(userName, "Edit Client", $"Client Edit this User : '{userName}'. ");
            if (id == null)
                return NotFound();

            var person = await _context.Persons.FindAsync(id);
            if (person == null || person.IsDeleted)
                return NotFound();

            return View(person);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, Person model)
        {
            var userName = HttpContext.Session.GetString("UserName") ?? "Unknown";
            await _activityLogger.LogAsync(userName, "Edit P Client", $"Client Id: {id} Edit this User : '{userName}'. ");

            if (id != model.Id)
                return NotFound();

            if (ModelState.IsValid)
            {
                var person = await _context.Persons.FindAsync(id);
                if (person == null || person.IsDeleted)
                    return NotFound();

                person.Name = model.Name;
                person.Address = model.Address;
                person.Phone = model.Phone;

                await _context.SaveChangesAsync();
                return RedirectToAction("Index");
            }
            return View(model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> CreateAjax([FromForm] Person model)
        {

            var userName = HttpContext.Session.GetString("UserName") ?? "Unknown";
            await _activityLogger.LogAsync(userName, "Add Client", $"Client Add with Ajax this User : '{userName}'. ");


            if (string.IsNullOrWhiteSpace(model.Name))
                return BadRequest("Name is required");

            model.CreatedAt = DateTime.Now;
            model.IsActive = true;
            model.CreatedByUserName = userName;

            _context.Persons.Add(model);
            await _context.SaveChangesAsync();

            return Json(new { id = model.Id, name = model.Name });
        }



        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Delete(int id)
        {
            var userName = HttpContext.Session.GetString("UserName") ?? "Unknown";
            await _activityLogger.LogAsync(userName, "Deleted Client", $"this Client : {id} is Deletd Add this User : '{userName}'. ");

            var person = await _context.Persons.FindAsync(id);
            if (person != null && !person.IsDeleted)
            {
                person.IsDeleted = true; 
                await _context.SaveChangesAsync();
            }

            return RedirectToAction("Index");
        }




    }
}
