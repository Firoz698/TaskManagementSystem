using Microsoft.AspNetCore.Mvc;
using TaskManagementSystem.Models;

namespace TaskManagementSystem.Controllers
{
    public class FrameEditorController : Controller
    {
        private readonly IWebHostEnvironment _environment;

        public FrameEditorController(IWebHostEnvironment environment)
        {
            _environment = environment;
        }

        public IActionResult Index()
        {
            var model = new FrameEditorViewModel
            {
                AvailableFrames = GetAvailableFrames()
            };
            return View(model);
        }

        [HttpPost]
        public async Task<IActionResult> UploadPhoto(IFormFile photo)
        {
            if (photo == null || photo.Length == 0)
            {
                return BadRequest("No file uploaded");
            }

            var uploadsFolder = Path.Combine(_environment.WebRootPath, "uploads");
            if (!Directory.Exists(uploadsFolder))
            {
                Directory.CreateDirectory(uploadsFolder);
            }

            var uniqueFileName = Guid.NewGuid().ToString() + Path.GetExtension(photo.FileName);
            var filePath = Path.Combine(uploadsFolder, uniqueFileName);

            using (var fileStream = new FileStream(filePath, FileMode.Create))
            {
                await photo.CopyToAsync(fileStream);
            }

            return Json(new { success = true, filePath = "/uploads/" + uniqueFileName });
        }

        private List<Frame> GetAvailableFrames()
        {
            return new List<Frame>
            {
                new Frame { Id = 1, Name = "Classic Frame", ImagePath = "/images/frames/frame1.svg" },
                new Frame { Id = 2, Name = "Vintage Frame", ImagePath = "/images/frames/frame2.svg" },
                new Frame { Id = 3, Name = "Modern Frame", ImagePath = "/images/frames/frame3.svg" },
                new Frame { Id = 4, Name = "Ornate Frame", ImagePath = "/images/frames/frame4.svg" },
                new Frame { Id = 5, Name = "Simple Frame", ImagePath = "/images/frames/frame5.svg" },
                new Frame { Id = 6, Name = "Heart Frame", ImagePath = "/images/frames/frame6.svg" }
            };
        }
    }
}
