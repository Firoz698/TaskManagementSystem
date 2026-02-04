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
                // Classic Frames (PNG)
                new Frame { Id = 1, Name = "Classic Frame", ImagePath = "/images/frames/frame1.png" },
                new Frame { Id = 2, Name = "Vintage Frame", ImagePath = "/images/frames/frame2.png" },
                new Frame { Id = 3, Name = "Modern Frame", ImagePath = "/images/frames/frame3.png" },
                new Frame { Id = 4, Name = "Ornate Frame", ImagePath = "/images/frames/frame4.png" },
                new Frame { Id = 5, Name = "Simple Frame", ImagePath = "/images/frames/frame5.png" },
                new Frame { Id = 6, Name = "Heart Frame", ImagePath = "/images/frames/frame6.png" },
                
                // Valentine's Day Frames (PNG)
                new Frame { Id = 7, Name = "Valentine Classic", ImagePath = "/images/frames/valentine-frame.png" },
                new Frame { Id = 8, Name = "Valentine Modern", ImagePath = "/images/frames/valentine-frame-2.png" },
                new Frame { Id = 9, Name = "Valentine Premium", ImagePath = "/images/frames/valentine-frame-3.png" },
                
                // Vote/Democracy/Political Frames (PNG)
                new Frame { Id = 10, Name = "Vote Classic", ImagePath = "/images/frames/vote-frame-1.png" },
                new Frame { Id = 11, Name = "Vote Gradient", ImagePath = "/images/frames/vote-frame-2.png" },
                new Frame { Id = 12, Name = "Vote Premium", ImagePath = "/images/frames/vote-frame-3.png" },
                new Frame { Id = 13, Name = "Vote Modern", ImagePath = "/images/frames/vote-frame-4.png" },
                new Frame { Id = 14, Name = "Political Frame", ImagePath = "/images/frames/political-frame-1.png" },
                
                // শহীদ দিবস Frames (PNG)
                new Frame { Id = 15, Name = "শহীদ দিবস ক্লাসিক", ImagePath = "/images/frames/shaheed-frame-1.png" },
                new Frame { Id = 16, Name = "শহীদ দিবস মডার্ন", ImagePath = "/images/frames/shaheed-frame-2.png" },
                new Frame { Id = 17, Name = "শহীদ দিবস প্রিমিয়াম", ImagePath = "/images/frames/shaheed-frame-3.png" },
                
                // স্বাধীনতা দিবস Frames (PNG)
                new Frame { Id = 18, Name = "স্বাধীনতা পতাকা", ImagePath = "/images/frames/independence-frame-1.png" },
                new Frame { Id = 19, Name = "স্বাধীনতা স্মৃতিসৌধ", ImagePath = "/images/frames/independence-frame-2.png" },
                
                // বিজয় দিবস Frame (PNG)
                new Frame { Id = 20, Name = "বিজয় সূর্যোদয়", ImagePath = "/images/frames/victory-frame-1.png" },
                
                // ঈদ মুবারক Frame (PNG)
                new Frame { Id = 21, Name = "ঈদ মুবারক", ImagePath = "/images/frames/eid-frame-1.png" },
                
                // পহেলা বৈশাখ Frame (PNG)
                new Frame { Id = 22, Name = "নববর্ষ উৎসব", ImagePath = "/images/frames/boishakh-frame-1.png" },
                
                // শিক্ষা Frame (PNG)
                new Frame { Id = 23, Name = "শিক্ষাই আলো", ImagePath = "/images/frames/education-frame-1.png" }
            };
        }
    }
}