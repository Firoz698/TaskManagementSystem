namespace TaskManagementSystem.Models
{
    public class FrameEditorViewModel
    {
        public List<Frame> AvailableFrames { get; set; } = new List<Frame>();
    }

    public class Frame
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public string ImagePath { get; set; } = string.Empty;
    }
}
