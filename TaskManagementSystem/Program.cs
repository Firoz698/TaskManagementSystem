using TaskManagementSystem.DataContext;
using TaskManagementSystem.Services;
using Microsoft.EntityFrameworkCore;
using Rotativa.AspNetCore;
using TaskManagementSystem.DataContext;
using TaskManagementSystem.Services;
using TaskManagementSystem.Helpers;

namespace EmployeePortal
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            // 1️⃣ Add MVC controllers with views
            builder.Services.AddControllersWithViews();

            // 2️⃣ Configure SQL Server connection
            builder.Services.AddDbContext<ApplicationDbContext>(options =>
                options.UseSqlServer(builder.Configuration.GetConnectionString("ECommerceDBConnection")));

            // 3️⃣ Register custom services
            builder.Services.AddScoped<EmployeeService>();
            builder.Services.AddHttpContextAccessor();
            builder.Services.AddScoped<ActivityLogger>();
            builder.Services.AddScoped<PermissionHelper>();



            // 4️⃣ Configure Session
            builder.Services.AddSession(options =>
            {
                options.IdleTimeout = TimeSpan.FromMinutes(30);
                options.Cookie.HttpOnly = true;
                options.Cookie.IsEssential = true;
            });

            var app = builder.Build();

            // 5️⃣ Configure HTTP pipeline
            if (app.Environment.IsDevelopment())
            {
                // Show detailed errors during development
                app.UseDeveloperExceptionPage();
            }
            else
            {
                // Use standard error handler in production
                app.UseExceptionHandler("/Home/Error");
                app.UseHsts();
            }

            app.UseHttpsRedirection();
            app.UseStaticFiles();

            app.UseRouting();

            // ✅ Enable session
            app.UseSession();

            // ⚠️ Authorization (optional if login implemented manually)
            app.UseAuthorization();

            // 6️⃣ Rotativa configuration (PDF generator)
            var rotativaRoot = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Rotativa");
           // RotativaConfiguration.Setup(rotativaRoot);

            // 7️⃣ Default route
            app.MapControllerRoute(
                name: "default",
                pattern: "{controller=Account}/{action=Login}/{id?}");

            app.Run();
        }
    }
}
