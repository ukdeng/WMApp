using Autofac;
using Autofac.Extensions.DependencyInjection;
using WMApp.Interfaces;
using WMApp.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();
builder.Services.AddControllersWithViews().AddRazorRuntimeCompilation();

// Use Autofac to inject dependencies
// Call UseServiceProviderFactory on the Host sub property
builder.Host.UseServiceProviderFactory(new AutofacServiceProviderFactory());

// Call ConfigureContainer on the Host sub property 
builder.Host.ConfigureContainer<ContainerBuilder>(builder =>
{
    builder.RegisterType<AccessReader>().As<IAccessReader>().SingleInstance();
    builder.RegisterType<ExcelReader>().As<IExcelReader>().SingleInstance();
    builder.RegisterType<FormFiller>().As<IFormFiller>().SingleInstance();
});


IronPdf.License.LicenseKey = "IRONSUITE.ZMOGXLUXQDGIUVGPKF.YTNHY.COM.2638-9D0E483FFF-FXKKUFIMXJ2KU7-HOOKS4OUSK4O-VDIJEKCTG36A-HYKE437PLH7A-DUXV4XK2LNEF-2CIDHZAYLFAP-4MH7V5-TFA5R5QVGKSNEA-DEPLOYMENT.TRIAL-JX65XP.TRIAL.EXPIRES.12.AUG.2024";
var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();
