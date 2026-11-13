var builder = DistributedApplication.CreateBuilder(args);
builder.AddProject<Projects.TemplateGenerator>("webui");
builder.Build().Run();
