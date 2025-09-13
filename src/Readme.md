
#  Direct Deposit System (DDS)
The system facilitates direct deposit transactions with a modern web interface and  API services.

---

## Prerequisites

Ensure you have the following installed:

### ✅ Windows & macOS
- [.NET 8 SDK](https://dotnet.microsoft.com/en-us/download/dotnet/8.0)
- [Visual Studio 2022 (Windows)](https://visualstudio.microsoft.com/)
- [Visual Studio Code (macOS)](https://code.visualstudio.com/)
- [SQL Server (Windows)](https://www.microsoft.com/en-us/sql-server/)

---

##  Installation

### 🔹 1. Clone the Repository
```sh
git clone https://github.com/muhammadyali/quelle.git
cd quelle
```

---

# Create Blazor WebAssembly Client (Frontend)
```sh
dotnet new blazorwasm -o Client --framework  net8.0 
```

# Create [ASP.NET](http://asp.net/) Core Web API (Backend)
```sh
dotnet new webapi -o Server --framework net8.0
```
# Create Shared Library (DTOs & Models)
```sh
dotnet new classlib -o Shared --framework net8.0
```
# Create a Solution File
```sh
dotnet new sln -n DDS-System
```
## Running the Project

### 🔹 1. Start the Backend API
```sh
cd Server
dotnet run
```
- The API will run at: [`https://localhost:7045`](https://localhost:7045)
- Access Swagger API documentation at: [`https://localhost:7045/swagger/index.html`](https://localhost:7045/swagger/index.html)

### 🔹 2. Start the Blazor WebAssembly Client
```sh
cd Client
dotnet run
```
- The UI will run at: [`http://localhost:5008`](http://localhost:5008)

💡 **Tip:** Use `dotnet watch run` for automatic rebuilds when making code changes:
```sh
dotnet watch run
```

---

## 📂 Project Structure

The project follows a structured architecture with clear separation between frontend, backend, and shared components.

```
/src/
├── .gitignore
├── DDS-System.sln                 # Solution File
├── Readme.md
│
├── Client/                         # Blazor WebAssembly Frontend
│   ├── App.razor
│   ├── Client.csproj
│   ├── Components/                  # Reusable UI Components
│   │   └── FileUploader.razor
│   ├── Layout/                      # UI Layouts
│   │   ├── MainLayout.razor
│   │   ├── MainLayout.razor.css
│   │   ├── NavMenu.razor
│   │   └── NavMenu.razor.css
│   ├── Models/                      # Frontend Models (DTOs)
│   │   ├── Transaction.cs
│   │   ├── FileUploadDto.cs
│   │   ├── ValidationResultDto.cs
│   │   ├── PostingHistoryDto.cs
│   │   ├── UserDto.cs
│   ├── Pages/                        # Blazor Pages
│   │   ├── LoadFunb.razor
│   │   ├── LoadFunb.razor.css
│   │   ├── ValidateTransaction.razor
│   │   ├── ValidateTransaction.razor.css
│   │   ├── Posting.razor
│   │   ├── Reporting.razor
│   │   ├── PreEditMaintenance.razor
│   │   ├── StateTreasurer.razor
│   │   ├── Login.razor
│   │   ├── Register.razor
│   ├── Services/                     # API Communication
│   │   ├── FunbFileService.cs
│   │   ├── ValidationService.cs
│   │   ├── PostingService.cs
│   │   ├── ReportService.cs
│   │   ├── TreasurerService.cs
│   │   ├── AuthenticationService.cs
│   ├── Program.cs
│   ├── Properties/
│   │   └── launchSettings.json
│   ├── wwwroot/                      # Static Assets
│   │   ├── css/
│   │   │   ├── app.css
│   │   │   └── bootstrap/
│   │   │       ├── bootstrap.min.css
│   │   │       └── bootstrap.min.css.map
│   │   ├── favicon.png
│   │   ├── icon-192.png
│   │   ├── index.html
│   │   ├── validation.jpg
│   ├── _Imports.razor
│
├── Server/                         # ASP.NET Core Web API Backend
│   ├── appsettings.Development.json
│   ├── appsettings.json
│   ├── Program.cs
│   ├── Server.csproj
│   ├── Properties/
│   │   └── launchSettings.json
│   ├── Controllers/                  # API Controllers
│   │   ├── FunbFileController.cs
│   │   ├── ValidationController.cs
│   │   ├── PostingController.cs
│   │   ├── ReportsController.cs
│   │   ├── TreasurerController.cs
│   │   ├── AuthenticationController.cs
│   ├── Services/                     # Business Logic Layer
│   │   ├── Interfaces/
│   │   │   ├── IFunbFileService.cs
│   │   │   ├── IValidationService.cs
│   │   │   ├── IPostingService.cs
│   │   │   ├── IReportService.cs
│   │   │   ├── ITreasurerService.cs
│   │   │   ├── IAuthenticationService.cs
│   │   ├── Implementations/
│   │   │   ├── FunbFileService.cs
│   │   │   ├── ValidationService.cs
│   │   │   ├── PostingService.cs
│   │   │   ├── ReportService.cs
│   │   │   ├── TreasurerService.cs
│   │   │   ├── AuthenticationService.cs
│   ├── Infrastructure/                     # Database & External Integrations
│   │   ├── AdoDbContext.cs
│   │   ├── HL7IntegrationService.cs
│   │   ├── StateTreasurerFileHandler.cs
│   ├── Models/                        # DTOs for API Responses
│   │   ├── FileUploadDto.cs
│   │   ├── ValidationResultDto.cs
│   │   ├── PostingHistoryDto.cs
│   │   ├── TreasurerFileDto.cs
│   │   ├── UserDto.cs
│
├── Shared/                           # Common Models Between Client & Server
│   ├── Shared.csproj
│   ├── Models/
│   │   ├── TransactionDto.cs
│   │   ├── FileUploadDto.cs
│   │   ├── ValidationResultDto.cs
│   │   ├── PostingHistoryDto.cs
│   │   ├── TreasurerFileDto.cs
│   │   ├── UserDto.cs

```

---

## 📜 License

This project is licensed under the **MIT License**.
