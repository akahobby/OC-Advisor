# OC Advisor

OC Advisor is a lightweight Windows utility that provides structured overclocking recommendations for CPU, GPU, and RAM configurations. It analyzes assumptions and presents organized tuning guidance without automatically applying changes.

## Features

- CPU, GPU, and RAM tuning suggestions  
- Structured BIOS guidance  
- Stability-focused recommendation logic  
- Clean, minimal desktop interface  
- No automatic system modifications  

## Requirements

- Windows 10 / 11 (x64)  
- .NET 8 SDK (for building from source)  

## Build From Source (PowerShell Only)

### 1. Install .NET 8 SDK

Download from:  
https://dotnet.microsoft.com/download  

Verify installation:

```powershell
dotnet --version
```

### 2. Restore & Build

Open PowerShell in the folder containing `OcAdvisor.csproj` and run:

```powershell
dotnet restore
dotnet build -c Release
```

Build output will be located in:

```
bin\Release\net8.0-windows\
```

## Create Portable Builds

### Option A — Self-Contained Folder

Bundles the .NET runtime but keeps supporting files separate.

```powershell
dotnet publish -c Release -r win-x64 --self-contained true
```

Output:

```
bin\Release\net8.0-windows\win-x64\publish\
```

Distribute the entire `publish` folder.

### Option B — Single Portable EXE (Recommended)

Creates a single standalone executable:

```powershell
dotnet publish -c Release -r win-x64 --self-contained true `
  /p:PublishSingleFile=true `
  /p:IncludeNativeLibrariesForSelfExtract=true `
  /p:IncludeAllContentForSelfExtract=true `
  /p:DebugType=None `
  /p:DebugSymbols=false
```

Output:

```
bin\Release\net8.0-windows\win-x64\publish\
```

Distribute the generated `.exe`.

This method:
- Bundles the .NET runtime  
- Bundles WPF native dependencies  
- Produces a clean single-file release  
- Does not require .NET installed on the target system  

## Notes

- Single-file builds are larger due to the bundled runtime.  
- No system changes are automatically applied by this application.  
- Overclocking adjustments must be applied manually via BIOS or tuning software.  

## Disclaimer

Overclocking carries risk. You are responsible for changes made to your system.

## License

MIT
