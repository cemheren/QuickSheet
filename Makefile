.PHONY: build run release clean help

help:
	@echo "Available targets:"
	@echo "  make build    - Build ExcelConsole.csproj"
	@echo "  make run      - Run QuickSheet in development mode"
	@echo "  make release  - Run QuickSheet in Release mode"
	@echo "  make clean    - Remove build artifacts"

build:
	dotnet build ExcelConsole.csproj

run:
	dotnet run --project ExcelConsole.csproj --

release:
	dotnet run -c Release --project ExcelConsole.csproj --

clean:
	dotnet clean ExcelConsole.csproj
	rm -rf bin obj
