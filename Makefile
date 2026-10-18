.PHONY: build run desktop release clean export-md export-html export-json

build:
	dotnet build ExcelConsole.csproj

run:
	dotnet run --project ExcelConsole.csproj

desktop:
	dotnet run --project ExcelConsole.csproj -- --desktop

release:
	dotnet run -c Release --project ExcelConsole.csproj -- --desktop

clean:
	dotnet clean ExcelConsole.csproj
	rm -rf bin obj

export-md:
	@test -n "$(CSV)" || (echo "Usage: make export-md CSV=data.csv OUT=data.md" && exit 1)
	dotnet run --project ExcelConsole.csproj -- $(CSV) --export-md $(OUT)

export-html:
	@test -n "$(CSV)" || (echo "Usage: make export-html CSV=data.csv OUT=data.html" && exit 1)
	dotnet run --project ExcelConsole.csproj -- $(CSV) --export-html $(OUT)

export-json:
	@test -n "$(CSV)" || (echo "Usage: make export-json CSV=data.csv OUT=data.json" && exit 1)
	dotnet run --project ExcelConsole.csproj -- $(CSV) --export-json $(OUT)
