# online.edu.loader
online.edu.loader 


# build and run
1. You need to install SDK 6.0.x https://dotnet.microsoft.com/en-us/download/dotnet/6.0
2. Зайти в папку online.edu.loader через командную строку.
3. dotnet build, the result should be successful.
4. dotnet run --project LoadOfDisciplines [X-CN-UUID] [OrganizationId] "[Шаблон дисциплины.xlsx]" [url to api of online.edu.ru]
4. dotnet run --project LoadOfEducationalPrograms [X-CN-UUID] [OrganizationId] "[Шаблон образовательной программы.xlsx]" [url to api of online.edu.ru]