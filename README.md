# XLSX to Markdown

XlsxToMarkdown is een C#-tool waarmee je eenvoudig tabellen uit Excel-bestanden (.xlsx) kunt converteren naar Markdown-formaat. Dit is handig om snel en efficiënt gegevens uit Excel te delen in bijvoorbeeld documentatie, GitHub README's, of Wiki-pagina's. Het is gebouwd met .NET 10.0.

## Features

- Converteer één of meerdere werkbladen uit een .xlsx-bestand naar overzichtelijke Markdown-tabellen
- Ondersteuning voor verschillende opmaakopties (bijv. uitlijning, kopteksten)
- Batchverwerking van meerdere bestanden mogelijk
- Eenvoudig te gebruiken via een command-line interface
- Snel en lichtgewicht

## Installatie

1. Clone deze repository:
   ```bash
   git clone https://github.com/Zaibatsu89/XlsxToMarkdown.git
   ```
2. Open het project in Visual Studio of een andere C# IDE.
3. Herstel de NuGet packages (indien nodig).
4. Bouw het project.

## Gebruik

1. Zorg dat je een .xlsx-bestand hebt met de gewenste tabellen.
2. Start de applicatie via de command line:

   ```bash
   XlsxToMarkdown.exe -i "<pad/naar/excelbestand.xlsx>" -o "<pad/naar/output.md>"
   ```

**Voorbeeld:**
```bash
XlsxToMarkdown.exe -i "C:\data\voorbeeld.xlsx" -o "C:\data\voorbeeld.md"
```

### Opties

- `-i`, `--input` : Pad naar het .xlsx-bestand (verplicht)
- `-o`, `--output` : Pad voor het gegenereerde .md-bestand (verplicht)
- `--sheet` : Specifiek werkblad om te converteren (optioneel)
- `--all-sheets` : Converteer alle werkbladen (optioneel)

## Voorbeeld output

```markdown
| Naam    | Leeftijd | Beroep      |
|---------|----------|-------------|
| Jan     | 32       | Programmeur |
| Lisa    | 28       | Designer    |
```

## Bijdragen

Bijdragen zijn welkom! Open een issue voor bug reports of feature requests, of maak een pull request voor verbeteringen.

## Licentie

Dit project is gelicenseerd onder de MIT-licentie.

## Contact

Voor vragen of feedback, maak gerust een issue aan op GitHub of neem contact op via [Zaibatsu89](https://github.com/Zaibatsu89).
