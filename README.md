# Co-author Affiliation Formatter

After having manually edited a list of ca. 75 co-authors (for Zuidema et al. 2022; https://doi.org/10.1038/s41561-022-00911-8), 
I created an R script that organizes and formats a list of coauthors and their institutional affiliations for use in scientific publications. 

It produces two output formats:
- A Word document with properly superscripted affiliations
- An HTML-formatted string that can be pasted into a Word document via an HTML viewer

---

## 📌 What It Does:

Given an Excel file with coauthor information (names, ORCID, order, and affiliations), the script:
1. Parses each author's affiliations
2. Assigns numeric superscripts to unique affiliations
3. Formats the author line and affiliation list
4. Outputs the result as:
   - A Word document (`output_author_list.docx`)
   - An HTML string copied to the clipboard

## 📌 What It Does not do:
- Organize the funding information of all co-authors
- Organize the list of people and entities that need to be acknowledged
You can organize this type of information in the same Excel (per co-author), but as it is less structured than the affiliations, 
I recommend doing the formatting of these two important parts manually.

---

## 📂 Input File

The input Excel file must:
- Be named: `Example author info.xlsx` (or similar — just update the file name in the script if needed)
- Contain a worksheet named `"Author Info"`
- Have the following columns starting on the second row:

| Column Name         | Description                                                              |
|---------------------|--------------------------------------------------------------------------|
| `name_in_publication` | Full name as it should appear in the manuscript                        |
| `surname`             | Surname (used for internal reference and checks)                       |
| `first_name`          | First name (used for internal reference and checks)                    |
| `order`               | Numeric value defining the author order (e.g., 1 = first author)       |
| `abrev`               | Initials or shorthand for internal use                                 |
| `aff_txt1`            | First affiliation (e.g., "University of São Paulo, Brazil")            |
| `aff_txt2`            | Second affiliation (if any)                                            |
| `ORCID`               | Author's ORCID iD (often solicited by journals)                        |
| `email`               | Email (not used in output, just stored for reference)                  |

*Note:* Additional affiliation columns can be added (e.g., `aff_txt3`) but must follow the same naming pattern.

---

## 🛠 How to Use

1. **Install R and RStudio** (if you haven’t already):
   - Download R: https://cran.r-project.org/
   - Download RStudio: https://posit.co/download/rstudio-desktop/

2. **Open the script in RStudio** and run it step by step or press "Source".

3. **Install Required Packages (first time only)**:
```r
install.packages(c("tidyverse", "readxl", "officer"))
```

4. **Ensure your Excel file is in your working directory (often the same folder as the R script)**.

5. **Run the script**:
   - It will produce a Word file (`output_author_list.docx`) in the same folder
   - It will also copy an HTML-formatted string to your clipboard

6. **(Optional) If using the HTML output**:
   - Go to: https://www.w3schools.com/html/tryit.asp?filename=tryhtml_editor
   - Paste the copied text into the editor and run it
   - Copy the rendered output into Word

---

## 🌍 Encoding Notes

If you're working with names or affiliations that contain accented or non-ASCII characters (e.g., "Gonçalves", "Göteborg", "Łódź"), the script includes a `Sys.setlocale()` call to handle encoding:

```r
Sys.setlocale(category = "LC_ALL", locale = "English_United States.1252")
```

This setting works for many Windows systems, but may need adjustment on macOS or Linux. If characters still render incorrectly, comment out, update the locale setting for your system, or change 
the "enconding" in R when loading the file.

---

## ✅ Output Files

- `output_author_list.docx`: Word file with formatted authors and affiliations
- HTML string: copied to clipboard (can be pasted in Word via online editor)

---

## 🧪 Testing

You can test the script using the included file:
- `Example_author_info_FAKE.xlsx`: A fake dataset with 10 fictional coauthors and affiliations

---

## 💡 Suggestions

- Add more affiliation columns if needed (`aff_txt3`, etc.)
- Customize font or size in the Word document by editing the `fp_text` settings
- Consider converting this into an RMarkdown document if you want inline previews or web output

---

## 📜 License

This script is provided under the MIT License. Feel free to use, adapt, or share it with attribution.

Peter Groenendijk
