# Editing workbook XML by hand

Read this before you change sheet XML yourself. Use `xlsxpatch.py` when one of its operations fits, and `colinsert.py` to insert columns. This file is for the rest: new rows, new notes, new parts.

Contents:

1. The package
2. Cells
3. Adding rows
4. Tables
5. Notes
6. New parts
7. Calculation and stored results
8. Traps in the patch code itself
9. Symptoms and causes

## 1. The package

An `.xlsx` file is a zip. The parts that matter:

| Part | Holds |
|---|---|
| `[Content_Types].xml` | The type of every part. A part without a type makes the file unreadable. |
| `xl/workbook.xml` and `xl/_rels/workbook.xml.rels` | Sheet names and order, defined names, calculation settings. The relationship file maps each sheet to its part. |
| `xl/worksheets/sheetN.xml` | Cells of one sheet. `N` is not the tab order: find the part through the relationship id of the sheet. |
| `xl/sharedStrings.xml` | Text values, referenced by index. |
| `xl/styles.xml` | Cell formats. A cell points to one with `s="index"`. |
| `xl/tables/tableN.xml` | Excel tables: range, column names, column formulas. |
| `xl/queryTables`, `xl/connections.xml`, `customXml/item1.xml` | Power Query: what loads where, and the query code. |
| `xl/commentsN.xml` and `xl/drawings/vmlDrawingN.vml` | Notes, and the shape that displays each note. |
| `xl/calcChain.xml` | The order in which Excel calculated the formulas. |

Change only the parts that your edit needs. Write every other part back with the same bytes, the same compression and in the same order.

## 2. Cells

```xml
<c r="B5" s="12"><v>40</v></c>                              number
<c r="B6" s="12" t="inlineStr"><is><t>Text</t></is></c>     text, written in place
<c r="B7" s="12" t="s"><v>31</v></c>                        text, index 31 of the shared strings
<c r="B8" s="12"><f>B5*2</f><v>80</v></c>                   formula with its stored result
<c r="B9" s="12" t="str"><f>A9&amp;"x"</f><v>ax</v></c>     formula that returns text
```

- Write new text as an inline string. The shared strings part then stays untouched. Excel moves it to the shared strings at its next save.
- Keep the `s` of the cell you replace, or copy it from a cell that looks right. A cell without `s` has the default format.
- A formula has no leading `=`. Escape `&`, `<` and `>`.
- Rows must be in ascending order. Cells in a row must be in column order, each one in its own row, with no repeats.
- If you add a cell beyond the old last column of a row, remove the `spans` attribute of the row. Excel computes it again.
- A date is a number with a date format.

## 3. Adding rows

The safe way is to clone a row that already has the right formats and formulas.

1. Copy the XML of the template row. Change `r` on the row and on each cell.
2. Shift the row-relative references in each formula. Absolute references stay. openpyxl's `Translator` does this correctly; a regular expression does not, because of names, strings and sheet references.
3. If a template cell is part of a shared formula (`<f t="shared" si="3"/>`), write the translated formula as a plain `<f>` in the new cell. Do not add the new cell to the group and do not rewrite the master of the group.
4. Remove stored results from the cloned formulas, or compute the new ones. A copied result is wrong for the new row.
5. Extend everything that describes the range:
   - `<dimension ref>` of the sheet,
   - the `ref` of an Excel table and of its `autoFilter`,
   - a sheet-level `autoFilter`,
   - defined names in `workbook.xml` that cover the range, including the hidden `_xlnm._FilterDatabase` and print areas,
   - `sqref` of conditional formats and data validations that should cover the new rows.
6. Run the checks.

If a query loads the table, do not add rows by hand. The next refresh rewrites the table. Add the rows to the source of the query.

Tables that people keep by hand and that formulas look up by key do not grow by themselves. When a new key appears in the data, add it there too, or the lookups return their default without any error.

## 4. Tables

- The text of each header cell must be equal to the `name` of its `tableColumn`. If they differ, Excel removes the table when it opens the file.
- Column names are unique without regard to case. Column ids are unique. `count` is the number of columns. `ref` includes the header row and the totals row.
- A calculated column has its formula in two places: `calculatedColumnFormula` in the table part (the default for new rows) and the formula of each existing cell. Change both: `table_rewrite` and `rewrite` in `xlsxpatch.py`.
- Inserting a column moves everything to its right: cells, references in formulas, column widths, conditional formats, validations, notes and their shapes, the table range and columns, query table fields, defined names. `colinsert.py` does all of that. Do not attempt it with ad hoc code.

## 5. Notes

A note needs four things that agree: the `<comment ref="B5">` in the comments part, a shape in the VML drawing with the same cell (`<x:Row>` and `<x:Column>`, counted from zero), the `<legacyDrawing r:id>` element of the sheet, and the relationships and content types of both parts. A note without its shape disappears, and Excel may report damage. Shape ids must be unique inside a drawing.

Threaded comments have a second representation (`xl/threadedComments`, `xl/persons`) next to a placeholder note. Do not edit them by hand unless you change both.

## 6. New parts

A new part needs three things: an `Override` in `[Content_Types].xml` (or a `Default` for its extension), a `Relationship` from the part that owns it, and the element that uses that relationship id. `xlsx_check.py` reports a part without a type, a relationship without a target, and an `r:id` that no relationship defines.

## 7. Calculation and stored results

- `fullCalcOnLoad="1"` on `<calcPr>` in `workbook.xml` makes Excel recalculate everything at the next open.
- Remove `xl/calcChain.xml` when you add or remove formulas, together with its relationship in `xl/_rels/workbook.xml.rels` and its `Override`. A chain that names a cell with no formula is damage; `xlsx_check.py` reports it. Excel builds a new chain.
- These two settings fix what a person sees in Excel. They do not put results in the file. Until Excel saves the workbook, a new formula has no `<v>`, and every reader other than Excel sees an empty cell. Either store the results (`xlsxpatch.py cached`) or have Excel open and save the file.
- A stored text result needs `t="str"` on the cell, an error needs `t="e"`. A shared formula dependent keeps its `<f t="shared" si="N"/>` and takes the `<v>` after it.

## 8. Traps in the patch code itself

These came from real defects.

- **A cell pattern that runs into the next cell.** `<c r="K49" s="3"/>` is self-closing. A pattern such as `<c r="K49"[^>]*>.*?</c>` then swallows the following cell and you overwrite the wrong one. Match the self-closing form first and never cross a `<c` tag. `xlsxpatch.py` has a correct pattern (`CELL_RE`).
- **An opening `<f>` pattern that matches `<f .../>`.** A shared-formula cell that depends on a master has `<f t="shared" si="3"/>`: no text and no `</f>`. A pattern `<f[^>]*>(.*?)</f>` takes the "/" as part of the attributes and runs on to the `</f>` of the next cell. Exclude it: `<f(?:\s[^>]*?)?(?<!/)>`.
- **Greedy attribute matches.** Use `[^>]*?`, not `[^>]*`, before a specific attribute or before `/>`.
- **The master of a shared formula.** The first cell of a group holds the formula text for all of them. If you replace it with a plain formula, the other cells lose theirs. To change some cells of a group, un-share the whole group; the `rewrite` operation does this and refuses a range that cuts a group.
- **Replace by position, not by text.** A formula text can occur in many cells. Find the cell, then change inside it.
- **Fail when the file is not as expected.** State the old formula and compare before you replace. A patch built for last week's file must stop on this week's file.
- **Names in attributes are escaped.** A table column `A & B` is `A &amp; B`, and a line break in a header is `_x000a_`.

## 9. Symptoms and causes

| What Excel shows | Usual cause |
|---|---|
| "We found a problem with some content", then "Removed Records: ... from /xl/worksheets/sheetN.xml" | Cells or rows out of order, a repeated cell, a shared string index that does not exist, a shared formula without its master |
| "Removed Feature: Table from /xl/tables/tableN.xml" or AutoFilter removed | A header cell that differs from the column name, a wrong `ref`, repeated column names |
| The file does not open at all | XML that is not well-formed, a part without a content type, a relationship to a missing part |
| Notes are gone | Comment without its VML shape, or the reverse |
| A person sees right numbers, a query or script sees empty cells | Formulas without stored results |
| Old numbers after the change | No `fullCalcOnLoad`, and a stale stored result |
| A number shows as text, or text as an error | Wrong `t` on the cell |

Run `xlsx_check.py` first. It names the part and the cell for most of these. Do not keep a file that Excel repaired: Excel has removed what it could not read. Go back to the base.
