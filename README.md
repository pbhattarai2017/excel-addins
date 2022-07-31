# excel-addins

This repository contains custom excel addins.

Functions in **NepaliMonthName.xlam** <br>
- `NEPALIMONTH(miti)`<br>
  Example: `=NEPALIMONTH("2078/01/01")` gives *Baishakh*<br>
  Example: `=NEPALIMONTH("2078/11/09")` gives *Falgun*
- `NEPALIMONTHNUMBER(miti)`<br>
  Example: `=NEPALIMONTHNUMBER("2078/01/01"`) gives *1*<br>
  Example: `=NEPALIMONTHNUMBER("2078/11/09")` gives *11*
  

#### Installation
1. Download <a href="https://github.com/pbhattarai2017/excel-addins/raw/main/NepaliMonthNames.xlam" target="_blank" download>NepaliMonthName.xlam</a>
2. Copy **NepaliMonthName.xlam** to *C:\Users\\%USERNAME%\AppData\Roaming\Microsoft\Excel\XLSTART*
3. Restart excel program.
4. Now, you can use the functions like, `NEPALIMONTH(miti)` and `NEPALIMONTHNUMBER(miti)`
---

To sort the month names in a **Pivot Table**; the following custom list is needed:

* Baishakh, Jestha, Ashard, Shrawan, Bhadra, Asoj, Kartik, Mangsir, Poush, Magh, Falgun, Chaitra
