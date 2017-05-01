# convert-quicken-report-to-excel
Convert a [Quicken itemized report](/example/itemized_categories_example.TXT) to an [excel document](/example/itemized_categories_example.TXT.xls) with transacations broken down by months.  

Excel output example with 2D line graph:  

![excep_output_with_graph](https://cloud.githubusercontent.com/assets/7596215/25567995/383bcaac-2dc7-11e7-9b12-9b098f214768.png)

## requires
* JDK 8+
* Groovy 2+

## usage
quickenReport2Excel.groovy can be run as a shell script:
```bash
quickenReport2Excel.groovy -f itemized_categories.TXT
```
## notes
Works with Quicken Home & Business 2014 and 2016
