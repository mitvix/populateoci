# PopulateOCI

Utility to populate missing SKU List Prices in OCI FOCUS Cost and Usage Reports.

## Overview

This utility, written in Go, can be used to populate Oracle OCI Focus detailed report files when the reports issued by Oracle do not provide all List Price unit values in the ListUnitPrice field. It also calculates the total ListCost based on the PricingQuantity.

### Instalation

To install Golang follow the how-to available in [https://go.dev/doc/tutorial/getting-started#install](go.dev).

Installation (Linux)
```
git clone https://github.com/mitvix/populateoci
cd populateoci
go build -o populateoci main.go
```

Installation (Windows)
```
export GOOS=windows
go build -o populateoci.exe main.go
```

### Usage


```
./populateoci Usage.xlsx ListPrice.xlsx
```

### In the end

A new filename will be created with "NewUsageFile.xlsx".

### Requirement

* CUR Data: A compiled XLSX file including all line items from the Oracle OCI Cost Usage Report for the desired period.

* Pricing Reference: An XLSX file mapping all SKUs to their unit prices.

* Structural Integrity: The column order in both spreadsheets must remain consistent with the requirements defined in the Go source code.

