package main

import (
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {

	if len(os.Args) < 2 {
		fmt.Println("Use: populate Usage.xlsx ListPrice.xlsx")
		return
	}
	fileUsage := os.Args[1]
	fileListPrice := os.Args[2]

	// Usage File Rows
	fu, err := excelize.OpenFile(fileUsage)
	if err != nil {
		log.Fatalf("Error to open file '%s': %v", fileUsage, err)
		return
	}
	defer fu.Close()

	fuSheetName := fu.GetSheetName(fu.GetActiveSheetIndex())
	urows, err := fu.GetRows(fuSheetName)
	if err != nil {
		log.Fatalf("Error to get rows from 'Sheet1': %v", err)
	}
	if len(urows) < 2 {
		log.Fatal("The sheet is empty or header not found.")
	}

	// List Price Rows
	lp, err := excelize.OpenFile(fileListPrice)
	if err != nil {
		log.Fatalf("Error to open file '%s': %v", fileListPrice, err)
		return
	}
	defer lp.Close()

	lpSheetName := lp.GetSheetName(lp.GetActiveSheetIndex())
	lrows, err := lp.GetRows(lpSheetName)
	if err != nil {
		log.Fatalf("Error to get rows from 'Sheet1': %v", err)
	}
	if len(lrows) < 2 {
		log.Fatal("The sheet is empty or header not found.")
	}

	ListPriceSKU := map[string]string{}
	for n, lrow := range lrows {
		if n == 0 {
			continue
		}
		lSkuId := lrow[0]
		lValue := lrow[3]
		ListPriceSKU[lSkuId] += lValue
	}

	for i, urow := range urows {
		if i == 0 { // Ignore header
			continue
		}

		// Extract data
		if len(urow) <= 32 {
			continue
		}
		uLine := i + 1
		SkuId := urow[32] // SkuId position 32
		if SkuId == "" {
			continue
		}
		PriceQuantity := urow[22] // PriceQuantity position 22
		ListPriceUnit := urow[20] // ListPriceUnit position 20
		if ListPriceUnit == "" {
			value, exist := ListPriceSKU[SkuId]
			if exist {
				valueFloat, _ := strconv.ParseFloat(value, 64)
				quantityFloat, _ := strconv.ParseFloat(PriceQuantity, 64)
				totalListCost := valueFloat * quantityFloat

				listPriceUnitValue := strings.Replace(value, ".", ",", 1)
				colName, _ := excelize.ColumnNumberToName(21) // convert 21 to U (ListUnitPrice)
				celPos := fmt.Sprintf("%s%d", colName, uLine)
				err = fu.SetCellValue(fuSheetName, celPos, listPriceUnitValue)
				if err != nil {
					log.Fatal(err)
				}

				colName, _ = excelize.ColumnNumberToName(20) // convert 20 to T (ListCost)
				celPos = fmt.Sprintf("%s%d", colName, uLine)
				err = fu.SetCellValue(fuSheetName, celPos, totalListCost)
				if err != nil {
					log.Fatal(err)
				}
				fmt.Printf("Linha vazia %d próximo a: %s \t SKU de ListPriceSKU %s\t for SKU Id %s\t CelPosition %s\n", uLine, urow[17], listPriceUnitValue, SkuId, celPos)
			}
		}
	}

	if err := fu.SaveAs("NewUsageFile.xlsx"); err != nil {
		log.Fatal(err)
	}

	fmt.Println("\nProccess Finished!")

}
