package xlsxparser

import (
	"log"
	"math"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

type Parser struct {
	file *excelize.File
}

func NewParser(fileName string) (Parser, error) {
	f, err := excelize.OpenFile(fileName)
	return Parser{file: f}, err
}

type productRow struct {
	ProductName string  `xlsx:"0"`
	Category    string  `xlsx:"1"`
	Price       float64 `xlsx:"2"`
	StockLevel  int     `xlsx:"3"`
}

type productSheet struct {
	productRows []productRow
	errors      []error
}

type productTypes map[string]productSheet

func (p *Parser) ParseFile() (productTypes, bool) {
	sheetMap := p.file.GetSheetMap()
	productTypesList := make(map[string]productSheet, 0)

	hasErrs := false
	for _, sheetKey := range sheetMap {
		var sheet productSheet
		products, errList := p.parseSheet(sheetKey)
		if len(errList) > 0 {
			hasErrs = true
		}
		sheet.productRows = products
		sheet.errors = errList
		productTypesList[sheetKey] = sheet
	}

	return productTypesList, hasErrs
}

func (p *Parser) parseSheet(sheetKey string) ([]productRow, []error) {
	errList := make([]error, 0)
	// Get all the rows in the Sheet1.
	rows := p.file.GetRows(sheetKey)
	products := make([]productRow, 0)
	for rowNumber, row := range rows {
		product := productRow{}

		if row[0] == "Product Name" {
			continue
		}
		product.ProductName = row[0]
		product.Category = row[1]
		price, err := strconv.ParseFloat(row[2], 64)
		if err != nil {
			log.Printf("can't parse price on line: %d, %s", rowNumber, err.Error())
			errList = append(errList, err)
		}
		product.Price = price
		stockLevelFloat, err := strconv.ParseFloat(row[3], 64)
		if err != nil {
			log.Printf("can't parse stock level on line: %d, %s", rowNumber, err.Error())
			errList = append(errList, err)
		}

		product.StockLevel = int(math.Round(stockLevelFloat))
		products = append(products, product)
		log.Printf("%v", product)
	}
	return products, errList
}
