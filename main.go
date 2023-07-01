package main

import (
	"encoding/csv"
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type round struct {
	Name   string
	Date   time.Time
	Scores []string
	Total  int
	Hits   int
	Golds  int
	Bow    string
}

func main() {
	files, err := os.ReadDir("/Users/robert/Documents/Archery/old_scores")
	if err != nil {
		log.Fatal(err)
	}

	rounds := make([]round, 0)

	for _, file := range files {
		f, err := os.Open("/Users/robert/Documents/Archery/old_scores/" + file.Name())
		if err != nil {
			log.Fatal(err)
		}
		defer f.Close()

		csvr := csv.NewReader(f)
		csvr.FieldsPerRecord = -1

		lines, err := csvr.ReadAll()
		if err != nil {
			log.Fatal(err)
		}

		record := round{}
		scores := false
		for _, line := range lines {
			switch line[0] {
			case "Scoresheet for":
				record.Name = line[1]
				continue
			case "Shot on":
				date, err := time.Parse("02/01/2006", line[1])
				if err != nil {
					log.Fatal("unable to convert date", err)
				}
				record.Date = date
				if record.Date.After(time.Date(2023, 01, 01, 0, 0, 0, 0, time.Local)) {
					record.Bow = "Elite Verdict"
				} else {
					record.Bow = "hoyt Razortec"
				}
				continue
			case "Grand Total:":
				scores = false
				total, err := strconv.Atoi(line[1])
				if err != nil {
					log.Fatal("unable to convert total", err)
				}
				record.Total = total
				continue
			case "Number Of Hits:":
				hits, err := strconv.Atoi(line[1])
				if err != nil {
					log.Fatal("unable to convert hits", err)
				}
				record.Hits = hits
				continue
			case "Number Of Whites:", "Number Of Golds:":
				golds, err := strconv.Atoi(line[1])
				if err != nil {
					log.Fatal("unable to convert golds", err)
				}
				record.Golds = golds
				continue
			case "Scores:":
				scores = true
				continue
			}
			if scores {
				// parse scores
				if line[0] == "Ends at" {
					continue
				}
				endScore := make([]string, 0)

				if len(line) == 8 || len(line) == 9 || len(line) == 10 || len(line) == 11 {
					endScore = append(endScore, line[2:7]...)
				} else {
					continue
				}

				if len(endScore) != 0 {
					record.Scores = append(record.Scores, "[\""+strings.Join(endScore, "\",\"")+"\"]")
				}
			}
		}
		rounds = append(rounds, record)
	}

	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			log.Fatal(err)
		}
	}()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "bow")
	f.SetCellValue(sheet, "B1", "date")
	f.SetCellValue(sheet, "C1", "hits")
	f.SetCellValue(sheet, "D1", "golds")
	f.SetCellValue(sheet, "E1", "round")
	f.SetCellValue(sheet, "F1", "score")
	f.SetCellValue(sheet, "G1", "sheet")

	for i, round := range rounds {
		r := fmt.Sprint(i + 2)

		f.SetCellValue(sheet, "A"+r, round.Bow)
		f.SetCellValue(sheet, "B"+r, round.Date.Format("2006-01-02"))
		f.SetCellValue(sheet, "C"+r, round.Hits)
		f.SetCellValue(sheet, "D"+r, round.Golds)
		f.SetCellValue(sheet, "E"+r, round.Name)
		f.SetCellValue(sheet, "F"+r, round.Total)
		f.SetCellValue(sheet, "G"+r, "["+strings.Join(round.Scores, ",")+"]")
	}

	if err := f.SaveAs("output.xlsx"); err != nil {
		log.Fatal(err)
	}
}
