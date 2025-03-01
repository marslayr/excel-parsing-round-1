package main

import (
	"errors"
	"fmt"
	"log"
	"math"
	"os"
	"sort"
	"strconv"

	"github.com/xuri/excelize/v2"
)

type Student struct {
	Index       int
	ClassNumber int
	EmpID       string
	CampusID    string
	Quiz        float64
	Midsem      float64
	LabTest     float64
	WeeklyLabs  float64
	PreCompre   float64
	Compre      float64
	Total       float64
	Branch      string
}

func main() {
	if len(os.Args) != 2 {
		log.Fatal("Usage: go run main.go <path-to-excel-file>")
	}

	excelPath := os.Args[1]
	file, err := excelize.OpenFile(excelPath)
	if err != nil {
		log.Fatalf("Failed to open Excel file: %v", err)
	}
	defer file.Close()

	sheetName := file.GetSheetName(0)
	rows, err := file.GetRows(sheetName)
	if err != nil {
		log.Fatalf("Failed to read sheet: %v", err)
	}

	var students []Student
	headers := map[string]int{}

	for i, row := range rows {
		if i == 0 {
			// Mapping column indexes
			for j, col := range row {
				headers[col] = j
			}
			continue
		}

		if len(row) < 11 {
			continue
		}

		studentIndex, _ := strconv.Atoi(row[headers["Sl No"]])
		classNumber, _ := strconv.Atoi(row[headers["Class No."]])
		empID := row[headers["EmplID"]]
		campusID := row[headers["Campus ID"]]
		quiz, _ := strconv.ParseFloat(row[headers["Quiz (30)"]], 64)
		midSem, _ := strconv.ParseFloat(row[headers["Mid-Sem (75)"]], 64)
		labTest, _ := strconv.ParseFloat(row[headers["Lab Test (60)"]], 64)
		weeklyLabs, _ := strconv.ParseFloat(row[headers["Weekly Labs (30)"]], 64)
		preCompre, _ := strconv.ParseFloat(row[headers["Pre-Compre (195)"]], 64)
		compre, _ := strconv.ParseFloat(row[headers["Compre (105)"]], 64)
		total, _ := strconv.ParseFloat(row[headers["Total (300)"]], 64)
		branch, _ := parseBranch(campusID)

		// fmt.Printf("%.2f\n", total)
		students = append(students, Student{studentIndex, classNumber, empID, campusID, quiz, midSem, labTest, weeklyLabs, preCompre, compre, total, branch})

	}

	mismatchChan := computeMismatch(students)
	avgChan := computeAverages(students)
	branchAvgChan := computeBranchAverages(students)
	topRankingsChan := computeTopRankings(students)

	// Print results as they become available
	fmt.Println(<-mismatchChan)
	fmt.Println(<-avgChan)
	fmt.Println(<-branchAvgChan)
	fmt.Println(<-topRankingsChan)
}

func parseBranch(campusID string) (string, error) {
	if campusID[4] == 'A' {
		branch := campusID[4:6]
		return branch, nil
	} else if campusID[4] == 'B' {
		branch := campusID[4:9]
		return branch, nil
	} else {
		return "", errors.New("invalid campus ID")
	}
}

func computeMismatch(students []Student) <-chan string {
	result := make(chan string)
	go func() {
		var mismatch string
		for _, s := range students {
			computedTotal := s.Quiz + s.Midsem + s.LabTest + s.WeeklyLabs + s.Compre
			if math.Abs(computedTotal-s.Total) > 1e-3 {
				mismatch += fmt.Sprintf("Mismatch in total for Campus ID %s: Computed %f, Found %f\n", s.CampusID, computedTotal, s.Total)
			}
		}
		result <- mismatch
	}()
	return result
}

func computeAverages(students []Student) <-chan string {
	result := make(chan string)

	go func() {
		var totalSum, quizSum, midSemSum, labSum, weeklyLabSum, preCompreSum, compreSum float64
		count := len(students)

		for _, s := range students {
			totalSum += s.Total
			quizSum += s.Quiz
			midSemSum += s.Midsem
			labSum += s.LabTest
			weeklyLabSum += s.WeeklyLabs
			preCompreSum += s.PreCompre
			compreSum += s.Compre
		}

		result <- fmt.Sprintf("General Averages:\nQuiz: %.2f\nMid Sem: %.2f\nLab Test: %.2f\nWeekly Labs: %.2f\nPre-Compre: %.2f\nCompre: %.2f\nOverall Average: %.2f\n",
			float64(quizSum)/float64(count),
			float64(midSemSum)/float64(count),
			float64(labSum)/float64(count),
			float64(weeklyLabSum)/float64(count),
			float64(preCompreSum)/float64(count),
			float64(compreSum)/float64(count),
			float64(totalSum)/float64(count))
	}()
	return result
}

// func printStudents(students []Student) {
// 	for _, s := range students {
// 		fmt.Printf("%+v\n", s)
// 	}
// }

func computeBranchAverages(students []Student) <-chan string {
	result := make(chan string)
	go func() {
		branchTotals := make(map[string]float64)
		branchCounts := make(map[string]float64)

		for _, s := range students {
			if len(s.Branch) == 2 {
				branchTotals[s.Branch] += s.Total
				branchCounts[s.Branch]++
			}
		}

		var output string
		output += "\nBranch-wise Averages for 2024 batch:\n"
		for branch, sum := range branchTotals {
			output += fmt.Sprintf("%s: %.2f\n", branch, float64(sum)/float64(branchCounts[branch]))
		}
		result <- output
	}()
	return result
}

func computeTopRankings(students []Student) <-chan string {

	result := make(chan string)
	go func() {
		var output string
		studentsCopy := make([]Student, len(students))
		copy(studentsCopy, students)
		categories := map[string]func(Student) float64{
			"Quiz":       func(s Student) float64 { return s.Quiz },
			"MidSem":     func(s Student) float64 { return s.Midsem },
			"LabTest":    func(s Student) float64 { return s.LabTest },
			"WeeklyLabs": func(s Student) float64 { return s.WeeklyLabs },
			"PreCompre":  func(s Student) float64 { return s.PreCompre },
			"Compre":     func(s Student) float64 { return s.Compre },
		}

		for category, getMarks := range categories {
			sort.Slice(studentsCopy, func(i, j int) bool {
				if getMarks(studentsCopy[i]) == getMarks(studentsCopy[j]) {
					return studentsCopy[i].CampusID < studentsCopy[j].CampusID
				}
				return getMarks(studentsCopy[i]) > getMarks(studentsCopy[j])
			})

			output += fmt.Sprintf("\nTop 3 Students in %s:\n", category)
			var rank int = 1
			prevMarks := -1.0
			for i := 0; i < 3 && i < len(studentsCopy); i++ {
				marks := getMarks(studentsCopy[i])

				if marks > prevMarks {
					rank = i + 1
				}
				output += fmt.Sprintf("%d. Campus ID: %s, Marks: %.2f\n", rank, studentsCopy[i].CampusID, marks)
				prevMarks = marks
			}
		}
		result <- output
	}()
	return result
}
