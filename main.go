package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"io/ioutil"
	"os"
	"strconv"
	"strings"
	"log"
	"time"
)

var (
	choiceAnswer = []string{"D", "D", "D", "B", "A", "C", "A", "C", "A", "C", "D", "B", "C", "A", "B", "C", "C", "B", "B", "C"}
	blankAnswer  = []string{"extends", "implements", "StringBuilder", "ArrayList", "m.put(\"语文\",90)"}
	fexamDir     = "./fexam/"
)

func right(stuAnswer []string, keyAnswer []string) int {
	rightNum := 0
	for i := 0; i < len(stuAnswer); i++ {
		if strings.EqualFold(stuAnswer[i], keyAnswer[i]) {
			rightNum = rightNum + 1
		}
	}
	return rightNum
}

func main() {
	var examFiles []string
	var scoreFile *xlsx.File
	var row *xlsx.Row
	files, err := ioutil.ReadDir(fexamDir)
	if err != nil {
		log.Println("Error reading fexam directory!")
		os.Exit(1)
	}
	for _, file := range files {
		filename := file.Name()
		if strings.HasSuffix(filename, ".xlsx") {
			examFiles = append(examFiles, fexamDir+filename)
		}
	}
	//create score xlsx file to save student score
	scoreFile = xlsx.NewFile()
	scoreSheet, err := scoreFile.AddSheet("Sheet1")
	if err != nil {
		log.Println(err.Error())
	}
	//loop student exam xlsx file ,read sheet1 col2 to judge
	startTime := time.Now().UnixNano()
	for _, examFile := range examFiles {
		score := 0
		xlFile, error := xlsx.OpenFile(examFile) //open excel file
		if error != nil {
			log.Println("Error reading examfile")
		}
		//get student's name,no,class
		//example:-Unlicensed-13715050_袁慧敏_13医器_java.xlsx
		s := strings.Split(examFile[20:strings.Index(examFile, ".xlsx")], "_")
		stuNo := s[0]
		stuName := s[1]
		stuClass := s[2]
		sheet := xlFile.Sheets[0]
		var stuAnswer []string
		for _, row := range sheet.Rows {
			if row != nil {
				cell := row.Cells[2]
				stuAnswer = append(stuAnswer, fmt.Sprintf("%s", cell.String()))
			}
		}
		fmt.Println(stuAnswer)
		//get the choice right number
		rightChoiceNum := right(stuAnswer[1:21], choiceAnswer)
		//get the blank right number
		rightBlankNum := right(stuAnswer[22:27], blankAnswer)
		score = rightChoiceNum*3 + rightBlankNum*5
		fmt.Println(score)
		row = scoreSheet.AddRow()
		//add student no
		cellstuno := row.AddCell()
		cellstuno.Value = stuNo
		//add student name
		cellname := row.AddCell()
		cellname.Value = stuName
		//add student class
		cellclass := row.AddCell()
		cellclass.Value = stuClass
		//add student score
		cellscore := row.AddCell()
		cellscore.Value = strconv.Itoa(score)
	}
	err = scoreFile.Save("./score/score.xlsx")
	endTime := time.Now().UnixNano()
	log.Printf("试卷批改完成 共加载:%d 条记录, 所花时间:%.1f ms\n", len(examFiles), float64(endTime-startTime)/1000000)
	if err != nil {
		log.Println(err.Error())
	}

}
