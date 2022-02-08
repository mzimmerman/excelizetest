package main

import (
	"bytes"
	"fmt"
	"io"
	"os"
	"strconv"
	"testing"

	"github.com/mzimmerman/xlsxwriter"
)

func WriteExcel(data [][]string, out io.Writer) error {
	if len(data) == 0 {
		return nil // nothing to write, no error
	}
	file, err := xlsxwriter.New(out)
	if err != nil {
		return err
	}
	defer file.Close()
	return file.WriteLines(data)
}

// func WriteExcelStream(data [][]string, out io.Writer) error {
// 	file := excelize.NewFile()
// 	if len(data) == 0 {
// 		return nil // nothing to write, no error
// 	}
// 	sw, err := file.NewStreamWriter("Sheet1")
// 	if err != nil {
// 		return err
// 	}
// 	dataIn := make(chan []interface{})
// 	go func() {
// 		defer close(dataIn)
// 		for _, line := range data {
// 			lineInterface := make([]interface{}, len(line))
// 			for x := range line {
// 				lineInterface[x] = line[x]
// 			}
// 			dataIn <- lineInterface
// 		}
// 	}()
// 	cell, _ := excelize.CoordinatesToCellName(1, 1)
// 	err = sw.SetRows(context.Background(), cell, dataIn)
// 	if err != nil {
// 		return err
// 	}
// 	sw.Flush()
// 	_, err = file.WriteTo(out)
// 	return err
// }

func BenchmarkExcelize10x10(b *testing.B) {
	benchmarkExcelize(10, 10, b)
}

func BenchmarkExcelize100x100(b *testing.B) {
	benchmarkExcelize(100, 100, b)
}

func BenchmarkExcelize1000x1000(b *testing.B) {
	benchmarkExcelize(1000, 1000, b)
}

func BenchmarkExcelize10000x10000(b *testing.B) {
	benchmarkExcelize(10000, 10000, b)
}

func BenchmarkExcelize1000x10(b *testing.B) {
	benchmarkExcelize(1000, 10, b)
}

func BenchmarkExcelize10000x10(b *testing.B) {
	benchmarkExcelize(10000, 10, b)
}

func BenchmarkExcelize100000x10(b *testing.B) {
	benchmarkExcelize(100000, 10, b)
}

func BenchmarkExcelize100000x100(b *testing.B) {
	benchmarkExcelize(100000, 100, b)
}

func BenchmarkExcelize10000x1000(b *testing.B) {
	benchmarkExcelize(10000, 1000, b)
}

func benchmarkExcelize(rows, cols int, b *testing.B) {
	buf := bytes.Buffer{}
	for n := 0; n < b.N; n++ {
		b.StopTimer()
		buf.Reset()
		count := 0
		data := make([][]string, rows)
		for x := range data {
			data[x] = make([]string, cols)
			for y := range data[x] {
				data[x][y] = strconv.Itoa(count)
				count++
			}
		}
		b.StartTimer()
		err := WriteExcel(data, &buf)
		if err != nil {
			b.Fatalf("error writing excel - %v", err)
		}
	}
}

func TestExcel(t *testing.T) {
	data := [][]string{{"hi", "there"}, {"yes", "no"}}
	fo, err := os.Create(fmt.Sprintf("tmp-%s.xlsx", "test"))
	if err != nil {
		t.Fatalf("error - %v", err)
		return
	}
	err = WriteExcel(data, fo)
	if err != nil {
		t.Fatalf("error - %v", err)
		return
	}
	fo.Close()
}
