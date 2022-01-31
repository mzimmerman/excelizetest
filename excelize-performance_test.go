package main

import (
	"bytes"
	"io"
	"strconv"
	"testing"

	"github.com/xuri/excelize/v2"
)

func WriteExcel(data [][]string, out io.Writer) error {
	file := excelize.NewFile()
	if len(data) == 0 {
		return nil // nothing to write, no error
	}
	sw, err := file.NewStreamWriter("Sheet1")
	if err != nil {
		return err
	}
	lineInterface := make([]interface{}, len(data[0]))
	for excelLineNum, line := range data {
		lineInterface = lineInterface[:0]
		for x := range line {
			lineInterface = append(lineInterface, line[x])
		}
		cell, _ := excelize.CoordinatesToCellName(1, excelLineNum+1)
		err = sw.SetRow(cell, lineInterface)
		if err != nil {
			return err
		}
	}
	_, err = file.WriteTo(out)
	return err
}

func BenchmarkExcelize10x10(b *testing.B) {
	benchmarkExcelize(10, 10, b)
}

func BenchmarkExcelize100x100(b *testing.B) {
	benchmarkExcelize(100, 100, b)
}

func BenchmarkExcelize1000x1000(b *testing.B) {
	benchmarkExcelize(1000, 1000, b)
}

// func BenchmarkExcelize10000x10000(b *testing.B) {
// 	benchmarkExcelize(10000, 10000, b)
// }

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
	// run the Fib function b.N times
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
