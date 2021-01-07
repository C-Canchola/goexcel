module github.com/C-Canchola/goexcel

go 1.15

require github.com/360EntSecGroup-Skylar/excelize/v2 v2.3.1

replace (
	github.com/C-Canchola/goexcel/parse => ./parse
    github.com/C-Canchola/goexcel/schema => ./schema
    github.com/C-Canchola/goexcel/writing => ./writing
)