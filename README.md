# xlsx
struct save as xlsx

```go
package main

import (
	"github.com/ruyewuxin/xlsx"
)

type employee struct {
	ID    int    `xlsx:"工号"`
	Name  string `xlsx:"姓名"`
	Email string `xlsx:"邮箱"`
}

func main() {
	var es []interface{}
	e1 := employee{1, "张三", "zhangsan@***.com"}
	e2 := employee{2, "李四", "lisi@***.com"}
	e3 := employee{3, "王五", "wangwu@***.com"}
	es = append(es, e1, e2, e3)

	var x1 xlsx.Xlsx
	x1.XlsxName = "Book1.xlsx"
	x1.SheetName = "文档1"
	x1.Headers = xlsx.GetHeaders(es)
	x1.Rows = xlsx.GetRows(es)
	x1.ToXlsx()
}

```
