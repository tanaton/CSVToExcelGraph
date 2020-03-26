package graph

import (
	"errors"
	"fmt"
	"path/filepath"
	"strings"

	ole "github.com/go-ole/go-ole"
	"github.com/tanaton/go-ole-msoffice/excel"
)

type ExcelGraph struct {
	obj *excel.Application
}

type GraphItem struct {
	x     int
	count int
	rg    *excel.Range
	leg   []string
}

// 終了処理
func (ex *ExcelGraph) quit() {
	ex.obj.Quit()
	ex.obj.Release()
}

// 既存のブックを開く
func (ex *ExcelGraph) openFile(p string) *excel.Workbook {
	workbooks := ex.obj.GetWorkbooks()
	workbook := workbooks.Open(p)
	return workbook
}

// ブックを保存
func (ex *ExcelGraph) saveBook(book *excel.Workbook, p string) {
	ex.obj.SetDisplayAlerts(false)
	book.SaveAs(p, excel.XlWorkbookDefault)
}

// グラフの作成
func (ex *ExcelGraph) createChartObject(sheet *excel.Worksheet, x, y, xx, yy int, name string) *excel.ChartObject {
	chs := sheet.ChartObjects()
	g := chs.Add(x, y, xx, yy)
	g.SetName(name)
	return g
}

// グラフに描画するためのデータを取得
func (ex *ExcelGraph) getGraphRange(sheet *excel.Worksheet) (string, []GraphItem) {
	x := 1
	xname := ""
	arr := []GraphItem{}
	maxCol := sheet.GetCells().GetItem(1, sheet.GetColumns().GetCount()).GetEnd(excel.XlToLeft).GetColumn() + 1

	// 気合で探索
	for {
		v := sheet.GetCells().GetItem(1, x).GetValue()
		if v.Value() == nil || sheet.Err != nil {
			break
		}
		if x == 1 {
			// X軸の名称を取得
			xname = v.ToString()
		}

		// 列の探索
		leg := []string{}
		i := x + 1
		for ; i < maxCol; i++ {
			v := sheet.GetCells().GetItem(1, i).GetValue()
			if v.Value() == nil || sheet.Err != nil {
				break
			} else {
				leg = append(leg, v.ToString())
			}
		}

		item := GraphItem{
			x:     x,                       // データ開始位置
			count: ((i - 1) - (x + 1)) + 1, // データの列数
			rg:    sheet.GetRange(sheet.GetColumns().GetItem(x+1), sheet.GetColumns().GetItem(i-1)),
			leg:   leg,
		}
		arr = append(arr, item)
		x = i + 1
	}
	return xname, arr
}

// シートからグラフを作る
func (ex *ExcelGraph) sheetToChart(g *excel.ChartObject, sheet *excel.Worksheet, secondary []int) {
	j := 1
	xname, arr := ex.getGraphRange(sheet)
	if len(arr) <= 0 {
		fmt.Println("シートにグラフ化できるデータが無いみたい")
		return
	}

	priname := []string{}
	secname := []string{}
	sec := map[int]struct{}{}
	for _, it := range secondary {
		sec[it] = struct{}{}
	}
	// 一つのレンジにまとめる
	union := arr[0].rg
	for i := 1; i < len(arr); i++ {
		union = ex.obj.Union(union, arr[i].rg)
	}

	chart := g.GetChart()
	// データの設定
	chart.SetSourceData(union, excel.XlColumns)
	// グラフの種類を設定
	chart.SetChartType(excel.XlXYScatterLinesNoMarkers)
	// 凡例の位置を修正
	legend := chart.GetLegend()
	legend.SetPosition(excel.XlLegendPositionBottom)
	// 要素の設定
	for _, it := range arr {
		xcell := sheet.GetCells().GetItem(2, it.x)
		for k := 1; k <= it.count; k++ {
			// 線ごとにX軸の設定
			sc := chart.SeriesCollection().Item(j)
			if _, ok := sec[j]; ok {
				// 2軸
				sc.SetAxisGroup(excel.XlSecondary)
				secname = append(secname, it.leg[k-1])
			} else {
				priname = append(priname, it.leg[k-1])
			}
			end := xcell.GetEnd(excel.XlDown)
			rg := sheet.GetRange(xcell, end)
			sc.SetXValues(rg)
			j++
		}
	}
	// グラフの軸についての設定
	ex.setGraphAxis(g, xname, strings.Join(priname, " / "))
	// 指定した要素を第二軸へ移動
	if len(secondary) > 0 {
		ex.setGraphAxisSecondary(g, strings.Join(secname, " / "))
	}
}

// グラフの軸を設定
func (ex *ExcelGraph) setGraphAxis(g *excel.ChartObject, xname, yname string) {
	chart := g.GetChart()
	cp := chart.Axes(excel.XlCategory, excel.XlPrimary)
	vp := chart.Axes(excel.XlValue, excel.XlPrimary)

	// X軸の目盛線の表示
	cp.SetHasMajorGridlines(true)
	cp.SetHasMinorGridlines(true)
	// Y軸の目盛線の表示
	vp.SetHasMinorGridlines(true)
	// 目盛線の位置を下に移動
	cp.SetTickLabelPosition(excel.XlTickLabelPositionLow)
	// X軸ラベルを表示
	cp.SetHasTitle(true)
	cp.GetAxisTitle().SetText(xname)
	// Y軸ラベルを表示
	vp.SetHasTitle(true)
	vp.GetAxisTitle().SetText(yname)
}

// 指定した要素を第二軸に移動
func (ex *ExcelGraph) setGraphAxisSecondary(g *excel.ChartObject, name string) {
	chart := g.GetChart()
	// 2軸目のY軸ラベルを表示
	vs := chart.Axes(excel.XlValue, excel.XlSecondary)
	vs.SetHasTitle(true)
	at := vs.GetAxisTitle()
	at.SetText(name)
}

// タイトルを設定
func (ex *ExcelGraph) setGraphTitle(g *excel.ChartObject, title string) {
	chart := g.GetChart()
	chart.SetHasTitle(true)
	ct := chart.GetChartTitle()
	ct.SetText(title)
	ct.SetPosition(excel.XlChartElementPositionAutomatic)
	ct.SetIncludeInLayout(false) // タイトルをグラフと重ねる
}

// グラフを画像にして保存
func (ex *ExcelGraph) saveGraphImage(chart *excel.Chart, ip string) {
	chart.Export(ip, "PNG")
}

// グラフを新しいシートに移動
func (ex *ExcelGraph) moveNewGraphSheet(g *excel.ChartObject, name string) *excel.Chart {
	chart := g.GetChart()
	return chart.Location(excel.XlLocationAsNewSheet, name)
}

// スクリーン更新停止
func (ex *ExcelGraph) lockScreen() {
	ex.obj.SetScreenUpdating(false)
}

// スクリーン更新許可
func (ex *ExcelGraph) unlockScreen() {
	ex.obj.SetScreenUpdating(true)
}

func Excelgraph(rp, wp, ip string, secondary []int) (err error) {
	// COMの初期化
	ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED|ole.COINIT_DISABLE_OLE1DDE)
	// 確実に行う必要があるため
	defer ole.CoUninitialize()

	// エクセルオブジェクトの生成
	e := excel.ThisApplication()
	if e == nil {
		err = errors.New("Excelの起動に失敗しました")
		return
	}
	ex := ExcelGraph{obj: e}
	// Excelを閉じてリソース解放
	defer ex.quit()
	// 生成してる感を出すためアプリケーションを表示する
	//ex.obj.SetVisible(true)
	// 既存のブックの読み込み
	book := ex.openFile(rp)
	// シートの取得
	sheet := book.GetWorksheets().GetItem(1)

	ex.lockScreen()
	// 空グラフの生成
	graph := ex.createChartObject(sheet, 30, 30, 500, 300, "csvexcelgraph")
	// シート内容をグラフに変換
	ex.sheetToChart(graph, sheet, secondary)
	// タイトルを設定
	_, name := filepath.Split(rp)
	ex.setGraphTitle(graph, strings.TrimSuffix(name, filepath.Ext(name)))
	// グラフオブジェクトをグラフシートに移動
	chart := ex.moveNewGraphSheet(graph, "Graph1")
	ex.unlockScreen()

	if ip != "" {
		// グラフを画像で保存
		ex.saveGraphImage(chart, ip)
	}
	// ブックを保存
	ex.saveBook(book, wp)
	// ブックを閉じる
	book.Close()
	return
}
