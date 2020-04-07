package app

import (
	"context"
	"flag"
	"fmt"
	"image"
	"image/png"
	"io"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"testing"

	"github.com/tanaton/CSVToExcelGraph/app/config"
	"github.com/tanaton/CSVToExcelGraph/app/graph"
	"go.uber.org/atomic"
	"go.uber.org/zap"
	"go.uber.org/zap/zapcore"
)

// Newline 改行コードの指定
const Newline = "\r\n"

var pngEncoder = png.Encoder{
	CompressionLevel: png.BestCompression,
	BufferPool:       newPngPool(),
}

// CSV csv用データ構造
type CSV struct {
	hmax        int
	linenum     int
	columnlist  []int
	secondaries []int
	reduceFunc  func(linenum int, cells []string) bool
	bufcolumns  [256]string
}

var log *zap.SugaredLogger

var (
	confpath = flag.String("c", "config.json", "設定ファイルのパス")
)

func init() {
	testing.Init() // flag.Parseによりテストフラグが処理されてテストが失敗するので追加
	flag.Parse()
}

// Cui CUI用メイン処理
func Cui(ctx context.Context, cdir string) error {
	defer log.Sync()
	c := config.NewConfig(cdir)
	c.SetCurrent(*confpath)
	err := c.Load()
	if err != nil {
		return err
	}
	rp := flag.CommandLine.Arg(0)
	if rp == "" {
		return fmt.Errorf("CSVファイルの指定がありません。")
	}
	if _, err := CreateGraph(c, rp); err != nil {
		log.Warnw("グラフ生成失敗", "error", err)
	} else {
		log.Infow("グラフ生成成功")
	}
	return nil
}

// Gui GUI用メイン処理
func Gui(ctx context.Context, cdir string) error {
	c := config.NewConfig(cdir)
	c.SetCurrent(*confpath)
	err := c.Load()
	if err != nil {
		log.Warnw("設定ファイルの読み込みに失敗しました。", "error", err)
		return err
	}
	mmw := &MyMainWindow{
		conf:       c,
		converting: atomic.NewBool(false),
	}
	// ログ更新
	UpdateLogger(mmw)
	defer log.Sync()
	return mmw.CreateDialog(ctx)
}

// UpdateLogger ロガーの出力先を変更する
func UpdateLogger(out io.Writer) {
	if log != nil {
		log.Sync()
	}
	c := zap.NewDevelopmentEncoderConfig()
	c.LineEnding = Newline
	logger := zap.New(zapcore.NewCore(
		zapcore.NewConsoleEncoder(c),
		zapcore.AddSync(out),
		zap.InfoLevel,
	))
	log = logger.Sugar()
}

// GetLog アプリ部で使用するロガーの取得
func GetLog() *zap.SugaredLogger {
	return log
}

func checkExtList(list []string, ext string) bool {
	if len(list) <= 0 {
		return false
	}
	for _, name := range list {
		if ext != filepath.Ext(name) {
			return false
		}
	}
	return true
}

// CreateGraph グラフ生成メイン処理呼び出し
func CreateGraph(c *config.Config, rp string) (string, error) {
	dir, name := filepath.Split(rp)
	csvname := strings.TrimRight(name, filepath.Ext(rp)) + "_graph.csv"
	dp, _ := filepath.Abs(filepath.Join(dir, csvname))
	// 間引き
	csv, err := reduceCSV(c, rp, dp)
	if err != nil {
		return "", err
	}
	wp := dp + ".xlsx"
	ip := wp + ".png"
	// スレッドを固定する
	runtime.LockOSThread()
	// グラフ描画
	err = graph.Excelgraph(dp, wp, ip, csv.secondaries)
	// スレッドの固定を解除する（※ゴールーチンを抜けると自動でアンロックされる）
	runtime.UnlockOSThread()
	if err != nil {
		return "", err
	}
	// 一時ファイルの削除
	err = os.Remove(dp)
	if err != nil {
		return "", err
	}
	// 生成されたpngの圧縮率が微妙なので再圧縮
	err = regenePNG(ip)
	if err != nil {
		return "", err
	}
	return ip, nil
}

func reduceCSV(c *config.Config, rp, wp string) (*CSV, error) {
	swc, err := NewScanWriteCloser(rp, wp)
	if err != nil {
		return nil, err
	}
	defer swc.Close()
	csv := NewCSVReducer(c)
	// ヘッダー
	err = csv.scanHeader(swc, c)
	if err != nil {
		return nil, err
	}
	// データ
	err = csv.scanData(swc)
	if err != nil {
		return nil, err
	}
	return csv, nil
}

// NewCSVReducer CSV間引き用構造体生成
func NewCSVReducer(c *config.Config) *CSV {
	var f func(int, []string) bool
	if c.ReduceRows > 0 {
		f = func(rows int) func(int, []string) bool {
			return func(linenum int, _ []string) bool {
				if linenum%rows == 0 {
					return false
				}
				// trueで間引く
				return true
			}
		}(c.ReduceRows)
	}
	return &CSV{
		hmax:        0,
		linenum:     0,
		columnlist:  make([]int, 0, len(c.YColumns)+1),
		secondaries: make([]int, 0, len(c.YColumns)+1),
		reduceFunc:  f,
	}
}

func (csv *CSV) scanHeader(swc ScanWriteCloser, c *config.Config) error {
	// ヘッダー
	if swc.Scan() == false {
		return swc.Err()
	}
	csv.linenum++
	cl := append([]config.Column{c.XColumn}, c.YColumns...)
	cells := strings.Split(swc.Text(), ",")
	csv.hmax = len(cells)
	swc.WriteString(csv.headerString(cells, cl) + Newline)
	return nil
}

func (csv *CSV) headerString(cells []string, cl []config.Column) string {
	for i, it := range cl {
		col := int(parseColumn(it.Axis))
		if col >= csv.hmax {
			log.Infow(
				"設定で指定された列番号が存在しません",
				"設定の列指定", it.Axis,
				"設定の列指定数値", col,
				"読み込んだCSVの最右列", formatColumn(uint64(csv.hmax)),
				"読み込んだCSVの最右列数値", csv.hmax,
			)
			continue
		}
		cell := cells[col]
		if it.AxisTitle != "" {
			cell = it.AxisTitle
		}
		csv.bufcolumns[i] = cell
		csv.columnlist = append(csv.columnlist, col)
		if i > 0 && it.AxisSecondary {
			csv.secondaries = append(csv.secondaries, i)
		}
	}
	return strings.Join(csv.bufcolumns[:len(csv.columnlist)], ",")
}

func (csv *CSV) scanData(swc ScanWriteCloser) error {
	if csv.linenum <= 0 {
		return fmt.Errorf("CSVのヘッダーを読み込んでいません。")
	}
	for swc.Scan() {
		csv.linenum++
		cells := strings.Split(swc.Text(), ",")
		if len(cells) < csv.hmax-1 {
			return fmt.Errorf("csvの区切り文字数が最初より少なくなりました。ヘッダーの区切り文字数:%d, %d行目の区切り文字数:%d", csv.hmax, csv.linenum, len(cells))
		}
		if csv.reduceFunc != nil && csv.reduceFunc(csv.linenum, cells) {
			// 間引く行の場合
		} else {
			swc.WriteString(csv.dataString(cells) + Newline)
		}
	}
	return nil
}

func (csv *CSV) dataString(cells []string) string {
	for i, it := range csv.columnlist {
		csv.bufcolumns[i] = cells[it]
	}
	return strings.Join(csv.bufcolumns[:len(csv.columnlist)], ",")
}

// Go標準ライブラリ[src/strconv/itoa.go formatBits]を参考に改造
const digits = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

func formatColumn(u uint64) string {
	var a [64 + 1]byte
	i := len(a)
	for u >= 26 {
		i--
		q := u / 26
		a[i] = digits[uint(u-q*26)]
		u = q - 1 // Z→AAに繰り上がる場合、10進数で9→00になるのと同じなので改造
	}
	i--
	a[i] = digits[uint(u)]
	return string(a[i:])
}

// Go標準ライブラリ[src/strconv/atoi.go ParseUint]を参考に改造
const intSize = 32 << (^uint(0) >> 63)
const maxUint64 = 1<<64 - 1

func lower(c byte) byte {
	return c | ('x' - 'X')
}
func parseColumn(s string) uint64 {
	if s == "" {
		return 0
	}
	bitSize := int(intSize)
	maxVal := uint64(1)<<uint(bitSize) - 1
	var cutoff uint64 = maxUint64/26 + 1
	var n uint64
	for i, c := range []byte(s) {
		var d byte
		switch {
		case 'a' <= lower(c) && lower(c) <= 'z':
			d = lower(c) - 'a'
		default:
			return 0
		}
		if d >= 26 {
			return 0
		}

		if n >= cutoff {
			// n*base overflows
			return maxVal
		}
		// AA=26とした場合、10進数で00=10とするのと同じなので改造
		if i >= 1 {
			n = (n + 1) * 26
		} else {
			n *= 26
		}

		n1 := n + uint64(d)
		if n1 < n || n1 > maxVal {
			// n+v overflows
			return maxVal
		}
		n = n1
	}
	return n
}

func regenePNG(ip string) error {
	img, err := openImage(ip)
	if err != nil {
		return err
	}
	wp, err := os.Create(ip)
	if err != nil {
		return err
	}
	defer wp.Close()
	return pngEncoder.Encode(wp, img)
}

func openImage(ip string) (image.Image, error) {
	rp, err := os.Open(ip)
	if err != nil {
		return nil, err
	}
	defer rp.Close()
	return png.Decode(rp)
}

type pngPool struct {
	pool sync.Pool
}

func newPngPool() *pngPool {
	return &pngPool{
		pool: sync.Pool{
			New: func() interface{} {
				return &png.EncoderBuffer{}
			},
		},
	}
}
func (p *pngPool) Get() *png.EncoderBuffer {
	return p.pool.Get().(*png.EncoderBuffer)
}
func (p *pngPool) Put(buf *png.EncoderBuffer) {
	p.pool.Put(buf)
}
