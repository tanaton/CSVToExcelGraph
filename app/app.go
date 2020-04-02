package app

import (
	"bufio"
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

var pngEncoder = png.Encoder{
	CompressionLevel: png.BestCompression,
	BufferPool:       NewPngPool(),
}

type CSV struct {
	raww        io.WriteCloser
	rawr        io.ReadCloser
	w           *bufio.Writer
	scanner     *bufio.Scanner
	finish      bool
	hmax        int
	linenum     int
	columnlist  []int
	secondaries []int
	reduceFunc  func(linenum int, cells []string) bool
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
	c.LineEnding = "\r\n"
	logger := zap.New(zapcore.NewCore(
		zapcore.NewConsoleEncoder(c),
		zapcore.AddSync(out),
		zap.InfoLevel,
	))
	log = logger.Sugar()
}

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

func CreateGraph(c *config.Config, rp string) (string, error) {
	dir, name := filepath.Split(rp)
	csvname := strings.TrimRight(name, filepath.Ext(rp)) + "_graph.csv"
	dp, _ := filepath.Abs(filepath.Join(dir, csvname))
	// 間引き
	csv, err := ReduceCSV(c, rp, dp)
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
	os.Remove(dp)
	if err != nil {
		return "", err
	}
	// 生成されたpngの圧縮率が微妙なので再圧縮
	err = regenePNG(ip)
	return ip, err
}

func ReduceCSV(c *config.Config, rp, wp string) (*CSV, error) {
	csv, err := NewCSVReducer(c, rp, wp)
	if err != nil {
		return nil, err
	}
	defer csv.Close()
	// ヘッダー
	err = csv.scanHeader(c)
	if err != nil {
		return nil, err
	}
	// データ
	err = csv.scanData()
	if err != nil {
		return nil, err
	}
	return csv, nil
}

func NewCSVReducer(c *config.Config, rp, wp string) (*CSV, error) {
	rawr, err := os.Open(rp)
	if err != nil {
		return nil, err
	}
	raww, werr := os.Create(wp)
	if werr != nil {
		rawr.Close()
		return nil, werr
	}
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
		raww:        raww,
		rawr:        rawr,
		w:           bufio.NewWriterSize(raww, 128*1024),
		scanner:     bufio.NewScanner(rawr),
		finish:      false,
		hmax:        0,
		linenum:     0,
		columnlist:  make([]int, 1, len(c.Columns)+1),
		secondaries: make([]int, 0, len(c.Columns)),
		reduceFunc:  f,
	}, nil
}

func (csv *CSV) Close() error {
	var err error
	if csv.w != nil {
		e := csv.w.Flush()
		if e != nil {
			err = e
		}
		csv.w = nil
	}
	if csv.raww != nil {
		e := csv.raww.Close()
		if e != nil {
			err = e
		}
		csv.raww = nil
	}
	if csv.rawr != nil {
		e := csv.rawr.Close()
		if e != nil {
			err = e
		}
		csv.rawr = nil
	}
	return err
}

func (csv *CSV) scanHeader(c *config.Config) error {
	xaxis := int(parseColumn(c.XAxis))
	// ヘッダー
	if csv.scanner.Scan() {
		csv.linenum++
		headers := make([]string, 1, len(c.Columns)+1)
		cells := strings.Split(csv.scanner.Text(), ",")
		csv.hmax = len(cells)
		// X軸設定
		if xaxis < csv.hmax {
			cell := cells[xaxis]
			if c.XAxisTitle != "" {
				cell = c.XAxisTitle
			}
			headers[0] = cell
			csv.columnlist[0] = xaxis
		} else {
			return fmt.Errorf("X軸設定が変です。xaxis:%d, cell_num:%d", xaxis, csv.hmax)
		}
		// Y軸設定
		for i, it := range c.Columns {
			col := int(parseColumn(it.YAxis))
			if col >= csv.hmax {
				log.Infow(
					"設定で指定された列番号が存在しません",
					"設定の列指定", it.YAxis,
					"設定の列指定数値", col,
					"読み込んだCSVの最右列", formatColumn(uint64(csv.hmax)),
					"読み込んだCSVの最右列数値", csv.hmax,
				)
				continue
			}
			cell := cells[col]
			if it.YAxisSecondary {
				csv.secondaries = append(csv.secondaries, i+1)
			}
			if it.YAxisTitle != "" {
				cell = it.YAxisTitle
			}
			headers = append(headers, cell)
			csv.columnlist = append(csv.columnlist, col)
		}
		fmt.Fprint(csv.w, strings.Join(headers, ",")+"\r\n")
	} else {
		return csv.scanner.Err()
	}
	return nil
}

func (csv *CSV) scanData() error {
	if csv.linenum <= 0 {
		return fmt.Errorf("CSVのヘッダーを読み込んでいません。")
	}
	columns := make([]string, len(csv.columnlist))
	for csv.scanner.Scan() {
		csv.linenum++
		cells := strings.Split(csv.scanner.Text(), ",")
		if len(cells) < csv.hmax-1 {
			return fmt.Errorf("csvの区切り文字数が最初より少なくなりました。ヘッダーの区切り文字数:%d, %d行目の区切り文字数:%d", csv.hmax, csv.linenum, len(cells))
		}
		if csv.reduceFunc != nil && csv.reduceFunc(csv.linenum, cells) {
			// 間引く行の場合
		} else {
			for i, it := range csv.columnlist {
				columns[i] = cells[it]
			}
			fmt.Fprint(csv.w, strings.Join(columns, ",")+"\r\n")
		}
	}
	csv.finish = true
	return nil
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

type PngPool struct {
	pool sync.Pool
}

func NewPngPool() *PngPool {
	return &PngPool{
		pool: sync.Pool{
			New: func() interface{} {
				return &png.EncoderBuffer{}
			},
		},
	}
}
func (p *PngPool) Get() *png.EncoderBuffer {
	return p.pool.Get().(*png.EncoderBuffer)
}
func (p *PngPool) Put(buf *png.EncoderBuffer) {
	p.pool.Put(buf)
}
