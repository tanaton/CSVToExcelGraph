package app

import (
	"bufio"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"image"
	"image/png"
	"io"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"

	"github.com/tanaton/CSVToExcelGraph/app/graph"
	"go.uber.org/zap"
	"go.uber.org/zap/zapcore"
)

var pngEncoder = png.Encoder{
	CompressionLevel: png.BestCompression,
	BufferPool:       NewPngPool(),
}
var rCRLF = strings.NewReplacer(
	"\r\n", "\r\n",
	"\n", "\r\n",
)

type Graph struct {
	xaxis       int
	secondaries []int
	colmap      map[int]struct{}
}

type Column struct {
	YAxis          string
	YAxisTitle     string `json:",omitempty"`
	YAxisSecondary bool   `json:",omitempty"`
}

type Config struct {
	XAxis      string
	XAxisTitle string `json:",omitempty"`
	Columns    []Column

	cdir     string
	current  string
	namelist []string
	namemap  map[string]string
}

var log *zap.SugaredLogger

// Cui CUI用メイン処理
func Cui(ctx context.Context, cdir, rp string) error {
	config := NewConfig(cdir)
	err := config.Load()
	if err != nil {
		return err
	}
	if _, err := CreateGraph(config, rp); err != nil {
		log.Warnw("グラフ生成失敗", "error", err)
	} else {
		log.Infow("グラフ生成成功")
	}
	return nil
}

// Gui GUI用メイン処理
func Gui(ctx context.Context, cdir string) error {
	config := NewConfig(cdir)
	err := config.Load()
	if err != nil {
		log.Warnw("設定ファイルの読み込みに失敗しました。", "error", err)
		return err
	}
	mmw := &MyMainWindow{
		config:     config,
		converting: false,
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
	logger := zap.New(zapcore.NewCore(
		zapcore.NewJSONEncoder(zap.NewProductionEncoderConfig()),
		zapcore.AddSync(out),
		zap.InfoLevel,
	))
	log = logger.Sugar()
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

func CreateGraph(c *Config, rp string) (string, error) {
	dir, name := filepath.Split(rp)
	csvname := strings.TrimRight(name, filepath.Ext(rp)) + "_graph.csv"
	dp, _ := filepath.Abs(filepath.Join(dir, csvname))
	// 間引き
	g, err := ReduceCSV(c, rp, dp)
	if err != nil {
		return "", err
	}
	wp := dp + ".xlsx"
	ip := wp + ".png"
	// スレッドを固定する
	runtime.LockOSThread()
	// グラフ描画
	err = graph.Excelgraph(dp, wp, ip, g.secondaries)
	// スレッドの固定を解除する
	runtime.UnlockOSThread()
	os.Remove(dp)
	if err != nil {
		return "", err
	}
	// 生成されたpngの圧縮率が微妙なので再圧縮
	err = regenePNG(ip)
	return ip, err
}

func ReduceCSV(c *Config, rp, wp string) (*Graph, error) {
	rfp, err := os.Open(rp)
	if err != nil {
		return nil, err
	}
	defer rfp.Close()
	wfp, werr := os.Create(wp)
	if werr != nil {
		return nil, werr
	}
	defer wfp.Close()
	// 各種初期化
	hmax := 0
	columnlist := []int{}
	g := &Graph{}
	g.colmap = make(map[int]struct{})
	g.xaxis = int(parseColumn(c.XAxis))

	r := bufio.NewScanner(rfp)
	// ヘッダー
	if r.Scan() {
		headers := []string{}
		cells := strings.Split(r.Text(), ",")
		hmax = len(cells)
		// X軸設定
		if g.xaxis < hmax {
			cell := cells[g.xaxis]
			if c.XAxisTitle != "" {
				cell = c.XAxisTitle
			}
			headers = append(headers, cell)
			columnlist = append(columnlist, g.xaxis)
		} else {
			return nil, errors.New("X軸設定が変です")
		}
		// Y軸設定
		for _, it := range c.Columns {
			col := int(parseColumn(it.YAxis))
			if col < hmax {
				cell := cells[col]
				if it.YAxisSecondary {
					g.secondaries = append(g.secondaries, len(headers))
				}
				if it.YAxisTitle != "" {
					cell = it.YAxisTitle
				}
				headers = append(headers, cell)
				columnlist = append(columnlist, col)
			}
		}
		fmt.Fprint(wfp, strings.Join(headers, ",")+"\r\n")
	} else {
		return nil, errors.New("読み取れなかったよ")
	}
	// データ
	for r.Scan() {
		cells := strings.Split(r.Text(), ",")
		if len(cells) < hmax-1 {
			return nil, errors.New("csvのカンマの数が少ないよ")
		}
		lines := []string{}
		for _, it := range columnlist {
			lines = append(lines, cells[it])
		}
		fmt.Fprint(wfp, strings.Join(lines, ",")+"\r\n")
	}
	return g, nil
}

func NewConfig(cdir string) *Config {
	c := &Config{
		cdir: cdir,
	}
	c.init()
	return c
}

func (c *Config) init() {
	rowlist, err := filepath.Glob(filepath.Join(c.cdir, "*.json"))
	if err != nil || len(rowlist) == 0 {
		log.Warnw("設定ファイルが見つかりませんでした。", "error", err)
		rowlist = []string{}
	}
	c.namelist = []string{}
	c.namemap = make(map[string]string)
	if len(rowlist) > 0 {
		c.current = rowlist[0]
		for _, path := range rowlist {
			_, name := filepath.Split(path)
			c.namelist = append(c.namelist, name)
			c.namemap[name] = path
		}
	}
}

func (c Config) GetNameList() []string {
	return c.namelist
}

func (c *Config) SetCurrent(name string) {
	p, ok := c.namemap[name]
	if ok {
		c.current = p
	} else {
		log.Warnw("選択された設定ファイルが無いよ", "filename", name)
	}
}

func (c *Config) Load() error {
	return c.ReadFile(c.current)
}

func (c *Config) ReadFile(p string) error {
	rfp, err := os.Open(p)
	if err != nil {
		return err
	}
	defer rfp.Close()
	dec := json.NewDecoder(rfp)
	return dec.Decode(c)
}
func (c Config) WriteFile(p string) error {
	wfp, err := os.Create(p)
	if err != nil {
		return err
	}
	defer wfp.Close()
	enc := json.NewEncoder(wfp)
	return enc.Encode(c)
}
func (c Config) Text() string {
	txt, _ := json.MarshalIndent(c, "", "  ")
	return string(txt)
}

// Go標準ライブラリ[src/strconv/itoa.go formatBits]から丸パクリ
const digits = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

func formatColumn(u uint64) string {
	var a [64 + 1]byte
	i := len(a)
	for u >= 26 {
		i--
		q := u / 26
		a[i] = digits[uint(u-q*26)]
		u = q
	}
	i--
	a[i] = digits[uint(u)]
	return string(a[i:])
}

// Go標準ライブラリ[src/strconv/atoi.go ParseUint]から丸パクリ
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
	for _, c := range []byte(s) {
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
		n *= 26

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
