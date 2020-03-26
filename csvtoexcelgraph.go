package main

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

	"github.com/lxn/walk"
	"github.com/lxn/walk/declarative"
	"github.com/tanaton/CSVToExcelGraph/format"
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
	yaxislist []int
	colmap    map[int]struct{}
}

type Column struct {
	Title          string `json:",omitempty"`
	YAxisSecondary bool   `json:",omitempty"`
}

type Config struct {
	XAxis      string
	XAxisTitle string
	Columns    map[string]Column

	cdir     string
	current  string
	namelist []string
	namemap  map[string]string
}

type MyMainWindow struct {
	*walk.MainWindow
	xcComboConfigList *walk.ComboBox
	xcEditConfig      *walk.TextEdit
	xcImage           *walk.ImageView
	xcEditLog         *walk.TextEdit
	config            *Config
	converting        bool
}

var log *zap.SugaredLogger

func init() {
	updateLogger(os.Stderr)
}

func main() {
	defer func() {
		if err := recover(); err != nil {
			log.Warnw("panic!!!", "error", err)
			os.Exit(1)
		}
	}()
	os.Exit(_main())
}

func _main() int {
	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()

	// カレントディレクトリ取得
	cdir, err := os.Getwd()
	if err != nil {
		log.Warnw("カレントディレクトリの取得に失敗", "error", err)
		return 1
	}
	if len(os.Args) > 1 {
		// CUIで動作
		err = cui(ctx, cdir, os.Args[1])
	} else {
		// GUI起動
		err = gui(ctx, cdir)
	}
	if err != nil {
		log.Warnw("ツール起動に失敗", "error", err)
		return 1
	}
	return 0
}

func updateLogger(out io.Writer) {
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

func cui(ctx context.Context, cdir, rp string) error {
	config := NewConfig(cdir)
	err := config.Load()
	if err != nil {
		return err
	}
	if _, err := config.CreateGraph(rp); err != nil {
		log.Warnw("グラフ生成失敗", "error", err)
	} else {
		log.Infow("グラフ生成成功")
	}
	return nil
}

func gui(ctx context.Context, cdir string) error {
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
	updateLogger(mmw)
	defer log.Sync()
	// ダイアログオブジェクト生成
	dialog := declarative.MainWindow{
		AssignTo: &mmw.MainWindow,
		Title:    "CSVToExcelGraph",
		Size:     declarative.Size{Width: 800, Height: 600},
		Layout:   declarative.VBox{},
		OnDropFiles: func(files []string) {
			if mmw.converting {
				log.Infow("グラフ生成中に新しくドロップされたファイルは無視されます。")
			} else if checkExtList(files, ".csv") {
				mmw.converting = true
				num := runtime.NumCPU()
				log.Infow("ファイルがドロップされました。", "ファイル数", len(files), "並列数", num)
				// メッセージループを止めないようにgoroutineを起動させる
				go func(files []string, num int) {
					c := make(chan struct{}, num)
					defer func() {
						close(c)
						mmw.converting = false
					}()
					for _, file := range files {
						c <- struct{}{}
						go func(file string) {
							defer func() {
								<-c
							}()
							mmw.CreateGraph(file)
						}(file)
					}
					for num > 0 {
						num--
						c <- struct{}{}
					}
				}(files, num)
			} else {
				log.Infow("csv以外の拡張子のファイルはグラフ化できません。")
			}
		},
		Children: []declarative.Widget{
			declarative.ComboBox{
				AssignTo:     &mmw.xcComboConfigList,
				Model:        mmw.config.GetNameList(),
				CurrentIndex: 0, // 初期値
				OnCurrentIndexChanged: func() {
					txt := mmw.xcComboConfigList.Text()
					mmw.ReadConfigName(txt)
				},
			},
			declarative.HSplitter{
				Children: []declarative.Widget{
					declarative.TextEdit{
						AssignTo:  &mmw.xcEditConfig,
						VScroll:   true,
						MaxLength: 0x7FFFFFFF,
						Text:      rCRLF.Replace(mmw.config.Text()),
					},
					declarative.ImageView{
						AssignTo:   &mmw.xcImage,
						Background: declarative.SolidColorBrush{Color: walk.RGB(255, 191, 0)},
						Margin:     3,
						Mode:       declarative.ImageViewModeZoom,
						MinSize:    declarative.Size{Width: 580, Height: 300},
					},
				},
			},
			declarative.TextEdit{
				AssignTo:  &mmw.xcEditLog,
				VScroll:   true,
				MaxLength: 0x7FFFFFFF,
			},
			declarative.PushButton{
				Text: "ファイルを選択する",
				OnClicked: func() {
					dlg := &walk.FileDialog{}
					dlg.FilePath = ""
					dlg.Title = "グラフ化したいCSVを選択してください"
					dlg.Filter = "Exe files (*.csv)|*.csv|All files (*.*)|*.*"
					if ok, err := dlg.ShowOpen(mmw); err != nil {
						log.Warnw("選択したファイルが開けませんでした。", "error", err)
					} else if !ok {
						log.Infow("キャンセルしました。")
					} else {
						rp := dlg.FilePath
						log.Infow("選択", "path", rp)
						mmw.CreateGraph(rp)
					}
				},
			},
		},
	}
	_, err = dialog.Run()
	return err
}

func (mmw *MyMainWindow) Write(b []byte) (n int, err error) {
	str := rCRLF.Replace(string(b))
	// メッセージループが重い処理等で停止していると制御が戻らなくなるので注意
	// Synchronizeを挟むと回避できるが、メッセージループが停止する事自体が問題なのでそっちで気を付ける
	//mmw.MainWindow.Synchronize(func() { mmw.xcEditLog.AppendText(str) })
	mmw.xcEditLog.AppendText(str)
	return len(b), nil
}

func (mmw *MyMainWindow) ReadConfigName(name string) {
	mmw.config.SetCurrent(name)
	err := mmw.config.Load()
	if err != nil {
		mmw.xcEditConfig.SetText("error")
		log.Warnw("設定ファイル読み込み異常", "error", err)
	} else {
		mmw.xcEditConfig.SetText(rCRLF.Replace(mmw.config.Text()))
		log.Infow("設定ファイルの読み込みができました。")
	}
}

func (mmw *MyMainWindow) CreateGraph(rp string) {
	log.Infow("グラフ生成開始", "path", rp)
	if ip, err := mmw.config.CreateGraph(rp); err != nil {
		log.Warnw("グラフ生成異常", "error", err)
	} else {
		img, err := walk.NewImageFromFileForDPI(ip, 96)
		if err != nil {
			log.Warnw("グラフ画像生成失敗", "error", err)
		} else {
			mmw.xcImage.SetImage(img)
			log.Infow("グラフ変換が完了しました。")
		}
	}
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
func (c Config) CreateGraph(rp string) (string, error) {
	dir, name := filepath.Split(rp)
	csvname := strings.TrimRight(name, filepath.Ext(rp)) + "_graph.csv"
	dp, _ := filepath.Abs(filepath.Join(dir, csvname))
	// 間引き
	g, err := c.ReduceCSV(rp, dp)
	if err != nil {
		return "", err
	}
	wp := dp + ".xlsx"
	ip := wp + ".png"
	// スレッドを固定する
	runtime.LockOSThread()
	// グラフ描画
	err = format.Excelgraph(dp, wp, ip, g.yaxislist)
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
func (c Config) ReduceCSV(rp, wp string) (*Graph, error) {
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
	g := &Graph{}
	g.colmap = make(map[int]struct{})
	g.yaxislist = []int{}

	r := bufio.NewScanner(rfp)
	// ヘッダー
	if r.Scan() {
		line := r.Text()
		headers := []string{}
		cells := strings.Split(line, ",")
		xaxis := c.XAxis
		if xaxis == "" {
			xaxis = "A"
		}
		for i, cell := range cells {
			fc := formatColumn(uint64(i))
			col, ok := c.Columns[fc]
			if fc == xaxis {
				g.colmap[i] = struct{}{}
				if c.XAxisTitle != "" {
					cell = c.XAxisTitle
				}
				headers = append(headers, cell)
			} else if ok {
				g.colmap[i] = struct{}{}
				if col.YAxisSecondary {
					g.yaxislist = append(g.yaxislist, len(headers))
				}
				if col.Title != "" {
					cell = col.Title
				}
				headers = append(headers, cell)
			}
		}
		hmax = len(cells)
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
		for i, it := range cells {
			if _, ok := g.colmap[i]; ok {
				lines = append(lines, it)
			}
		}
		fmt.Fprint(wfp, strings.Join(lines, ",")+"\r\n")
	}
	return g, nil
}

const digits = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

// Go標準ライブラリ[src/strconv/itoa.go formatBits]から丸パクリ
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
