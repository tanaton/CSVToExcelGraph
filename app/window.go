package app

import (
	"context"
	"flag"
	"runtime"
	"strings"
	"sync"

	"github.com/lxn/walk"
	"github.com/lxn/walk/declarative"
	"github.com/tanaton/CSVToExcelGraph/app/config"
	"go.uber.org/atomic"
)

// MyMainWindow メインウインドウ構造体
type MyMainWindow struct {
	*walk.MainWindow
	xcComboConfigList *walk.ComboBox
	xcEditConfig      *walk.TextEdit
	xcImage           *walk.ImageView
	xcEditLog         *walk.TextEdit
	conf              *config.Config
	converting        *atomic.Bool
	onInitDialogOnce  sync.Once
}

// AppTitle アプリケーション名
const AppTitle = "CSVToExcelGraph"

var rCRLF = strings.NewReplacer(
	"\r\n", "\r\n",
	"\n", "\r\n",
)

// CreateDialog ダイアログ生成
func (mmw *MyMainWindow) CreateDialog(ctx context.Context) error {
	// ダイアログオブジェクト生成
	dialog := declarative.MainWindow{
		AssignTo:    &mmw.MainWindow,
		Title:       declarative.Bind("OnInitDialog()"),
		Size:        declarative.Size{Width: 1024, Height: 768},
		Layout:      declarative.VBox{},
		OnDropFiles: mmw.OnDropFiles,
		Functions: map[string]func(args ...interface{}) (interface{}, error){
			"OnInitDialog": mmw.OnInitDialog,
		},
		Children: []declarative.Widget{
			declarative.ComboBox{
				AssignTo:     &mmw.xcComboConfigList,
				Model:        mmw.conf.GetNameList(),
				CurrentIndex: mmw.conf.GetCurretIndex(), // 初期値
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
						Text:      rCRLF.Replace(mmw.conf.Text()),
					},
					declarative.ImageView{
						AssignTo:   &mmw.xcImage,
						Background: declarative.SolidColorBrush{Color: walk.RGB(255, 217, 0)},
						Margin:     3,
						Mode:       declarative.ImageViewModeZoom,
						MinSize:    declarative.Size{Width: 806, Height: 430},
					},
				},
			},
			declarative.TextEdit{
				AssignTo:  &mmw.xcEditLog,
				VScroll:   true,
				MaxLength: 0x7FFFFFFF,
				HScroll:   true,
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
					} else if mmw.converting.Load() {
						log.Infow("グラフ生成中に新しく選択されたファイルは無視されます。")
					} else {
						rp := dlg.FilePath
						log.Infow("選択", "path", rp)
						mmw.converting.Store(true)
						go func() {
							defer mmw.converting.Store(false)
							mmw.CreateGraph(rp)
						}()
					}
				},
			},
		},
	}
	// メッセージループ起動
	_, err := dialog.Run()
	return err
}

// OnInitDialog ダイアログ生成時に呼び出される
func (mmw *MyMainWindow) OnInitDialog(args ...interface{}) (interface{}, error) {
	// OnInitDialog的なものの定義
	// MainWindowのTitleプロパティを借りて動作するので文字列を返す必要あり
	mmw.onInitDialogOnce.Do(func() {
		rp := flag.CommandLine.Arg(0)
		if rp != "" && mmw.converting.Load() == false {
			// 引数ありで起動された場合＆グラフ変換中ではない場合
			mmw.MainWindow.Synchronize(func() {
				// Synchronizeで後回しにすると上手く動く
				log.Infow("起動時引数を元に変換を開始します。", "path", rp)
				mmw.converting.Store(true)
				go func() {
					defer mmw.converting.Store(false)
					mmw.CreateGraph(rp)
				}()
			})
		}
	})
	return AppTitle, nil
}

func (mmw *MyMainWindow) Write(b []byte) (n int, err error) {
	// メッセージループが重い処理等で停止していると制御が戻らなくなるので注意
	// Synchronizeを挟むと回避できるが、メッセージループが停止する事自体が問題なのでそっちで気を付ける
	//mmw.MainWindow.Synchronize(func() { mmw.xcEditLog.AppendText(str) })
	mmw.xcEditLog.AppendText(string(b))
	return len(b), nil
}

// OnDropFiles ドロップイベントハンドラ
func (mmw *MyMainWindow) OnDropFiles(files []string) {
	if mmw.converting.Load() {
		log.Infow("グラフ生成中に新しくドロップされたファイルは無視されます。")
	} else if checkExtList(files, ".csv") {
		num := runtime.NumCPU()
		log.Infow("ファイルがドロップされました。", "ファイル数", len(files), "並列数", num)
		// メッセージループを止めないようにgoroutineを起動させる
		mmw.converting.Store(true)
		go func(files []string, num int) {
			c := make(chan struct{}, num)
			defer func() {
				close(c)
				mmw.converting.Store(false)
			}()
			for _, file := range files {
				c <- struct{}{}
				// ある程度並列で動作させる
				go func(file string) {
					defer func() {
						<-c
					}()
					// グラフ化実行
					mmw.CreateGraph(file)
				}(file)
			}
			// 並列で動作している処理の待機
			for num > 0 {
				num--
				c <- struct{}{}
			}
		}(files, num)
	} else {
		log.Infow("csv以外の拡張子のファイルはグラフ化できません。")
	}
}

// ReadConfigName 設定ファイル読み込み
func (mmw *MyMainWindow) ReadConfigName(name string) {
	err := mmw.conf.SetCurrent(name)
	if err != nil {
		log.Warnw("指定されたファイルがありません。", "error", err)
	} else {
		err = mmw.conf.Load()
		if err != nil {
			mmw.xcEditConfig.SetText("error")
			log.Warnw("設定ファイル読み込み異常", "error", err)
		} else {
			mmw.xcEditConfig.SetText(rCRLF.Replace(mmw.conf.Text()))
			log.Infow("設定ファイルの読み込みができました。")
		}
	}
}

// CreateGraph GUI側グラフ生成関数読み出し
func (mmw *MyMainWindow) CreateGraph(rp string) {
	log.Infow("グラフ生成開始", "path", rp)
	if ip, err := CreateGraph(mmw.conf, rp); err != nil {
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
