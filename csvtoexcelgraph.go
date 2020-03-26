package main

import (
	"context"
	"log"
	"os"

	"github.com/tanaton/CSVToExcelGraph/app"
)

func init() {
	app.UpdateLogger(os.Stderr)
}

func main() {
	defer func() {
		if err := recover(); err != nil {
			log.Println("panic!!!", err)
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
		log.Println("カレントディレクトリの取得に失敗", err)
		return 1
	}
	if len(os.Args) > 1 {
		// CUIで動作
		err = app.Cui(ctx, cdir, os.Args[1])
	} else {
		// GUI起動
		err = app.Gui(ctx, cdir)
	}
	if err != nil {
		log.Println("ツール起動に失敗", err)
		return 1
	}
	return 0
}
