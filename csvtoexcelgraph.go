package main

import (
	"context"
	"log"
	"os"
	"path/filepath"

	"github.com/shirou/gopsutil/process"
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

	// exeファイルが置いてあるディレクトリの取得
	// ファイルドロップで起動した場合、ファイルドロップ元がカレントディレクトリになるため対応
	cdir, err := getExePath(ctx)
	if err != nil {
		log.Println("カレントディレクトリの取得に失敗", err)
		return 1
	}
	if isCUI(ctx) {
		// CUIで動作
		err = app.Cui(ctx, cdir)
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

// exeファイルのディレクトリを取得
func getExePath(ctx context.Context) (string, error) {
	proc, err := process.NewProcess(int32(os.Getpid()))
	if err != nil {
		return "", err
	}
	exe, err := proc.ExeWithContext(ctx)
	if err != nil {
		return "", err
	}
	return filepath.Dir(exe), nil
}

// CUI判定
func isCUI(ctx context.Context) bool {
	proc, err := process.NewProcess(int32(os.Getppid()))
	if err != nil {
		return false
	}
	name, err := proc.NameWithContext(ctx)
	if err != nil {
		return false
	}
	return name == "cmd.exe"
}
