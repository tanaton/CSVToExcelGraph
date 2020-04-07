package app

import (
	"bufio"
	"io"
	"os"
)

const writeBuffSize = 128 * 1024

// ScanWriteCloser 読み書き用
type ScanWriteCloser interface {
	io.StringWriter
	io.Writer
	io.Closer
	Err() error
	Scan() bool
	Text() string
}

type scannerWriter struct {
	*bufio.Scanner
	*bufio.Writer
	raww io.WriteCloser
	rawr io.ReadCloser
}

// NewScanWriteCloser ScanWriteCloser生成用
func NewScanWriteCloser(rp, wp string) (ScanWriteCloser, error) {
	rawr, err := os.Open(rp)
	if err != nil {
		return nil, err
	}
	raww, werr := os.Create(wp)
	if werr != nil {
		rawr.Close()
		return nil, werr
	}
	return &scannerWriter{
		Writer:  bufio.NewWriterSize(raww, writeBuffSize),
		Scanner: bufio.NewScanner(rawr),
		raww:    raww,
		rawr:    rawr,
	}, nil
}

func (rw *scannerWriter) Close() error {
	var err error
	if rw.Writer != nil {
		e := rw.Writer.Flush()
		if e != nil {
			err = e
		}
		rw.Writer = nil
	}
	if rw.raww != nil {
		e := rw.raww.Close()
		if e != nil {
			err = e
		}
		rw.raww = nil
	}
	if rw.rawr != nil {
		e := rw.rawr.Close()
		if e != nil {
			err = e
		}
		rw.rawr = nil
	}
	return err
}
