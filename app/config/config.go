package config

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
)

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

func (c Config) GetCurretIndex() int {
	for i, name := range c.namelist {
		if name == c.current {
			return i
		}
	}
	return 0
}

func (c *Config) SetCurrent(name string) error {
	p, ok := c.namemap[name]
	if !ok {
		return fmt.Errorf("ファイルがありません。filename：%s", name)
	}
	c.current = p
	return nil
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
