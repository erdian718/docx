package docx

import (
	"archive/zip"
	"bytes"
	"io"
	"io/ioutil"
	"os"
)

// File is a docx file.
type File struct {
	files map[string]([]byte)
}

// OpenFile opens a docx.File from a file.
func OpenFile(name string) (*File, error) {
	zr, err := zip.OpenReader(name)
	if err != nil {
		return nil, err
	}
	defer zr.Close()
	return OpenZip(&zr.Reader)
}

// OpenReader opens a docx.File from a reader.
func OpenReader(r io.Reader) (*File, error) {
	buf, err := ioutil.ReadAll(r)
	if err != nil {
		return nil, err
	}
	zr, err := zip.NewReader(bytes.NewReader(buf), int64(len(buf)))
	if err != nil {
		return nil, err
	}
	return OpenZip(zr)
}

// OpenZip opens a docx.File from a zip reader.
func OpenZip(zr *zip.Reader) (*File, error) {
	files := make(map[string]([]byte))
	for _, f := range zr.File {
		body, err := readZipFile(f)
		if err != nil {
			return nil, err
		}
		files[f.Name] = body
	}
	return &File{
		files: files,
	}, nil
}

// Document gets the document data.
func (a *File) Document() []byte {
	return a.files["word/document.xml"]
}

// SetDocument sets the document data.
func (a *File) SetDocument(body []byte) {
	a.files["word/document.xml"] = body
}

// WriteFile writes the docx to a file.
func (a *File) WriteFile(name string) error {
	f, err := os.Create(name)
	if err != nil {
		return err
	}
	defer f.Close()
	if err := a.Write(f); err != nil {
		return err
	}
	return f.Sync()
}

// Write writes the docx to a writer.
func (a *File) Write(w io.Writer) error {
	return a.WriteZip(zip.NewWriter(w))
}

// WriteZip writes the docx to a zip writer.
func (a *File) WriteZip(zw *zip.Writer) error {
	for name, body := range a.files {
		w, err := zw.Create(name)
		if err != nil {
			return err
		}
		if _, err := w.Write(body); err != nil {
			return err
		}
	}
	return zw.Close()
}

func readZipFile(f *zip.File) ([]byte, error) {
	r, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer r.Close()
	return ioutil.ReadAll(r)
}
