package main

import (
	"bufio"
	"errors"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/schollz/progressbar/v3"
	"golang.org/x/exp/slog"
)

var (
	ignore = flag.String("g", "_", "ExcelでPDF作成対象外とするシート名の先頭文字")
)

var (
	ErrOpenFile   = errors.New("ファイルのオープンに失敗しました。")
	ErrConvertPdf = errors.New("PDFファイルへの変換に失敗しました。")
)

type ConsoleOutput struct {
	filetype int
	log      string
}

const (
	MsoTriStateMsoFalse = 0
	MsoTriStateMsoTrue  = -1
)

func main() {
	if len(os.Args) < 2 {
		slog.Error("引数を指定してください。")
		slog.Error("Usage: Office2PDF.exe [対象フォルダ]")
		os.Exit(1)
	}

	flag.Usage = usage
	flag.Parse()
	args := flag.Args()
	if len(args) == 0 {
		slog.Error("PDF変換対象フォルダのパスを指定してください。")
		slog.Error("Usage: Office2PDF.exe [対象フォルダ]")
		os.Exit(1)
	}

	targetPath := args[0]

	// 処理対象フォルダから、PDF変換対象ファイルの一覧を取得する。
	xlsPaths, docPaths, pptPaths, err := getFilePaths(targetPath)
	if err != nil {
		slog.Error("ファイル一覧の取得に失敗しました。", err, "path", targetPath)
		os.Exit(1)
	}
	// PDFに変換するファイルが存在しない場合は、処理終了。
	if len(xlsPaths) == 0 && len(docPaths) == 0 && len(pptPaths) == 0 {
		slog.Info("PDF変換対象フィルが存在しません。", "path", targetPath)
		return
	}

	fmt.Printf("Excel: %d件, Word: %d件, PowerPoint: %d件のファイルを変換します。\n", len(xlsPaths), len(docPaths), len(pptPaths))

	wg := sync.WaitGroup{}
	errChan := make(chan error, len(xlsPaths)+len(docPaths)+len(pptPaths))

	// Excel変換のゴルーチン
	wg.Add(1)
	go func() {
		defer wg.Done()
		if len(xlsPaths) > 0 {
			fmt.Println("Excel変換の進行状況:")
			bar := progressbar.NewOptions(len(xlsPaths),
				progressbar.OptionEnableColorCodes(true),
				progressbar.OptionShowCount(),
				progressbar.OptionSetWidth(50),
				progressbar.OptionSetDescription("[cyan]Excel→PDF[reset]"),
				progressbar.OptionSetTheme(progressbar.Theme{
					Saucer:        "[green]=[reset]",
					SaucerHead:    "[green]>[reset]",
					SaucerPadding: " ",
					BarStart:      "[",
					BarEnd:        "]",
				}))

			if err := convertExcelFileToPdf(xlsPaths, *ignore, bar); err != nil {
				errChan <- err
			}
		}
	}()

	// Word変換のゴルーチン
	wg.Add(1)
	go func() {
		defer wg.Done()
		if len(docPaths) > 0 {
			fmt.Println("Word変換の進行状況:")
			bar := progressbar.NewOptions(len(docPaths),
				progressbar.OptionEnableColorCodes(true),
				progressbar.OptionShowCount(),
				progressbar.OptionSetWidth(50),
				progressbar.OptionSetDescription("[blue]Word→PDF[reset]"),
				progressbar.OptionSetTheme(progressbar.Theme{
					Saucer:        "[green]=[reset]",
					SaucerHead:    "[green]>[reset]",
					SaucerPadding: " ",
					BarStart:      "[",
					BarEnd:        "]",
				}))

			if err := convertWordFileToPdf(docPaths, bar); err != nil {
				errChan <- err
			}
		}
	}()

	// PowerPoint変換のゴルーチン
	wg.Add(1)
	go func() {
		defer wg.Done()
		if len(pptPaths) > 0 {
			fmt.Println("PowerPoint変換の進行状況:")
			bar := progressbar.NewOptions(len(pptPaths),
				progressbar.OptionEnableColorCodes(true),
				progressbar.OptionShowCount(),
				progressbar.OptionSetWidth(50),
				progressbar.OptionSetDescription("[magenta]PPT→PDF[reset]"),
				progressbar.OptionSetTheme(progressbar.Theme{
					Saucer:        "[green]=[reset]",
					SaucerHead:    "[green]>[reset]",
					SaucerPadding: " ",
					BarStart:      "[",
					BarEnd:        "]",
				}))

			if err := convertPptFileToPdf(pptPaths, bar); err != nil {
				errChan <- err
			}
		}
	}()

	wg.Wait()
	close(errChan)

	flag := true
	for err := range errChan {
		if flag {
			slog.Error("PDF変換でエラーが発生しました。")
			flag = false
		}
		slog.Error("error", err)
	}

	if !flag {
		fmt.Print("エラーが発生しました。何かキーを押してください。\n")
		scanner := bufio.NewScanner(os.Stdin)
		scanner.Scan()
		os.Exit(1)
	} else {
		fmt.Println("すべての変換処理が完了しました！")
	}
}

// PowerPointファイルをPDFに変換する。
func convertPptFileToPdf(files []string, bar *progressbar.ProgressBar) (rErr error) {
	if len(files) == 0 {
		return nil
	}

	// COMオブジェクトの初期化
	if err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED); err != nil {
		return err
	}
	defer ole.CoUninitialize()

	// PowerPointアプリケーションの作成
	var pptApp *ole.IDispatch
	pptApp, rErr = createPowerPointApp()
	if rErr != nil {
		return rErr
	}
	defer pptApp.Release()

	defer func() {
		_, err := oleutil.CallMethod(pptApp, "Quit")
		if err != nil {
			rErr = errors.Join(rErr, err)
		}
		slog.Info("PowerPointを終了しました.")
	}()
	slog.Info("PowerPointを起動しました.")

	for i, path := range files {
		fullpath, err := filepath.Abs(path)
		if err != nil {
			return err
		}

		// 変換元Wordファイルのパスから、PDFファイルのパス（相対パス、絶対パス）を取得する。
		pdfPath, pdfFullPath, err := getPdfPath(path)
		if err != nil {
			return err
		}

		name := filepath.Base(path)
		bar.Describe(fmt.Sprintf("[magenta]PPT→PDF[reset] (%d/%d): %s", i+1, len(files), name))

		rErr = convertPptxToPdf(pptApp, fullpath, pdfFullPath)
		if rErr != nil {
			slog.Error(name+" 変換失敗", "err", rErr, "PDFファイル", pdfPath)
			return err
		} else {
			slog.Info(name+" 変換完了", "PDFファイル", pdfPath)
		}
		bar.Add(1)
		time.Sleep(100 * time.Millisecond) // プログレスバーが更新されるのを待つ
	}

	return nil
}

// PowerPointファイルをPDFに変換する
func convertPptxToPdf(powerpoint *ole.IDispatch, pptPath, pdfFilePath string) error {
	pptname := filepath.Base(pptPath)

	pres, err := oleutil.GetProperty(powerpoint, "Presentations")
	if err != nil {
		return err
	}
	defer pres.ToIDispatch().Release()

	// PowerPointドキュメントを開く
	ppt, err := openPptFile(pres.ToIDispatch(), pptPath)
	if err != nil {
		return fmt.Errorf("%w: %s", ErrOpenFile, err.Error())
	}
	defer ppt.Release()

	slides, err := oleutil.GetProperty(ppt, "Slides")
	if err != nil {
		return err
	}
	defer slides.ToIDispatch().Release()

	count := (int)(oleutil.MustGetProperty(slides.ToIDispatch(), "Count").Val)
	slog.Info(pptname, "スライド数", count)

	ps, err := oleutil.GetProperty(ppt, "PageSetup")
	if err != nil {
		return err
	}
	defer ps.ToIDispatch().Release()

	sp := (int)(oleutil.MustGetProperty(ps.ToIDispatch(), "FirstSlideNumber").Val)
	slog.Info(pptname, "スライド開始ページ番号", sp)

	po, err := oleutil.GetProperty(ppt, "PrintOptions")
	if err != nil {
		return err
	}
	defer po.ToIDispatch().Release()

	r, err := oleutil.GetProperty(po.ToIDispatch(), "Ranges")
	if err != nil {
		return err
	}
	defer r.ToIDispatch().Release()

	pr, err := oleutil.CallMethod(r.ToIDispatch(), "Add", sp, count+(sp-1))
	if err != nil {
		return err
	}

	_, err = oleutil.CallMethod(ppt, "ExportAsFixedFormat", pdfFilePath, 2, 2, 0, 1, 1, 0, pr, 1, "", false, false, false, false, false)
	if err != nil {
		return fmt.Errorf("%w: %s", ErrConvertPdf, err.Error())
	}

	_, err = oleutil.PutProperty(ppt, "Saved", true)
	if err != nil {
		return err
	}
	_, err = oleutil.CallMethod(ppt, "Close")
	if err != nil {
		return err
	}

	return nil
}

// PowerPointのファイルをオープンする。
func openPptFile(pres *ole.IDispatch, path string) (*ole.IDispatch, error) {
	ppt, err := oleutil.CallMethod(pres, "Open", path, MsoTriStateMsoTrue, MsoTriStateMsoFalse, MsoTriStateMsoFalse)
	if err != nil {
		return nil, fmt.Errorf("%w: %s", ErrOpenFile, err.Error())
	}
	return ppt.ToIDispatch(), nil
}

func createPrintRange(pptname string, ppt *ole.VARIANT) (*ole.VARIANT, error) {
	ps, err := oleutil.GetProperty(ppt.ToIDispatch(), "PageSetup")
	if err != nil {
		return nil, err
	}
	defer ps.ToIDispatch().Release()

	slides, err := oleutil.GetProperty(ppt.ToIDispatch(), "Slides")
	if err != nil {
		return nil, err
	}
	defer slides.ToIDispatch().Release()

	count := (int)(oleutil.MustGetProperty(slides.ToIDispatch(), "Count").Val)
	slog.Info(pptname, "スライド数", count)

	sp := (int)(oleutil.MustGetProperty(ps.ToIDispatch(), "FirstSlideNumber").Val)
	slog.Info(pptname, "スライド開始ページ番号", sp)

	po, err := oleutil.GetProperty(ppt.ToIDispatch(), "PrintOptions")
	if err != nil {
		return nil, err
	}

	defer po.ToIDispatch().Release()
	r, err := oleutil.GetProperty(po.ToIDispatch(), "Ranges")
	if err != nil {
		return nil, err
	}

	defer r.ToIDispatch().Release()
	pr, err := oleutil.CallMethod(r.ToIDispatch(), "Add", sp, count+(sp-1))
	if err != nil {
		return nil, err
	}

	return pr, nil
}

// WordファイルをPDFに変換する。
func convertWordFileToPdf(files []string, bar *progressbar.ProgressBar) (rErr error) {
	if len(files) == 0 {
		return nil
	}

	// COMオブジェクトの初期化
	if err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED); err != nil {
		return err
	}
	defer ole.CoUninitialize()

	// Wordアプリケーションの作成
	var wordApp *ole.IDispatch
	wordApp, rErr = createWordApp()
	if rErr != nil {
		return rErr
	}
	defer wordApp.Release()

	defer func() {
		_, err := oleutil.CallMethod(wordApp, "Quit")
		if err != nil {
			rErr = errors.Join(rErr, err)
		}
		slog.Info("Wordを終了しました.")
	}()
	slog.Info("Wordを起動しました.")

	// Wordウィンドウを表示しないようにする
	if _, err := oleutil.PutProperty(wordApp, "Visible", false); err != nil {
		return err
	}

	for i, path := range files {
		fullpath, err := filepath.Abs(path)
		if err != nil {
			return err
		}

		// 変換元Wordファイルのパスから、PDFファイルのパス（相対パス、絶対パス）を取得する。
		pdfPath, pdfFullPath, err := getPdfPath(path)
		if err != nil {
			return err
		}

		name := filepath.Base(path)
		bar.Describe(fmt.Sprintf("[blue]Word→PDF[reset] (%d/%d): %s", i+1, len(files), name))

		rErr = convertDocxToPdf(wordApp, fullpath, pdfFullPath)
		if rErr != nil {
			slog.Error(name+" 変換失敗", "err", rErr, "PDFファイル", pdfPath)
			return err
		} else {
			slog.Info(name+" 変換完了", "PDFファイル", pdfPath)
		}
		bar.Add(1)
		time.Sleep(100 * time.Millisecond) // プログレスバーが更新されるのを待つ
	}

	return nil
}

// WordファイルをPDFに変換する
func convertDocxToPdf(word *ole.IDispatch, dcPath, pdfFilePath string) error {
	documents, err := oleutil.GetProperty(word, "documents")
	if err != nil {
		return err
	}
	defer documents.ToIDispatch().Release()

	// Wordドキュメントを開く
	doc, err := oleutil.CallMethod(documents.ToIDispatch(), "Open", dcPath)
	if err != nil {
		return err
	}
	defer doc.ToIDispatch().Release()

	// PDFに変換する
	_, err = oleutil.CallMethod(doc.ToIDispatch(), "ExportAsFixedFormat", pdfFilePath, 17)
	if err != nil {
		return err
	}

	_, err = oleutil.CallMethod(doc.ToIDispatch(), "Close", false)
	if err != nil {
		return err
	}

	return nil
}

// ExcelファイルをPDFに変換する。
func convertExcelFileToPdf(files []string, ig string, bar *progressbar.ProgressBar) (rErr error) {
	if len(files) == 0 {
		return nil
	}

	// COMオブジェクトの初期化
	if err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED); err != nil {
		return err
	}
	defer ole.CoUninitialize()

	// Excelアプリケーションの生成
	var excelApp *ole.IDispatch
	excelApp, rErr = createExcelApp()
	if rErr != nil {
		return rErr
	}
	defer excelApp.Release()
	defer func() {
		_, err := oleutil.CallMethod(excelApp, "Quit")
		if err != nil {
			rErr = errors.Join(rErr, err)
		}
		slog.Info("Excelを終了しました.")
	}()
	slog.Info("Excelを起動しました.")

	for i, path := range files {
		fullpath, err := filepath.Abs(path)
		if err != nil {
			return err
		}

		// 変換元Excelファイルのパスから、PDFファイルのパス（相対パス、絶対パス）を取得する。
		pdfPath, pdfFullPath, err := getPdfPath(path)
		if err != nil {
			return err
		}

		name := filepath.Base(path)
		bar.Describe(fmt.Sprintf("[cyan]Excel→PDF[reset] (%d/%d): %s", i+1, len(files), name))

		rErr = convertXlsxToPdf(excelApp, fullpath, pdfFullPath, ig)
		if rErr != nil {
			slog.Error(name+" 変換失敗", "err", rErr, "PDFファイル", pdfPath)
			return err
		} else {
			slog.Info(name+" 変換完了", "PDFファイル", pdfPath)
		}
		bar.Add(1)
		time.Sleep(100 * time.Millisecond) // プログレスバーが更新されるのを待つ
	}

	return nil
}

// ExcelファイルをPDFに変換する
func convertXlsxToPdf(excel *ole.IDispatch, xlPath, pdfFilePath, ig string) error {
	xlname := filepath.Base(xlPath)
	workbooks, err := oleutil.GetProperty(excel, "Workbooks")
	if err != nil {
		return err
	}
	defer workbooks.ToIDispatch().Release()
	workbook, err := oleutil.CallMethod(workbooks.ToIDispatch(), "Open", xlPath)
	if err != nil {
		return err
	}
	defer workbook.ToIDispatch().Release()

	if ig == "" {
		// PDF形式で保存
		_, err = oleutil.CallMethod(workbook.ToIDispatch(), "ExportAsFixedFormat", 0, pdfFilePath, 0, false, false)
		if err != nil {
			return err
		}
	} else {
		worksheets, err := oleutil.GetProperty(workbook.ToIDispatch(), "Worksheets")
		if err != nil {
			return err
		}
		defer worksheets.ToIDispatch().Release()

		sheetCount := (int)(oleutil.MustGetProperty(worksheets.ToIDispatch(), "Count").Val)
		slog.Info(xlname, "シート数", sheetCount)

		var worksheet *ole.IDispatch
		for i := 1; i < sheetCount+1; i++ {
			worksheet = oleutil.MustGetProperty(workbook.ToIDispatch(), "Worksheets", i).ToIDispatch()
			defer worksheet.Release()
			name := oleutil.MustGetProperty(worksheet, "Name")
			if strings.HasPrefix(name.ToString(), ig) {
				slog.Info(xlname+" シート名によりスキップ", "シート名", name.ToString())
				continue
			} else {
				_, err := oleutil.CallMethod(worksheet, "Select", false)
				if err != nil {
					return err
				}
			}
		}

		activeSheet, err := oleutil.GetProperty(workbook.ToIDispatch(), "ActiveSheet")
		if err != nil {
			return err
		}
		defer activeSheet.ToIDispatch().Release()

		_, err = oleutil.CallMethod(activeSheet.ToIDispatch(), "ExportAsFixedFormat", 0, pdfFilePath, 0, false, false)
		if err != nil {
			return err
		}
	}

	_, err = oleutil.PutProperty(workbook.ToIDispatch(), "Saved", true)
	if err != nil {
		return err
	}
	_, err = oleutil.CallMethod(workbook.ToIDispatch(), "Close", false)
	if err != nil {
		return err
	}

	return nil
}

// pathから拡張子を除いたファイル名を返す
func getFileNameWithoutExt(path string) string {
	return filepath.Base(path[:len(path)-len(filepath.Ext(path))])
}

// patshから拡張子を除いたパスを返す
func getPathWithoutExt(path string) string {
	return path[:len(path)-len(filepath.Ext(path))]
}

// Wordアプリケーションの作成
func createWordApp() (*ole.IDispatch, error) {
	if unknown, err := oleutil.CreateObject("Word.Application"); err != nil {
		return nil, err
	} else {
		wordApp, err := unknown.QueryInterface(ole.IID_IDispatch)
		if err != nil {
			return nil, err
		}
		return wordApp, nil
	}
}

// Excelアプリケーションの作成
func createExcelApp() (*ole.IDispatch, error) {
	if unknown, err := oleutil.CreateObject("Excel.Application"); err != nil {
		return nil, err
	} else {
		excelApp, err := unknown.QueryInterface(ole.IID_IDispatch)
		if err != nil {
			return nil, err
		}
		return excelApp, nil
	}
}

// PowerPointオブジェクトの生成
func createPowerPointApp() (*ole.IDispatch, error) {
	if unknown, err := oleutil.CreateObject("PowerPoint.Application"); err != nil {
		return nil, err
	} else {
		ppointApp, err := unknown.QueryInterface(ole.IID_IDispatch)
		if err != nil {
			return nil, err
		}
		return ppointApp, nil
	}
}

// folderPath で指定されたフォルダから、サブフォルダも含めたPDF変換対象ファイルの一覧を取得する。
// PDF変換対象ファイルの一覧は、Excel、Word、PowerPointに分けて、配列で返す。
func getFilePaths(folderPath string) ([]string, []string, []string, error) {
	var xslPaths, docPaths, pptPaths []string
	err := filepath.Walk(folderPath, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		// フォルダと~で始まるファイルはスキップ
		if !info.IsDir() && !strings.HasPrefix(filepath.Base(path), "~") {
			switch filepath.Ext(info.Name()) {
			case ".xlsx", ".xls":
				xslPaths = append(xslPaths, path)
			case ".docx", ".doc":
				docPaths = append(docPaths, path)
			case ".pptx", ".ppt":
				pptPaths = append(pptPaths, path)
			}
		}
		return nil
	})
	if err != nil {
		return nil, nil, nil, err
	}
	return xslPaths, docPaths, pptPaths, nil
}

func usage() {
	fmt.Println("usage: PDFConverterGO [flags] path")
	flag.PrintDefaults()
}

func getPdfPath(path string) (string, string, error) {
	pdfPath := getPathWithoutExt(path) + ".pdf"
	pdfFullPath, err := filepath.Abs(pdfPath)
	if err != nil {
		return "", "", err
	}
	return pdfPath, pdfFullPath, nil
}
