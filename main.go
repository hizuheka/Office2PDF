package main

import (
	"errors"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"golang.org/x/exp/slog"
)

var (
	ignore = flag.String("g", "_", "ExcelでPDF作成対象外とするシート名の先頭文字")
)

var (
	ErrOpenFile   = errors.New("ファイルのオープンに失敗しました。")
	ErrConvertPdf = errors.New("PDFファイルへの変換に失敗しました。")
)

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

 wg := sync.WaitGroup{}
 errChan := make(chan error, len(xlsPaths) +len(docPaths)+len(pptPaths) ) 

 wg.Add(1)
 go func() {
            defer wg.Done()
		if err := convertExcelFileToPdf(xlsPaths, *ignore); err != nil {
                errChan <- err
            }
	})

	wg.Add(1)
 go func() {
            defer wg.Done()
		if err := convertWordFileToPdf(docPaths); err != nil {
                errChan <- err
            }
	})

	wg.Add(1)
 go func() {
            defer wg.Done()
		if err := convertPptFileToPdf(pptPaths); err != nil {
                errChan <- err
            }
	})
 
 wg.Wait()
 close(errChan)

 flag := true
 for err := range errChan {
    if flag {
      slog.Error("PDF変換でエラーが発生しました。")
      flag = false
    }
    slog.Error(err)
 }

	if !flag {
		os.Exit(1)
	}
}

// PowerPointファイルをPDFに変換する。
func convertPptFileToPdf(files []string) (rErr error) {
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

	// // PowerPointウィンドウを表示しないようにする
	// if _, err := oleutil.PutProperty(pptApp, "Visible", false); err != nil {
	// 	return err
	// }

	for _, path := range files {
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
		rErr = convertPptxToPdf(pptApp, fullpath, pdfFullPath)
		if rErr != nil {
			slog.Error(name+" 変換失敗", "err", rErr, "PDFファイル", pdfPath)
			return err
		} else {
			slog.Info(name+" 変換完了", "PDFファイル", pdfPath)
		}
	}

	return nil
}

// PowerPointファイルをPDFに変換する
func convertPptxToPdf(powerpoint *ole.IDispatch, pptPath, pdfFilePath string) error {
	pptname := filepath.Base(pptPath)

	// 　 Dim ppt As New PowerPoint.Application
	// 　 Dim pres As PowerPoint.Presentation
	// 　 Dim save_path As String, file_name As String
	// 　 Dim Target As String
	// 　 Target = Application.GetOpenFilename("PowerPoint,*.pptx")
	// 　 If Target = "False" Then Exit Sub
	// 　 Set pres = ppt.Presentations.Open(Target, WithWindow:=MsoTriState.msoFalse)
	//
	// 　 With pres
	// 　　 save_path = CreateObject("WScript.Shell").SpecialFolders("Desktop")
	// 　　 file_name = "Test"
	// 　　 .ExportAsFixedFormat _
	// 　　　　　 Path:=save_path & "\" & file_name & ".pdf", _
	// 　　　　　 FixedFormatType:=ppFixedFormatTypePDF
	// 　 End With

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

	// pr, err := oleutil.CallMethod(r.ToIDispatch(), "Add", 1, count)
	pr, err := oleutil.CallMethod(r.ToIDispatch(), "Add", sp, count+(sp-1))
	if err != nil {
		return err
	}
	// defer pr.ToIDispatch().Release()

	// PDFに変換する
	// ExportAsFixedFormat (
	// 	Path,
	// 	FixedFormatType, : ppFixedFormatTypePDF(2)
	// 	Intent, : ppFixedFormatIntentPrint(2)
	// 	FrameSlides, : msoFalse(0)
	// 	HandoutOrder, : ppPrintHandoutVerticalFirst(1)
	// 	OutputType, : ppPrintOutputSlides(1)
	// 	PrintHiddenSlides, : msoFalse(0)
	// 	PrintRange,
	// 	RangeType, : ppPrintAll(1)
	// 	SlideShowName, : ""
	// 	IncludeDocProperties, : false
	// 	KeepIRMSettings, : false
	// 	DocStructureTags, : false
	// 	BitmapMissingFonts, : false
	// 	UseISO19005_1, : false
	// 	ExternalExporter : nil
	//)
	// _, err = oleutil.CallMethod(ppt.ToIDispatch(), "ExportAsFixedFormat", pdfFilePath, 2, 2, 0, 1, 1, 0, pr, 1, "", false, false, false, false, false, nil)
	_, err = oleutil.CallMethod(ppt, "ExportAsFixedFormat", pdfFilePath, 2, 2, 0, 1, 1, 0, pr, 1, "", false, false, false, false, false)
	//   ppFixedFormatTypePDF, ppFixedFormatIntentScreen, msoCTrue, ppPrintHandoutHorizontalFirst, ppPrintOutputBuildSlides, msoFalse, , , , False, False, False, False, False
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
	// Open (FileName、 ReadOnly、 Untitled、 WithWindow)
	//  FileName	必須	文字列型 (String)	開くファイルの名前を指定します。
	//  ReadOnly	省略可能	MsoTriState	読み取り/書き込み可能な状態でファイルを開くか、または読み取り専用で開くかを指定します。
	//		msoFalse	既定値です。 読み取り/書き込み可能な状態でファイルを開きます。
	//		msoTrue	読み取り専用でファイルを開きます。
	//  Untitled	省略可能	MsoTriState	ファイルにタイトルを設定するかどうかを指定します。
	//		msoFalse	既定値です。 ファイル名が自動的に、開かれたプレゼンテーションのタイトルとなります。
	//		msoTrue	タイトルなしにファイルを開きます。 これは、ファイルのコピーを作成することと同じです。
	//  WithWindow	省略可能	MsoTriState	ファイルを表示するかどうかを指定します。
	//		msoFalse	開かれたプレゼンテーションを非表示にします。
	//		msoTrue	既定値です。 ファイルを表示可能なウィンドウで開きます。
	ppt, err := oleutil.CallMethod(pres, "Open", path, MsoTriStateMsoTrue, MsoTriStateMsoFalse, MsoTriStateMsoFalse)
	// ppt, err := oleutil.CallMethod(pres, "Open", path)
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
	// pr, err := oleutil.CallMethod(r.ToIDispatch(), "Add", 1, count)
	pr, err := oleutil.CallMethod(r.ToIDispatch(), "Add", sp, count+(sp-1))
	if err != nil {
		return nil, err
	}
	// defer pr.ToIDispatch().Release()

	return pr, nil
}

// WordファイルをPDFに変換する。
func convertWordFileToPdf(files []string) (rErr error) {
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

	for _, path := range files {
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
		rErr = convertDocxToPdf(wordApp, fullpath, pdfFullPath)
		if rErr != nil {
			slog.Error(name+" 変換失敗", "err", rErr, "PDFファイル", pdfPath)
			return err
		} else {
			slog.Info(name+" 変換完了", "PDFファイル", pdfPath)
		}
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
func convertExcelFileToPdf(files []string, ig string) (rErr error) {
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

	for _, path := range files {
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
		rErr = convertXlsxToPdf(excelApp, fullpath, pdfFullPath, ig)
		if rErr != nil {
			slog.Error(name+" 変換完了", err, "PDFファイル", pdfPath)
			return err
		} else {
			slog.Info(name+" 変換完了", "PDFファイル", pdfPath)
		}
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
				// defer selected.ToIDispatch().Release()
			}
		}

		activeSheet, err := oleutil.GetProperty(workbook.ToIDispatch(), "ActiveSheet")
		if err != nil {
			return err
		}
		defer activeSheet.ToIDispatch().Release()

		_, err = oleutil.CallMethod(activeSheet.ToIDispatch(), "ExportAsFixedFormat", 0, pdfFilePath, 0, false, false)
		// _, err = oleutil.CallMethod(workbook.ToIDispatch(), "ExportAsFixedFormat", 0, pdfFilePath, 0, false, false)
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

func convertFileToPdf() filepath.WalkFunc {
	// var eFlag bool
	// var wFlag, pFlag bool
	// var excelApp *ole.IDispatch
	// var wordApp, ppointApp *ole.IDispatch

	return func(path string, info os.FileInfo, err error) (rErr error) {
		if err != nil {
			return err
		}
		if info.IsDir() || strings.HasPrefix(filepath.Base(path), "~") {
			return nil
		}
		fullpath, err := filepath.Abs(path)
		if err != nil {
			return err
		}
		ext := filepath.Ext(fullpath)
		if ext != ".docx" && ext != ".xlsx" && ext != ".xls" && ext != ".pptx" {
			return nil
		}
		// pdfPath := getPathWithoutExt(path) + ".pdf"
		// pdfFullPath, err := filepath.Abs(pdfPath)
		if err != nil {
			return err
		}

		switch ext {
		// case ".docx":
		// 	if !wFlag {
		// 		wFlag = true
		// 		// Wordアプリケーションの作成
		// 		wordApp, err2 = createWordApp()
		// 		if err2 != nil {
		// 			return err2
		// 		}
		// 		defer wordApp.Release()
		// 	}
		// 	err = convertDocxToPdf(wordApp, fullpath, pdfPath)
		case ".xlsx", ".xls":
			// if !eFlag {
			// 	eFlag = true
			// 	// Excelアプリケーションの生成
			// 	excelApp, rErr = createExcelApp()
			// 	if rErr != nil {
			// 		return rErr
			// 	}
			// 	defer excelApp.Release()
			// 	defer func() {
			// 		_, err := oleutil.CallMethod(excelApp, "Quit")
			// 		if err != nil {
			// 			rErr = errors.Join(rErr, err)
			// 		}
			// 		slog.Info("Excel has exited.")
			// 	}()
			// 	slog.Info("Excel launched.")
			// }
			// rErr = convertXlsxToPdf(excelApp, fullpath, pdfFullPath)
			// if rErr != nil {
			// 	slog.Error("Failed to convert.", err, path, pdfPath)
			// } else {
			// 	slog.Info("Converted.", "excel-path", path, "pdf-path", pdfPath)
			// }
		case ".pptx":
			// if !pFlag {
			// 	pFlag = true
			// 	// PowerPointオブジェクトの生成
			// 	ppointApp, err2 = createPowerPointApp()
			// 	if err2 != nil {
			// 		return err2
			// 	}
			// 	defer ppointApp.Release()
			// }
			// err = convertPptxToPdf(ppointApp, fullpath, pdfPath)
		}
		return nil
	}
}

// func convertDocxToPdf(docxPath string, pdfPath string) error {
// doc, err := document.Open(docxPath)
// if err != nil {
// 	return err
// }
// defer doc.Close()
// pdf, err := os.Create(pdfPath)
// if err != nil {
// 	return err
// }
// defer pdf.Close()
// err = doc.Save(pdf, document.SaveOptionPDFPageWidth(8.5), document.SaveOptionPDFPageHeight(11))
// if err != nil {
// 	return err
// }
// 	return nil
// }

// func convertPptxToPdf(pptxPath string, pdfPath string) error {
// 	// prs, err := presentation.Open(pptxPath)
// 	// if err != nil {
// 	// 	return err
// 	// }
// 	// defer prs.Close()
// 	// pdf, err := os.Create(pdfPath)
// 	// if err != nil {
// 	// 	return err
// 	// }
// 	// defer pdf.Close()
// 	// err = prs.SaveToPDF(pdf)
// 	// if err != nil {
// 	// 	return err
// 	// }
// 	return nil
// }

// func convertToPDF(filepath string) error {
// 	// Wordオブジェクトを取得する
// 	word, err := unknown.QueryInterface(ole.IID_IDispatch)
// 	if err != nil {
// 		return err
// 	}
// 	defer word.Release()
// 	// Wordウィンドウを表示しないようにする
// 	_, err = oleutil.PutProperty(word, "Visible", false)
// 	if err != nil {
// 		return err
// 	}
// 	// Wordドキュメントを開く
// 	doc, err := oleutil.CallMethod(word, "Documents", "Open", filepath)
// 	if err != nil {
// 		return err
// 	}
// 	defer doc.Release()
// 	// PDFに変換する
// 	_, err = oleutil.CallMethod(doc, "ExportAsFixedFormat", filepath+".pdf", 17)
// 	if err != nil {
// 		return err
// 	}
// 	return nil
// }

// func convertToPDF(filePath string) error {
// 	// PowerPointオブジェクトをPowerPoint.Applicationオブジェクトにキャスト
// 	powerPoint, err := unknown.QueryInterface(ole.IID_IDispatch)
// 	if err != nil {
// 		return fmt.Errorf("PowerPoint.Applicationオブジェクトの生成に失敗しました: %v", err)
// 	}
// 	defer powerPoint.Release()
// 	// プレゼンテーションの読み込み
// 	presentation, err := oleutil.CallMethod(powerPoint, "Presentations", filePath)
// 	if err != nil {
// 		return fmt.Errorf("プレゼンテーションの読み込みに失敗しました: %v", err)
// 	}
// 	// PDFファイルへの変換
// 	pdfFilePath := filePath + ".pdf"
// 	_, err = oleutil.CallMethod(presentation.ToIDispatch(), "SaveAs", pdfFilePath, 32)
// 	if err != nil {
// 		return fmt.Errorf("PDFへの変換に失敗しました: %v", err)
// 	}
// 	return nil
// }

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
		if !info.IsDir() || strings.HasPrefix(filepath.Base(path), "~") {
			switch filepath.Ext(info.Name()) {
			case ".xlsx", ".xls":
				xslPaths = append(xslPaths, path)
			case ".docx", "doc":
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
	slog.Info("usage: PDFConverterGO [flags] path")
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
