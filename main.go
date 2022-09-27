package main

import (
	"context"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"os"
	"time"

	"cloud.google.com/go/storage"
	firebase "firebase.google.com/go"
	"github.com/ConvertAPI/convertapi-go"
	"github.com/ConvertAPI/convertapi-go/config"
	"github.com/joho/godotenv"
	"github.com/xuri/excelize/v2"
	"google.golang.org/api/option"
)

func main() {
	f, err := excelize.OpenFile("./files/template.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	// バイナリファイル読み込み用
	// b, err := ioutil.ReadFile("./files/jobHistory.xlsx")
	// if err != nil {
	// 	fmt.Println(err)
	// 	return
	// }
	// f, err := excelize.OpenReader(bytes.NewReader(b))
	// if err != nil {
	// 	fmt.Println(err)
	// 	return
	// }

	// 1番目シート名を取得しています
	templateSheetName := f.GetSheetName(0)
	fmt.Println(templateSheetName)

	// シート追加
	//固有の名前つける
	addSheetName := time.Now().Format("20060102150405") + "求人ID"
	i := f.NewSheet(addSheetName)

	// シート削除
	// defer f.DeleteSheet(addSheetName)

	// シートコピー
	from := f.GetSheetIndex(templateSheetName)
	to := f.GetSheetIndex(addSheetName)
	if err := f.CopySheet(from, to); err != nil {
		fmt.Println(err)
		return
	}

	phoneNumber := "090-1234-5678"

	// 以下でセル毎に取得情報を入れていく
	f.SetCellValue(addSheetName, "A1", phoneNumber)

	// シート名取得　=> サンプル2 と出力される
	fmt.Println("シート名は:", f.GetSheetName(i))

	if err := f.SaveAs("sample2.xlsx"); err != nil {
		fmt.Println(err)
		return
	}

	err = godotenv.Load(".env")
	if err != nil {
		fmt.Println("Error loading .env file")
	}

	config.Default.Secret = os.Getenv("CONVERT_API_SECRET")

	if file, errs := convertapi.ConvertPath("./files/template.xlsx", "./files/result.pdf"); errs == nil {
		fmt.Println("PDF file saved to: ", file.Name())
	} else {
		fmt.Println(errs)
	}

	// another way
	// pdfRes := convertapi.ConvDef("xlsx", "pdf", param.NewPath("file", "test.xlsx", nil))

	// pdfRes.ToPath("/tmp/result.pdf")

	defer os.Remove("./files/result.pdf")
	config := &firebase.Config{
		StorageBucket: "storage-exp.appspot.com",
	}
	opt := option.WithCredentialsFile("storage-exp-key.json")
	app, err := firebase.NewApp(context.Background(), config, opt)
	if err != nil {
		log.Fatalln(err)
	}

	client, err := app.Storage(context.Background())
	if err != nil {
		log.Fatalln(err)
	}

	bucket, err := client.DefaultBucket()
	if err != nil {
		log.Fatalln(err)
	}

	localFilename := "./test.pdf" // ローカルのファイル名
	remoteFilename := "test.pdf"  // Bucketに保存されるファイル名
	contentType := "text/plain"
	ctx := context.Background()

	writer := bucket.Object(remoteFilename).NewWriter(ctx)
	writer.ObjectAttrs.ContentType = contentType
	writer.ObjectAttrs.CacheControl = "no-cache"
	writer.ObjectAttrs.ACL = []storage.ACLRule{
		{
			Entity: storage.AllUsers,
			Role:   storage.RoleReader,
		},
	}

	fileForUpload, err := os.Open(localFilename)
	if _, err = io.Copy(writer, fileForUpload); err != nil {
		log.Fatalln(err)
	}
	defer f.Close()

	if err := writer.Close(); err != nil {
		fmt.Println(err)
	}

	rc, err := bucket.Object(remoteFilename).NewReader(ctx)
	if err != nil {
		log.Fatalln(err)
	}
	defer rc.Close()

	data, err := ioutil.ReadAll(rc)
	if err != nil {
		log.Fatalln(err)
	}

	log.Printf("Downloaded contents: %v\n", string(data))

}
