package main

import (
	"fmt"
	"github.com/jmoiron/sqlx"
	_ "github.com/mattn/go-sqlite3"
	"github.com/nguyenthenguyen/docx"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"strconv"
	"strings"
	"wolfPowerSrbolabApp/model"
)

const (
	DriverName = "sqlite3"
	Connection = "file:data.db3"
)

func main() {
	uverenjaRows, err := getRowsFromExcelFile()
	if err != nil {
		log.Fatal(err)
	}

	db, err := sqlx.Connect(DriverName, Connection)
	if err != nil {
		log.Fatal(err)
	}

	defer db.Close()

	configStruct := model.ProcessConfig()

	var lastProcessed int
	rows, err := db.Query(`SELECT last FROM last`)
	if err != nil {
		log.Fatal(err)
	}
	for rows.Next() {
		err = rows.Scan(&lastProcessed)
		if err != nil {
			log.Fatal(err)
		}
	}
	rows.Close()

	newLastProcessed := lastProcessed
	for _, uverenje := range uverenjaRows {
		redniBroj, err := strconv.Atoi(strings.TrimSpace(uverenje.RedniBroj))
		if err != nil {
			log.Fatal(fmt.Sprintf("Greska u parsiranju rednog broja, greska: %s", err))
		}
		if redniBroj <= lastProcessed && (redniBroj > configStruct.LastNumToProcess || redniBroj < configStruct.FirstNumToProcess){
			continue
		}
		if redniBroj > lastProcessed {
			newLastProcessed = redniBroj
		}

		fillFiles(uverenje)
	}

	_, err = db.Exec("DELETE FROM last")
	if err != nil {
		log.Fatal(err)
	}
	_, err = db.Exec("INSERT INTO last VALUES ($1)", newLastProcessed)
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("Uspesno odradjeno :-)")
}

func fillFiles(uverenje model.UverenjeRow) {
	var tipIzvestaja string
	var tipNaloga string

	switch strings.ToLower(strings.TrimSpace(uverenje.VrstaIzvestaja)) {
	case "upotrebljavana vozila":
		tipIzvestaja = "izvestaj_upotrebljavana"
		tipNaloga = "nalog_upotrebljavana"
	case "vozila na tng":
		tipIzvestaja = "izvestaj_tng"
		tipNaloga = "nalog_tng"
	case "prepravljena vozila":
		tipIzvestaja = "izvestaj_prepravljana"
		tipNaloga = "nalog_prepravljana"
	case "stakla na vozilima":
		tipIzvestaja = "izvestaj_stakla"
		tipNaloga = "nalog_stakla"
	case "vozila na kpg":
		tipIzvestaja = "izvestaj_kpg"
		tipNaloga = "nalog_kpg"
	default:
		panic(fmt.Sprintf("Ne moze da se odredi tip izvestaja!!! Izvestaj: %s, redni broj: %s", uverenje.VrstaIzvestaja, uverenje.RedniBroj))
	}

	creteIzvestaj(tipIzvestaja, uverenje)

	createNalog(tipNaloga, uverenje)
}

func creteIzvestaj(tipIzvestaja string, uverenje model.UverenjeRow) {
	// Read from docx file
	r, err := docx.ReadDocxFile(fmt.Sprintf("templates/%s.docx", tipIzvestaja))
	if err != nil {
		log.Fatal(err)
	}
	docx1 := r.Editable()

	docx1.Replace("{broj_izvestaja}", uverenje.BrojIzvestaja, -1)
	docx1.Replace("{datum_izvestaja}", uverenje.DatumIzvestaja, -1)
	docx1.Replace("{broj_izvestaja2}", uverenje.BrojIzvestaja, -1)
	docx1.Replace("{datum_izvestaja2}", uverenje.DatumIzvestaja, -1)
	docx1.Replace("{broj_zahteva}", uverenje.BrojZahteva, -1)
	docx1.Replace("{datum_zahteva}}", uverenje.DatumZahteva , -1)
	docx1.Replace("{broj_zapisnika}", uverenje.BrojZapisnika, -1)
	docx1.Replace("{datum_zapisnika}", uverenje.DatumZapisnika , -1)
	docx1.Replace("{vlasnik}", uverenje.Vlasnik, -1)
	docx1.Replace("{prebivaliste}", uverenje.Prebivaliste, -1)
	docx1.Replace("{marka_i_oznaka}", uverenje.MarkaIOznaka, -1)
	docx1.Replace("{vin}", uverenje.IdentifikacionaOznaka, -1)

	var usaglaseno string
	if uverenje.Usaglaseno == "N" {
		usaglaseno = "NIJE USAGLASENO!!!!!!!"
		//docx1.Replace("<w:checkBox><w:sizeAuto></w:sizeAuto></w:checkBox>", "<w:checkBox><w:sizeAuto></w:sizeAuto><w:checked/></w:checkBox>", -1)
		//docx1.Replace("<w:checkBox><w:sizeAuto></w:sizeAuto><w:checked/></w:checkBox>", "<w:checkBox><w:sizeAuto></w:sizeAuto></w:checkBox>", -1)
	}

	if _, err = os.Stat(fmt.Sprintf("izvestaji/%s", uverenje.DatumIzvestaja)); os.IsNotExist(err) {
		err = os.Mkdir(fmt.Sprintf("izvestaji/%s", uverenje.DatumIzvestaja), 0777)
		if err != nil {
			log.Fatal(err)
		}
	}

	err = docx1.WriteToFile(fmt.Sprintf("izvestaji/%s/%s%s%s.docx", uverenje.DatumIzvestaja, tipIzvestaja, uverenje.BrojUverenja, usaglaseno))
	if err != nil {
		log.Fatal(err)
	}
	r.Close()
}

func createNalog(tipNaloga string, uverenje model.UverenjeRow) {
	// Read from docx file
	r, err := docx.ReadDocxFile(fmt.Sprintf("templates/%s.docx", tipNaloga))
	if err != nil {
		log.Fatal(err)
	}
	docx1 := r.Editable()

	docx1.Replace("{broj_zahteva}", uverenje.BrojZahteva , -1)
	docx1.Replace("{datumzahteva2}", uverenje.DatumZahteva , -1)
	docx1.Replace("{datumzahteva3}", uverenje.DatumZahteva , -1)
	docx1.Replace("{vlasnik}", uverenje.Vlasnik, -1)
	docx1.Replace("{prebivaliste}", uverenje.Prebivaliste, -1)
	docx1.Replace("{kontrolor1}", uverenje.Kontrolor1, -1)
	docx1.Replace("{kontrolor2}", uverenje.Kontrolor2, -1)
	docx1.ReplaceHeader("{broj_zahteva}", uverenje.BrojZahteva)
	docx1.ReplaceHeader("{datumzahteva}", uverenje.DatumZahteva)

	if _, err = os.Stat(fmt.Sprintf("nalozi/%s", uverenje.DatumIzvestaja)); os.IsNotExist(err) {
		err = os.Mkdir(fmt.Sprintf("nalozi/%s", uverenje.DatumIzvestaja), 0777)
		if err != nil {
			log.Fatal(err)
		}
	}

	err = docx1.WriteToFile(fmt.Sprintf("nalozi/%s/%s%s.docx", uverenje.DatumIzvestaja, tipNaloga, uverenje.BrojUverenja))
	if err != nil {
		log.Fatal(err)
	}
	r.Close()
}

func getRowsFromExcelFile() ([]model.UverenjeRow, error) {
	fileExcel, err := excelize.OpenFile("documents/Evidencija izvestaja.xlsx")
	if err != nil {
		fmt.Println(err)
		return nil, err
	}

	rows, err := fileExcel.GetRows("Evidencija izveštaja")
	if err != nil {
		fmt.Println(err)
	}

	uverenjeRows := []model.UverenjeRow{}
	for i, row := range rows {
		if i == 0 {
			continue
		}
		e := model.UverenjeRow{}

		if len(row) > 0 {
			e.RedniBroj = row[0]
		}
		if len(row) > 1 {
			e.BrojZahteva = row[1]
		} else {
			continue
		}
		if len(row) > 2 {
			e.DatumZahteva = row[2]
		}
		if len(row) > 3 {
			e.BrojUgovora = row[3]
		}
		if len(row) > 4 {
			e.PodnosilacZahteva = row[4]
		}
		if len(row) > 5 {
			e.Vlasnik = row[5]
		}
		if len(row) > 6 {
			e.Prebivaliste = row[6]
		}
		if len(row) > 7 {
			e.VrstaIzvestaja = row[7]
		}
		if len(row) > 8 {
			e.MarkaIOznaka = row[8]
		}
		if len(row) > 9 {
			e.IdentifikacionaOznaka = row[9]
		}
		if len(row) > 10 {
			e.BrojZapisnika = row[10]
		}
		if len(row) > 11 {
			e.DatumZapisnika = row[11]
		}
		if len(row) > 12 {
			e.BrojIzvestaja = row[12]
		}
		if len(row) > 13 {
			e.DatumIzvestaja = row[13]
		}
		if len(row) > 14 {
			e.BrojPrijaveUABS = row[14]
		}
		if len(row) > 15 {
			e.BrojUverenja = row[15]
		}
		if len(row) > 16 {
			e.Usaglaseno = row[16]
		}
		if len(row) > 17 {
			e.Napomena = row[17]
		}
		if len(row) > 18 {
			e.Kontrolor1 = row[18]
		}
		if len(row) > 19 {
			e.Kontrolor2 = row[19]
		}

		uverenjeRows = append(uverenjeRows, e)
	}

	return uverenjeRows, nil
}
