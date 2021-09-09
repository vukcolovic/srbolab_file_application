package model

import (
	"encoding/json"
	"io/ioutil"
	"log"
	"os"
)

type UverenjeRow struct {
	RedniBroj             string
	BrojZapisnika         string
	DatumZapisnika        string
	BrojUgovora           string
	PodnosilacZahteva     string
	BrojUverenja          string
	DatumUverenja         string
	BrojZahteva           string
	DatumZahteva          string
	MarkaIOznaka          string
	IdentifikacionaOznaka string
	BrojIzvestaja         string
	DatumIzvestaja        string
	BrojPrijaveUABS       string
	IspitnoMesto          string
	Vlasnik               string
	Prebivaliste          string
	Usaglaseno            string
	VrstaIzvestaja        string
	Napomena              string
	Kontrolor1            string
	Kontrolor2            string
}


type Config struct {
	FirstNumToProcess int	`json:"prvi_broj_za_procesuiranje"`
	LastNumToProcess 	int `json:"poslednji_broj_za_procesuiranje"`
}

func ProcessConfig() Config {
	configFile, err := os.Open("config.json")
	if err != nil {
		log.Println("Warning, can't open config.json file!!!")
		return Config{0, 0}
	}

	defer configFile.Close()

	byteValue, _ := ioutil.ReadAll(configFile)

	var config Config
	json.Unmarshal(byteValue, &config)

	if config.FirstNumToProcess < 0 || config.LastNumToProcess <0 {
		log.Println("Prvi broj za procesuiranje ni poslednji broj za procesuiranje ne mogu biti manji od 0!!!")
		return Config{0, 0}
	}
	if config.FirstNumToProcess > config.LastNumToProcess {
		log.Println("Prvi broj za procesuiranje ne moze biti veci od poslednjeg broja za procesuiranje!!!")
		return Config{0, 0}
	}

	return config
}
