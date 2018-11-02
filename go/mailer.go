package main

import (
  "os"
  "log"
  "fmt"
  "github.com/kylelemons/go-gypsy/yaml"
)

func main() {
	// If the file doesn't exist, create it, or append to the file
	logfile, _ := os.OpenFile("app.log", os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
	defer logfile.Close()
	logger := log.New(logfile, "example ", log.LstdFlags|log.Lshortfile|log.Lmicroseconds)
	logger.Println("init")
	config, err := yaml.ReadFile("mailer.yaml")
        if err != nil {
		fmt.Println(err)
		logger.Fatalln(err)
        }
	fmt.Println(config.Get("SmtpServer"))
	fmt.Println(config.Get("SmtpServerPort"))
	fmt.Println(config.Get("SendUsername"))
	fmt.Println(config.Get("SendPassword"))
}
