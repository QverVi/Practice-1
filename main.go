package main

import (
    "fmt"
    "log"
    "os"
    "strings"
    "net/http"
    "io"
    "github.com/tealeg/xlsx"
    tgbotapi "github.com/go-telegram-bot-api/telegram-bot-api/v5"
    godotenv "github.com/joho/godotenv"
)

var bot *tgbotapi.BotAPI
var newMessage = tgbotapi.NewMessage
var startMenu = tgbotapi.NewInlineKeyboardMarkup(
    tgbotapi.NewInlineKeyboardRow(tgbotapi.NewInlineKeyboardButtonData("загрузка xls/xlsx файла", "isXml_")),
)

// base functions
func sendMessage(msg tgbotapi.Chattable) {
    if _, err := bot.Send(msg); err != nil {
        log.Panicf("Send Message error %v", err)
    }
}
func send(msg tgbotapi.MessageConfig) { sendMessage(msg) }

// функция получения файла из telegram по file_id
func downloadFileFromTelegram(fileID string) (string, error) {
    // Получаем информацию о файле
    fileConfig := tgbotapi.FileConfig{FileID: fileID}
    file, err := bot.GetFile(fileConfig)
    if err != nil {
        return "", err
    }
    fileURL := file.Link(os.Getenv("token_telegram_bot"))

    // Создаем временный файл
    tempFile, err := os.CreateTemp("", "schedule_*.xlsx")
    if err != nil {
        return "", err
    }
    tempFileName := tempFile.Name()
    tempFile.Close() // закрываем, чтобы потом перезаписать

    // скачиваем файл по URL
    err = downloadFile(fileURL, tempFileName)
    if err != nil {
        os.Remove(tempFileName)
        return "", err
    }
    return tempFileName, nil
}

// функция скачивания файла по URL
func downloadFile(url, filepath string) error {
    resp, err := http.Get(url)
    if err != nil {
        return err
    }
    defer resp.Body.Close()

    out, err := os.Create(filepath)
    if err != nil {
        return err
    }
    defer out.Close()

    _, err = io.Copy(out, resp.Body)
    return err
}

// main
func main() {
    err := godotenv.Load(".env")
    if err != nil {
        fmt.Println(err)
        log.Fatal(".env not found")
    }
    bot, err = tgbotapi.NewBotAPI(os.Getenv("token_telegram_bot"))
    if err != nil {
        log.Fatalf("Failed to init api: %v", err)
    }

    updateConf := tgbotapi.NewUpdate(0)
    updateConf.Timeout = 30
    updates := bot.GetUpdatesChan(updateConf)

    for update := range updates {
        if update.CallbackQuery != nil {
            callbacks(update)
        } else if update.Message != nil && update.Message.IsCommand() {
            commands(update)
        } else if update.Message != nil && update.Message.Document != nil {
            handleFile(update.Message)
        }
    }
}

func callbacks(update tgbotapi.Update) {
    data := update.CallbackQuery.Data
    chatId := update.CallbackQuery.From.ID
    txt := "fef"
    switch data {
    case "isXml_":
        msg := newMessage(chatId, txt)
        send(msg)
    }
}

func commands(update tgbotapi.Update) {
    command := update.Message.Command()
    commandChatId := update.Message.Chat.ID
    switch command {
    case "help":
        msg := newMessage(commandChatId, "/start\n/help\n/wait get xls/xlsx\n/stop wait")
        send(msg)
    case "start":
        msg := newMessage(commandChatId, "file !")
        msg.ReplyMarkup = startMenu
        msg.ParseMode = "MarkDown"
        send(msg)
    }
}

// обработка файла
func handleFile(msg *tgbotapi.Message) {
    fileID := msg.Document.FileID
    // скачиваем файл из Telegram по file_id
    filePath, err := downloadFileFromTelegram(fileID)
    if err != nil {
        log.Printf("Error downloading file from telegram: %v", err)
        return
    }
    defer os.Remove(filePath) // удаляем временный файл после обработки

    // парсинг файла
    discCounts, err := parseExcel(filePath)
    if err != nil {
        log.Printf("Error parsing excel: %v", err)
        return
    }

    // формируем отчет
    report := ""
    for disc, count := range discCounts {
        report += fmt.Sprintf("%s: %d пар\n", disc, count)
    }

    // отправляем отчет
    chatID := msg.Chat.ID
    responseMsg := newMessage(chatID, report)
    send(responseMsg)
}

// парсинг Excel файла
func parseExcel(filepath string) (map[string]int, error) {
    file, err := xlsx.OpenFile(filepath)
    if err != nil {
        return nil, err
    }

    discCounts := make(map[string]int)
    if len(file.Sheets) == 0 {
        return nil, fmt.Errorf("no sheets found")
    }
    sheet := file.Sheets[0]

    for _, row := range sheet.Rows {
        if len(row.Cells) < 2 {
            continue
        }
        disc := row.Cells[1].String() // предполагается, что название дисциплины во 2-й колонке
        if strings.TrimSpace(disc) != "" {
            discCounts[disc]++
        }
    }

    return discCounts, nil
}