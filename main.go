package main

import (
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"regexp"
	"strconv"
	"strings"

	tgbotapi "github.com/go-telegram-bot-api/telegram-bot-api/v5"
	godotenv "github.com/joho/godotenv"
	excelize "github.com/xuri/excelize/v2"
)

// Глобальные переменные
var bot *tgbotapi.BotAPI
var textError string
var strbuild strings.Builder

// Карта для хранения режима обработки по chatID
var userMode = make(map[int64]string)

// функции для оптимизации
func errors(err error, textError string) {
	if err != nil {
		fmt.Println(err)
		log.Fatal(textError)
	}
}
func main() {
	// Загружаем переменные окружения
	err := godotenv.Load(".env")
	textError = ".env не найден"
	errors(err, textError)

	bot, err = tgbotapi.NewBotAPI(os.Getenv("token_telegram_bot"))
	textError = "Не удалось инициализировать api"
	errors(err, textError)

	// Настройка и получение обновлений
	updateConf := tgbotapi.NewUpdate(0)
	updateConf.Timeout = 30
	updates := bot.GetUpdatesChan(updateConf)

	// Обработка обновлений
	for update := range updates {
		strbuild.Reset()
		if update.Message != nil {
			if update.Message.IsCommand() {
				handleCommand(bot, update.Message)
			} else if update.Message.Document != nil {
				handleDocument(bot, update.Message)
			} else {
				bot.Send(tgbotapi.NewMessage(update.Message.Chat.ID, "Пожалуйста, отправьте Excel файл или используйте /start для выбора режима."))
			}
		} else if update.CallbackQuery != nil {
			handleCallback(bot, update.CallbackQuery)
		}
	}
}

func handleCommand(bot *tgbotapi.BotAPI, msg *tgbotapi.Message) {
	switch msg.Command() {
	case "start":
		bot.Send(tgbotapi.NewMessage(msg.Chat.ID, "Здравстуйте, это бот по обработке отчетов\nвоспользуйтесь /help для того чтобы узнать больше"))
	case "help":
		bot.Send(tgbotapi.NewMessage(msg.Chat.ID, "Отправьте XLSX/XLS файл, и я подготовлю нужный отчет.\nИспользуйте /setmode, чтобы выбрать режим обработки"))
	case "setmode":
		sendModeSelection(bot, msg.Chat.ID)
	default:
		bot.Send(tgbotapi.NewMessage(msg.Chat.ID, "Неизвестная команда. Используйте /start или /help"))
	}
}

func sendModeSelection(bot *tgbotapi.BotAPI, chatID int64) {
	msg := tgbotapi.NewMessage(chatID, "Выберите режим обработки:")
	keyboard := tgbotapi.InlineKeyboardMarkup{
		InlineKeyboard: [][]tgbotapi.InlineKeyboardButton{
			{
				tgbotapi.NewInlineKeyboardButtonData("Расписание групп", "mode_schedule"),
				tgbotapi.NewInlineKeyboardButtonData("Темы уроков", "mode_lessons"),
			},
			{
				tgbotapi.NewInlineKeyboardButtonData("Студенты", "mode_students"),
				tgbotapi.NewInlineKeyboardButtonData("Посещаемость", "mode_attendance"),
			},
			{
				tgbotapi.NewInlineKeyboardButtonData("Проверенные ДЗ", "mode_checked_homework"),
				tgbotapi.NewInlineKeyboardButtonData("Сданные ДЗ", "mode_submitted_homework"),
			},
		},
	}
	msg.ReplyMarkup = keyboard
	bot.Send(msg)
}

func handleCallback(bot *tgbotapi.BotAPI, callback *tgbotapi.CallbackQuery) {
	chatID := callback.Message.Chat.ID
	data := callback.Data

	switch data {
	case "mode_schedule":
		userMode[chatID] = "schedule"
	case "mode_lessons":
		userMode[chatID] = "lessons"
	case "mode_students":
		userMode[chatID] = "students"
	case "mode_attendance":
		userMode[chatID] = "attendance"
	case "mode_checked_homework":
		userMode[chatID] = "checked_homework"
	case "mode_submitted_homework":
		userMode[chatID] = "submitted_homework"
	}

	text := "Режим выбран: " + strings.Title(strings.ReplaceAll(data[5:], "_", " "))
	bot.Request(tgbotapi.NewCallback(callback.ID, text))
	bot.Send(tgbotapi.NewMessage(chatID, "Режим обработки установлен. Теперь отправьте файл для обработки."))
}

func handleDocument(bot *tgbotapi.BotAPI, msg *tgbotapi.Message) {
	chatID := msg.Chat.ID
	filename := msg.Document.FileName

	if !(strings.HasSuffix(filename, ".xlsx") || strings.HasSuffix(filename, ".xls")) {
		bot.Send(tgbotapi.NewMessage(chatID, "Пожалуйста, отправьте файл в формате Excel (.xlsx или .xls)"))
		return
	}

	sentMsg, _ := bot.Send(tgbotapi.NewMessage(chatID, "⏳ Обрабатываю файл..."))

	file, err := bot.GetFile(tgbotapi.FileConfig{FileID: msg.Document.FileID})
	if err != nil {
		bot.Send(tgbotapi.NewMessage(chatID, "Ошибка при получении файла"))
		return
	}
	url := file.Link(bot.Token)

	localPath := fmt.Sprintf("temp_%d_%s", msg.MessageID, filename)
	defer os.Remove(localPath)
	if err := downloadFile(url, localPath); err != nil {
		bot.Send(tgbotapi.NewMessage(chatID, "Ошибка при скачивании файла"))
		return
	}

	// Попытка определить тип файла автоматически
	category := determineFileType(localPath)

	// Проверка режима
	mode, boolMode := userMode[chatID]
	var res string
	var errProcess error

	if boolMode {
		switch mode {
		case "schedule", "lessons", "students", "attendance", "checked_homework", "submitted_homework":
		default:
			bot.Send(tgbotapi.NewMessage(chatID, "Некорректный режим обработки. Используйте /start для выбора режима."))
			return
		}
	} else {
		if category == "" {
			bot.Send(tgbotapi.NewMessage(chatID, "Не удалось определить тип файла. Пожалуйста, убедитесь, что выбран правильный файл."))
			return
		}
	}

	// Обработка по режиму или по определенному типу файла
	if boolMode {
		switch mode {
		case "schedule":
			res, errProcess = processSchedule(localPath)
		case "lessons":
			res, errProcess = processLessonTopics(localPath)
		case "students":
			res, errProcess = processStudents(localPath)
		case "attendance":
			res, errProcess = processAttendance(localPath)
		case "checked_homework":
			res, errProcess = processCheckedHomework(localPath)
		case "submitted_homework":
			res, errProcess = processSubmittedHomework(localPath)
		}
	} else {
		switch category {
		case "Расписание групп":
			res, errProcess = processSchedule(localPath)
		case "Темы уроков":
			res, errProcess = processLessonTopics(localPath)
		case "Отчет по студентам":
			res, errProcess = processStudents(localPath)
		case "Посещаемость по преподавателям":
			res, errProcess = processAttendance(localPath)
		case "Отчет по проверенным ДЗ":
			res, errProcess = processCheckedHomework(localPath)
		case "Отчет по сданным ДЗ":
			res, errProcess = processSubmittedHomework(localPath)
		default:
			bot.Send(tgbotapi.NewMessage(chatID, "Обработка этого типа файла не реализована или не распознана."))
			return
		}
	}

	if errProcess != nil {
		bot.Send(tgbotapi.NewMessage(chatID, fmt.Sprintf("Ошибка при обработке файла: %v", errProcess)))
		return
	}

	parts := splitMessage(res, 4000)
	bot.Send(tgbotapi.NewDeleteMessage(chatID, sentMsg.MessageID))
	for _, part := range parts {
		bot.Send(tgbotapi.NewMessage(chatID, part))
	}
}

func downloadFile(url, path string) error {
	responce, err := http.Get(url)
	if err != nil {
		return err
	}
	defer responce.Body.Close()
	if responce.StatusCode != http.StatusOK {
		return fmt.Errorf("HTTP статус %d", responce.StatusCode)
	}
	out, err := os.Create(path)
	if err != nil {
		return err
	}
	defer out.Close()
	_, err = io.Copy(out, responce.Body)
	return err
}

// Функция определения типа файла по содержимому
func determineFileType(filepath string) string {
	file, err := excelize.OpenFile(filepath)
	if err != nil {
		return ""
	}
	defer file.Close()
	sheets := file.GetSheetList()
	if len(sheets) == 0 {
		return ""
	}
	rows, err := file.GetRows(sheets[0])
	if err != nil || len(rows) == 0 {
		return ""
	}
	header := rows[0]
	txt := strings.ToLower(strings.Join(header, " "))

	if strings.Contains(txt, "группа") && strings.Contains(txt, "время") && strings.Contains(txt, "пара") {
		return "Расписание групп"
	}
	if strings.Contains(txt, "урок") || strings.Contains(txt, "тема") || strings.Contains(txt, "тема урока") {
		return "Темы уроков"
	}
	if strings.Contains(txt, "fio") || (strings.Contains(txt, "homework") && strings.Contains(txt, "classroom")) {
		return "Отчет по студентам"
	}
	if strings.Contains(txt, "фио преподавателя") && strings.Contains(txt, "средняя посещаемость") {
		return "Посещаемость по преподавателям"
	}
	if strings.Contains(txt, "форма обучения") && strings.Contains(txt, "фио преподавателя") ||
		(strings.Contains(txt, "месяц") || strings.Contains(txt, "неделя")) || strings.Contains(txt, "день") || strings.Contains(txt, "проверено") {
		return "Отчет по проверенным ДЗ"
	}
	if strings.Contains(txt, "fio") && (strings.Contains(txt, "percentage homework") || strings.Contains(txt, "домашнее")) {
		return "Отчет по сданным ДЗ"
	}
	return ""
}

// 1. Расписание групп
func processSchedule(filepath string) (string, error) {
	file, err := excelize.OpenFile(filepath)
	if err != nil {
		return "", err
	}
	defer file.Close()

	rows, err := file.GetRows(file.GetSheetName(0))
	if err != nil || len(rows) < 2 {
		return "Нет данных в файле", nil
	}
	header := rows[0]
	groupIndx, pairIndx, timeIndx := -1, -1, -1

	for i, col := range header {
		colLower := strings.ToLower(col)
		if strings.Contains(colLower, "группа") {
			groupIndx = i
		} else if strings.Contains(colLower, "пара") {
			if pairIndx == -1 {
				pairIndx = i
			}
		} else if strings.Contains(colLower, "время") {
			if timeIndx == -1 {
				timeIndx = i
			}
		}
	}

	if groupIndx == -1 || pairIndx == -1 || timeIndx == -1 {
		return "Не удалось найти колонки 'Группа' или 'Пара'", nil
	}

	groupStats := make(map[string]map[string]int)

	for _, row := range rows[1:] {
		if len(row) <= max(groupIndx, pairIndx) {
			continue
		}
		group := strings.TrimSpace(row[groupIndx])
		subject := strings.TrimSpace(row[pairIndx])
		if group == "" || subject == "" {
			continue
		}
		if _, ok := groupStats[group]; !ok {
			groupStats[group] = make(map[string]int)
		}
		groupStats[group][subject]++
	}

	strbuild.WriteString("📅 ОТЧЕТ ПО РАСПИСАНИЮ ГРУПП\n")
	strbuild.WriteString("Количество пар по дисциплинам:\n\n")
	for group, subjects := range groupStats {
		strbuild.WriteString(fmt.Sprintf("Группа: %s\n", group))
		for subj, count := range subjects {
			strbuild.WriteString(fmt.Sprintf("  %s: %d пар\n", subj, count))
		}
		strbuild.WriteString("\n")
	}
	return strbuild.String(), nil
}

// 2. Темы уроков
func processLessonTopics(filepath string) (string, error) {
	file, err := excelize.OpenFile(filepath)
	if err != nil {
		return "", err
	}
	defer file.Close()

	rows, err := file.GetRows(file.GetSheetName(0))
	if err != nil || len(rows) == 0 {
		return "Нет данных в файле", nil
	}

	topicCol := -1
	for i, col := range rows[0] {
		if strings.Contains(strings.ToLower(col), "тема урока") {
			topicCol = i
			break
		}
	}
	if topicCol == -1 {
		return "Не найдена колонка с темами уроков", nil
	}

	validTopics := []string{}
	invalidTopics := []string{}
	pattern := regexp.MustCompile(`^Урок №\s*\d+.*Тема:`)
	for _, row := range rows[1:] {
		if len(row) <= topicCol {
			continue
		}
		topic := strings.TrimSpace(row[topicCol])
		if topic == "" {
			continue
		}
		if pattern.MatchString(topic) {
			validTopics = append(validTopics, topic)

		} else {
			invalidTopics = append(invalidTopics, topic)

		}
	}

	strbuild.WriteString("📚 ОТЧЕТ ПО ТЕМАМ ЗАНЯТИЙ\n\n")
	if len(validTopics) > 0 {
		strbuild.WriteString("✅ Темы в правильном формате:\n")
		for _, t := range validTopics {
			strbuild.WriteString(fmt.Sprintf("• %s\n", t))
		}
		strbuild.WriteString("\n")
	}
	if len(invalidTopics) > 0 {
		strbuild.WriteString("❌ Темы в НЕправильном формате:\n")
		for _, t := range invalidTopics {
			strbuild.WriteString(fmt.Sprintf("• %s\n", t))
		}
	} else if len(validTopics) == 0 {
		strbuild.WriteString("Темы уроков не найдены")
	}
	return strbuild.String(), nil
}

// 3. Студенты со слабым оцениванием
func processStudents(filepath string) (string, error) {
	file, err := excelize.OpenFile(filepath)
	if err != nil {
		return "", err
	}
	defer file.Close()

	rows, err := file.GetRows(file.GetSheetName(0))
	if err != nil || len(rows) < 2 {
		return "Нет данных в файле", nil
	}
	header := rows[0]
	fioIndx, homeworkIndx, classworkIndx := -1, -1, -1
	for i, col := range header {
		switch strings.ToLower(col) {
		case "фио", "fio":
			fioIndx = i
		case "homework", "домашняя работа":
			homeworkIndx = i
		case "classwork", "классная работа":
			classworkIndx = i
		}
	}
	if fioIndx == -1 {
		return "Не найдена колонка с ФИО студентов", nil
	}
	var problemStudents []string
	for _, row := range rows[1:] {
		if len(row) <= max(fioIndx, homeworkIndx, classworkIndx) {
			continue
		}
		name := strings.TrimSpace(row[fioIndx])
		if name == "" {
			continue
		}
		if homeworkIndx != -1 && len(row) > homeworkIndx {
			if row[homeworkIndx] == "1" {
				problemStudents = append(problemStudents, fmt.Sprintf("%s (домашняя: 1)", name))
				continue
			}
		}
		if classworkIndx != -1 && len(row) > classworkIndx {
			gradeStr := strings.TrimSpace(row[classworkIndx])
			if grade, err := strconv.ParseFloat(gradeStr, 64); err == nil && grade < 3 {
				problemStudents = append(problemStudents, fmt.Sprintf("%s (классная: %.1f)", name, grade))
			}
		}
	}
	strbuild.WriteString("👨‍🎓 ОТЧЕТ ПО СТУДЕНТАМ\n\n")
	if len(problemStudents) > 0 {
		strbuild.WriteString("Студенты, требующие внимания:\n")
		for i, s := range problemStudents {
			strbuild.WriteString(fmt.Sprintf("%d. %s\n", i+1, s))
		}
	} else {
		strbuild.WriteString("✅ Все студенты успешно справляются")
	}
	return strbuild.String(), nil
}

// 4. Посещаемость преподавателей ниже 40%
func processAttendance(filepath string) (string, error) {
	file, err := excelize.OpenFile(filepath)
	if err != nil {
		return "", err
	}
	defer file.Close()

	rows, err := file.GetRows(file.GetSheetName(0))
	if err != nil || len(rows) < 2 {
		return "Нет данных в файле", nil
	}
	header := rows[0]
	teacherIndx, attendanceIndx := -1, -1
	for i, col := range header {
		switch strings.ToLower(col) {
		case "фио преподавателя":
			teacherIndx = i
		case "средняя посещаемость":
			attendanceIndx = i
		}
	}
	if teacherIndx == -1 || attendanceIndx == -1 {
		return "Не найдены необходимые колонки", nil
	}
	var lowAttendanceTeachers []string
	for _, row := range rows[1:] {
		if len(row) <= max(teacherIndx, attendanceIndx) {
			continue
		}
		teacher := strings.TrimSpace(row[teacherIndx])
		attStr := strings.TrimSpace(row[attendanceIndx])
		if teacher == "" || attStr == "" {
			continue
		}
		attStr = strings.TrimSuffix(attStr, "%")
		if att, err := strconv.ParseFloat(attStr, 64); err == nil {
			if att < 40 {
				lowAttendanceTeachers = append(lowAttendanceTeachers, fmt.Sprintf("%s (%.1f%%)", teacher, att))
			}
		}
	}
	strbuild.WriteString("👨‍🏫 ОТЧЕТ ПО ПОСЕЩАЕМОСТИ ПРЕПОДАВАТЕЛЕЙ\n\n")
	if len(lowAttendanceTeachers) > 0 {
		strbuild.WriteString("Преподаватели с посещаемостью ниже 40%:\n")
		for i, t := range lowAttendanceTeachers {
			strbuild.WriteString(fmt.Sprintf("%d. %s\n", i+1, t))
		}
	} else {
		strbuild.WriteString("✅ У всех преподавателей посещаемость 40% и выше")
	}
	return strbuild.String(), nil
}

// 5. Проверка проверенных домашних
func processCheckedHomework(filepath string) (string, error) {
	file, err := excelize.OpenFile(filepath)
	if err != nil {
		return "", err
	}
	defer file.Close()

	rows, err := file.GetRows(file.GetSheetName(0))
	if err != nil || len(rows) < 2 {
		return "Нет данных в файле", nil
	}
	header := rows[1]
	teacherIdx, checkedIdx, totalIdx := -1, -1, -1
	for i, col := range header {
		switch strings.ToLower(col) {
		case "фио преподавателя":
			teacherIdx = i
		case "проверено":
			checkedIdx = i
		case "получено":
			totalIdx = i
		}
	}
	if teacherIdx == -1 || checkedIdx == -1 || totalIdx == -1 {
		return "Не найдены необходимые колонки", nil
	}
	var lowPercentTeachers []string
	for _, row := range rows[1:] {
		if len(row) <= max(teacherIdx, checkedIdx, totalIdx) {
			continue
		}
		teacher := strings.TrimSpace(row[teacherIdx])
		checkedStr := strings.TrimSpace(row[checkedIdx])
		totalStr := strings.TrimSpace(row[totalIdx])
		if teacher == "" || checkedStr == "" || totalStr == "" {
			continue
		}
		checked, err1 := strconv.ParseFloat(checkedStr, 64)
		total, err2 := strconv.ParseFloat(totalStr, 64)
		if err1 == nil && err2 == nil && total > 0 {
			percent := (checked / total) * 100
			if percent < 70 {
				lowPercentTeachers = append(lowPercentTeachers, fmt.Sprintf("%s (%.1f%% проверено)", teacher, percent))
			}
		}
	}
	strbuild.WriteString("📝 ОТЧЕТ ПО ПРОВЕРЕННЫМ ДОМАШНИМ ЗАДАНИЯМ\n\n")
	if len(lowPercentTeachers) > 0 {
		strbuild.WriteString("Преподаватели с проверкой ниже 70%:\n")
		for i, t := range lowPercentTeachers {
			strbuild.WriteString(fmt.Sprintf("%d. %s\n", i+1, t))
		}
	} else {
		strbuild.WriteString("✅ Все преподаватели проверяют более 70% заданий")
	}
	return strbuild.String(), nil
}

func processSubmittedHomework(filepath string) (string, error) {
	file, err := excelize.OpenFile(filepath)
	if err != nil {
		return "", err
	}
	defer file.Close()

	rows, err := file.GetRows(file.GetSheetName(0))
	if err != nil || len(rows) < 2 {
		return "Нет данных в файле", nil
	}

	header := rows[0]
	var studentIdx, percentIdx int = -1, -1

	for i, col := range header {
		colLower := strings.ToLower(col)
		if colLower == "фио" || colLower == "fio" {
			studentIdx = i
		}
		if colLower == "percentage homework" {
			percentIdx = i
		}
	}

	if studentIdx == -1 || percentIdx == -1 {
		return "Не найдены колонки ФИО или процента выполнения", nil
	}

	strbuild.WriteString("ФИО студента - % выполнения\n\n")
	for _, row := range rows[1:] {
		if len(row) <= max(studentIdx, percentIdx) {
			continue
		}
		fio := strings.TrimSpace(row[studentIdx])
		percent := strings.TrimSpace(row[percentIdx])
		percentInt, err := strconv.Atoi(percent)
		if err != nil {
			continue
		}
		if percentInt < 70 {
			strbuild.WriteString(fmt.Sprintf("%s - %s%%\n", fio, percent))
		}
	}

	return strbuild.String(), nil
}

func splitMessage(text string, maxLen int) []string {
	if len(text) <= maxLen {
		return []string{text}
	}
	var parts []string
	for len(text) > maxLen {
		idx := strings.LastIndex(text[:maxLen], "\n")
		if idx == -1 {
			idx = maxLen
		}
		parts = append(parts, strings.TrimSpace(text[:idx]))
		text = strings.TrimSpace(text[idx:])
	}
	if len(text) > 0 {
		parts = append(parts, strings.TrimSpace(text))
	}
	return parts
}
