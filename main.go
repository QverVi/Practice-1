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
	"github.com/joho/godotenv"
	"github.com/xuri/excelize/v2"
)

var bot *tgbotapi.BotAPI

func main() {
	// Загружаем переменные окружения
	err := godotenv.Load(".env")
	if err != nil {
		fmt.Println(err)
		log.Fatal(".env not found")
	}

	bot, err = tgbotapi.NewBotAPI(os.Getenv("token_telegram_bot"))
	if err != nil {
		log.Fatalf("Failed to init api: %v", err)
	}

	// Настройка получения обновлений
	updateConf := tgbotapi.NewUpdate(0)
	updateConf.Timeout = 30
	updates := bot.GetUpdatesChan(updateConf)

	// Обработка обновлений
	for update := range updates {
		if update.Message != nil {
			log.Printf("Получено сообщение от %s: %s", update.Message.From.UserName, update.Message.Text)

			if update.Message.IsCommand() {
				handleCommand(bot, update.Message)
			} else if update.Message.Document != nil {
				handleDocument(bot, update.Message)
			} else {
				msg := tgbotapi.NewMessage(update.Message.Chat.ID, "Пожалуйста, отправьте Excel файл или команду /start")
				if _, err := bot.Send(msg); err != nil {
					log.Printf("Ошибка отправки сообщения: %v", err)
				}
			}
		}
	}
}

func handleCommand(bot *tgbotapi.BotAPI, msg *tgbotapi.Message) {
	switch msg.Command() {
	case "start":
		text := `Привет! Я бот для обработки учебных отчетов.

Поддерживаемые отчеты:
1. Отчет по расписанию групп
2. Отчет по темам занятий  
3. Отчет по студентам
4. Отчет по посещаемости преподавателей
5. Отчет по проверенным ДЗ
6. Отчет по сданным ДЗ студентами

Отправьте Excel файл — я определю тип и подготовлю отчет.`
		response := tgbotapi.NewMessage(msg.Chat.ID, text)
		if _, err := bot.Send(response); err != nil {
			log.Printf("Ошибка отправки сообщения: %v", err)
		}

	case "help":
		response := tgbotapi.NewMessage(msg.Chat.ID, "Отправьте XLS файл, и я подготовлю нужный отчет.")
		if _, err := bot.Send(response); err != nil {
			log.Printf("Ошибка отправки сообщения: %v", err)
		}

	default:
		response := tgbotapi.NewMessage(msg.Chat.ID, "Неизвестная команда. Используйте /start или /help")
		if _, err := bot.Send(response); err != nil {
			log.Printf("Ошибка отправки сообщения: %v", err)
		}
	}
}

func handleDocument(bot *tgbotapi.BotAPI, msg *tgbotapi.Message) {
	filename := msg.Document.FileName
	if !(strings.HasSuffix(filename, ".xlsx") || strings.HasSuffix(filename, ".xls")) {
		bot.Send(tgbotapi.NewMessage(msg.Chat.ID, "Пожалуйста, отправьте файл в формате Excel (.xlsx или .xls)"))
		return
	}

	sentMsg, _ := bot.Send(tgbotapi.NewMessage(msg.Chat.ID, "⏳ Обрабатываю файл..."))

	file, err := bot.GetFile(tgbotapi.FileConfig{FileID: msg.Document.FileID})
	if err != nil {
		bot.Send(tgbotapi.NewMessage(msg.Chat.ID, "Ошибка при получении файла"))
		return
	}
	url := file.Link(bot.Token)

	localPath := fmt.Sprintf("temp_%d_%s", msg.MessageID, filename)
	defer os.Remove(localPath)
	if err := downloadFile(url, localPath); err != nil {
		bot.Send(tgbotapi.NewMessage(msg.Chat.ID, "Ошибка при скачивании файла"))
		return
	}

	category := determineCategory(localPath)
	if category == "" {
		bot.Send(tgbotapi.NewMessage(msg.Chat.ID, "Не удалось определить тип файла. Проверьте формат."))
		return
	}

	var res string
	var errProc error

	switch category {
	case "Расписание групп":
		res, errProc = processSchedule(localPath)
	case "Темы уроков":
		res, errProc = processLessonTopics(localPath)
	case "Отчет по студентам":
		res, errProc = processStudents(localPath)
	case "Посещаемость по преподавателям":
		res, errProc = processAttendance(localPath)
	case "Отчет по проверенным ДЗ":
		res, errProc = processCheckedHomework(localPath)
	case "Отчет по сданным ДЗ":
		res, errProc = processSubmittedHomework(localPath)
	default:
		res = "Обработка этого типа файла не реализована."
	}

	if errProc != nil {
		bot.Send(tgbotapi.NewMessage(msg.Chat.ID, fmt.Sprintf("Ошибка: %v", errProc)))
		return
	}

	parts := splitMessage(res, 4000)
	bot.Send(tgbotapi.NewDeleteMessage(msg.Chat.ID, sentMsg.MessageID))
	for _, p := range parts {
		bot.Send(tgbotapi.NewMessage(msg.Chat.ID, p))
	}
}

func downloadFile(url, path string) error {
	resp, err := http.Get(url)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	if resp.StatusCode != http.StatusOK {
		return fmt.Errorf("HTTP статус %d", resp.StatusCode)
	}
	out, err := os.Create(path)
	if err != nil {
		return err
	}
	defer out.Close()
	_, err = io.Copy(out, resp.Body)
	return err
}

func determineCategory(filepath string) string {
	f, err := excelize.OpenFile(filepath)
	if err != nil {
		return ""
	}
	defer f.Close()
	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return ""
	}
	rows, err := f.GetRows(sheets[0])
	if err != nil || len(rows) == 0 {
		return ""
	}
	header := rows[0]
	txt := strings.ToLower(strings.Join(header, " "))

	// 1. Расписание групп
	if strings.Contains(txt, "группа") && strings.Contains(txt, "время") && strings.Contains(txt, "пара") {
		return "Расписание групп"
	}

	// 2. Темы уроков
	if strings.Contains(txt, "урок") || strings.Contains(txt, "тема") || strings.Contains(txt, "тема урока") {
        return "Темы уроков"
	}
	// 3. Отчет по студентам
	if strings.Contains(txt, "fio") || (strings.Contains(txt, "homework") && strings.Contains(txt, "classroom")) {
		return "Отчет по студентам"
	}

	// 4. Посещаемость по преподавателям
	if strings.Contains(txt, "фио преподавателя") && strings.Contains(txt, "средняя посещаемость") {
		return "Посещаемость по преподавателям"
	}

	// 5. Отчет по проверенным ДЗ
	if strings.Contains(txt, "форма обучения") && strings.Contains(txt, "фио преподавателя")||
		(strings.Contains(txt, "месяц") || strings.Contains(txt, "неделя")) || strings.Contains(txt, "день") || strings.Contains(txt, "проверено") {
		return "Отчет по проверенным ДЗ"
	}

	// 6. Отчет по сданным ДЗ
	if strings.Contains(txt, "fio") && strings.Contains(txt, "percentage homework") || strings.Contains(txt, "домашнее") {
		return "Отчет по сданным ДЗ"
	}

	return ""
}

// 1. Отчет по расписанию групп
func processSchedule(filepath string) (string, error) {
	f, err := excelize.OpenFile(filepath)
	if err != nil {
		return "", err
	}
	defer f.Close()

	rows, err := f.GetRows(f.GetSheetName(0))
	if err != nil || len(rows) < 2 {
		return "Нет данных в файле", nil
	}

	header := rows[0]
	groupIdx, subjectIdx := -1, -1

	// Ищем колонки
	for i, col := range header {
		colLower := strings.ToLower(col)
		if strings.Contains(colLower, "группа") {
			groupIdx = i
		} else if strings.Contains(colLower, "время") || (len(colLower) > 0 && groupIdx != i) {
			// Если не нашли явно "предмет", берем первую не-группу колонку
			if subjectIdx == -1 {
				subjectIdx = i
			}
		}
	}

	if groupIdx == -1 || subjectIdx == -1 {
		return "Не удалось найти колонки 'Группа' и 'Предмет' в файле", nil
	}

	// Считаем пары по дисциплинам для каждой группы
	groupStats := make(map[string]map[string]int)

	for _, row := range rows[1:] {
		if len(row) <= max(groupIdx, subjectIdx) {
			continue
		}

		group := strings.TrimSpace(row[groupIdx])
		subject := strings.TrimSpace(row[subjectIdx])

		if group == "" || subject == "" {
			continue
		}

		if _, ok := groupStats[group]; !ok {
			groupStats[group] = make(map[string]int)
		}

		groupStats[group][subject]++
	}

	if len(groupStats) == 0 {
		return "Нет данных о расписании", nil
	}

	var result strings.Builder
	result.WriteString("📅 ОТЧЕТ ПО РАСПИСАНИЮ ГРУПП\n")
	result.WriteString("Количество пар по дисциплинам:\n\n")

	for group, subjects := range groupStats {
		result.WriteString(fmt.Sprintf("Группа: %s\n", group))
		for subject, count := range subjects {
			result.WriteString(fmt.Sprintf("  %s: %d пар\n", subject, count))
		}
		result.WriteString("\n")
	}

	return result.String(), nil
}

// 2. Отчет по темам занятий
func processLessonTopics(filepath string) (string, error) {
	f, err := excelize.OpenFile(filepath)
	if err != nil {
		return "", err
	}
	defer f.Close()

	rows, err := f.GetRows(f.GetSheetName(0))
	if err != nil || len(rows) == 0 {
		return "Нет данных в файле", nil
	}

	// Ищем колонку с темами
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

	var validTopics []string
	var invalidTopics []string
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

	var result strings.Builder
	result.WriteString("📚 ОТЧЕТ ПО ТЕМАМ ЗАНЯТИЙ\n\n")

	if len(validTopics) > 0 {
		result.WriteString("✅ Темы в правильном формате:\n")
		for _, topic := range validTopics {
			result.WriteString(fmt.Sprintf("• %s\n", topic))
		}
		result.WriteString("\n")
	}

	if len(invalidTopics) > 0 {
		result.WriteString("❌ Темы в НЕправильном формате:\n")
		for _, topic := range invalidTopics {
			result.WriteString(fmt.Sprintf("• %s\n", topic))
		}
	} else if len(validTopics) == 0 {
		result.WriteString("Темы уроков не найдены")
	}

	return result.String(), nil
}

// 3. Отчет по студентам
func processStudents(filepath string) (string, error) {
	f, err := excelize.OpenFile(filepath)
	if err != nil {
		return "", err
	}
	defer f.Close()

	rows, err := f.GetRows(f.GetSheetName(0))
	if err != nil || len(rows) < 2 {
		return "Нет данных в файле", nil
	}

	header := rows[0]
	fioIdx, homeworkIdx, classroomIdx := -1, -1, -1

	for i, col := range header {
		colLower := strings.ToLower(col)
		if strings.Contains(colLower, "fio") {
			fioIdx = i
		} else if strings.Contains(colLower, "homework") {
			homeworkIdx = i
		} else if strings.Contains(colLower, "classroom") {
			classroomIdx = i
		}
	}

	if fioIdx == -1 {
		return "Не найдена колонка с ФИО студентов", nil
	}

	var problemStudents []string

	for _, row := range rows[1:] {
		if len(row) <= max(fioIdx, homeworkIdx, classroomIdx) {
			continue
		}

		name := strings.TrimSpace(row[fioIdx])
		if name == "" {
			continue
		}

		// Проверяем домашнюю работу (средняя оценка = 1)
		if homeworkIdx != -1 && len(row) > homeworkIdx {
			homeworkStr := strings.TrimSpace(row[homeworkIdx])
			if homeworkStr == "1" {
				problemStudents = append(problemStudents, fmt.Sprintf("%s (домашняя работа: 1)", name))
				continue
			}
		}

		// Проверяем классную работу (< 3)
		if classroomIdx != -1 && len(row) > classroomIdx {
			classroomStr := strings.TrimSpace(row[classroomIdx])
			classroomGrade, err := strconv.ParseFloat(classroomStr, 64)
			if err == nil && classroomGrade < 3 {
				problemStudents = append(problemStudents, fmt.Sprintf("%s (классная работа: %.1f)", name, classroomGrade))
			}
		}
	}

	var result strings.Builder
	result.WriteString("👨‍🎓 ОТЧЕТ ПО СТУДЕНТАМ\n\n")

	if len(problemStudents) > 0 {
		result.WriteString("Студенты, требующие внимания:\n")
		for i, student := range problemStudents {
			result.WriteString(fmt.Sprintf("%d. %s\n", i+1, student))
		}
	} else {
		result.WriteString("✅ Все студенты успешно справляются с заданиями")
	}

	return result.String(), nil
}

// 4. Отчет по посещаемости преподавателей
func processAttendance(filepath string) (string, error) {
	f, err := excelize.OpenFile(filepath)
	if err != nil {
		return "", err
	}
	defer f.Close()

	rows, err := f.GetRows(f.GetSheetName(0))
	if err != nil || len(rows) < 2 {
		return "Нет данных в файле", nil
	}

	header := rows[0]
	teacherIdx, attendanceIdx := -1, -1

	for i, col := range header {
		colLower := strings.ToLower(col)
		if strings.Contains(colLower, "фио преподавателя") {
			teacherIdx = i
		} else if strings.Contains(colLower, "средняя посещаемость") {
			attendanceIdx = i
		}
	}

	if teacherIdx == -1 || attendanceIdx == -1 {
		return "Не найдены необходимые колонки в файле", nil
	}

	var lowAttendanceTeachers []string

	for _, row := range rows[1:] {
		if len(row) <= max(teacherIdx, attendanceIdx) {
			continue
		}

		teacher := strings.TrimSpace(row[teacherIdx])
		attendanceStr := strings.TrimSpace(row[attendanceIdx])

		if teacher == "" || attendanceStr == "" {
			continue
		}

		// Убираем знак процента если есть
		attendanceStr = strings.TrimSuffix(attendanceStr, "%")
		attendance, err := strconv.ParseFloat(attendanceStr, 64)
		if err != nil {
			continue
		}

		if attendance < 40 {
			lowAttendanceTeachers = append(lowAttendanceTeachers,
				fmt.Sprintf("%s (%.1f%%)", teacher, attendance))
		}
	}

	var result strings.Builder
	result.WriteString("👨‍🏫 ОТЧЕТ ПО ПОСЕЩАЕМОСТИ ПРЕПОДАВАТЕЛЕЙ\n\n")

	if len(lowAttendanceTeachers) > 0 {
		result.WriteString("Преподаватели с посещаемостью ниже 40%:\n")
		for i, teacher := range lowAttendanceTeachers {
			result.WriteString(fmt.Sprintf("%d. %s\n", i+1, teacher))
		}
	} else {
		result.WriteString("✅ У всех преподавателей посещаемость 40% и выше")
	}

	return result.String(), nil
}

// 5. Отчет по проверенным домашним заданиям
func processCheckedHomework(filepath string) (string, error) {
	f, err := excelize.OpenFile(filepath)
	if err != nil {
		return "", err
	}
	defer f.Close()

	rows, err := f.GetRows(f.GetSheetName(0))
	if err != nil || len(rows) < 2 {
		return "Нет данных в файле", nil
	}

	header := rows[0]
	teacherIdx, checkedIdx, totalIdx := -1, -1, -1

	for i, col := range header {
		colLower := strings.ToLower(col)
		if strings.Contains(colLower, "фио преподавателя") {
			teacherIdx = i
		} else if strings.Contains(colLower, "форма обучения") {
			checkedIdx = i
		} else if strings.Contains(colLower, "получено") {
			totalIdx = i
		}
	}

	if teacherIdx == -1 || checkedIdx == -1 || totalIdx == -1 {
		return "Не найдены необходимые колонки в файле", nil
	}

	var lowCheckTeachers []string

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

		if err1 != nil || err2 != nil || total == 0 {
			continue
		}

		percentage := (checked / total) * 100
		if percentage < 70 {
			lowCheckTeachers = append(lowCheckTeachers,
				fmt.Sprintf("%s (%.1f%% проверено)", teacher, percentage))
		}
	}

	var result strings.Builder
	result.WriteString("📝 ОТЧЕТ ПО ПРОВЕРЕННЫМ ДОМАШНИМ ЗАДАНИЯМ\n\n")

	if len(lowCheckTeachers) > 0 {
		result.WriteString("Преподаватели с процентом проверки ниже 70%:\n")
		for i, teacher := range lowCheckTeachers {
			result.WriteString(fmt.Sprintf("%d. %s\n", i+1, teacher))
		}
	} else {
		result.WriteString("✅ Все преподаватели проверяют более 70% заданий")
	}

	return result.String(), nil
}

// 6. Отчет по сданным домашним заданиям студентами
func processSubmittedHomework(filepath string) (string, error) {
	f, err := excelize.OpenFile(filepath)
	if err != nil {
		return "", err
	}
	defer f.Close()

	rows, err := f.GetRows(f.GetSheetName(0))
	if err != nil || len(rows) < 2 {
		return "Нет данных в файле", nil
	}

	header := rows[0]
	studentIdx, submittedIdx, totalIdx := -1, -1, -1

	for i, col := range header {
		colLower := strings.ToLower(col)
		if strings.Contains(colLower, "fio") {
			studentIdx = i
		} else if strings.Contains(colLower, "percentage homework") {
			submittedIdx = i
		} else if strings.Contains(colLower, "домашнее") {
			totalIdx = i
		}
	}

	if studentIdx == -1 {
		return "Не найдена колонка с ФИО студентов", nil
	}

	var lowSubmissionStudents []string

	for _, row := range rows[1:] {
		if len(row) <= max(studentIdx, submittedIdx, totalIdx) {
			continue
		}

		student := strings.TrimSpace(row[studentIdx])
		if student == "" {
			continue
		}

		// Если есть данные о сданных и всего заданиях
		if submittedIdx != -1 && totalIdx != -1 && len(row) > submittedIdx && len(row) > totalIdx {
			submittedStr := strings.TrimSpace(row[submittedIdx])
			totalStr := strings.TrimSpace(row[totalIdx])

			submitted, err1 := strconv.ParseFloat(submittedStr, 64)
			total, err2 := strconv.ParseFloat(totalStr, 64)

			if err1 == nil && err2 == nil && total > 0 {
				percentage := (submitted / total) * 100
				if percentage < 70 {
					lowSubmissionStudents = append(lowSubmissionStudents,
						fmt.Sprintf("%s (%.1f%% выполнено)", student, percentage))
				}
			}
		} else if submittedIdx != -1 && len(row) > submittedIdx { // Если есть колонка с процентом
			percentStr := strings.TrimSpace(row[submittedIdx])
			percentStr = strings.TrimSuffix(percentStr, "%")
			percentage, err := strconv.ParseFloat(percentStr, 64)
			if err == nil && percentage < 70 {
				lowSubmissionStudents = append(lowSubmissionStudents,
					fmt.Sprintf("%s (%.1f%% выполнено)", student, percentage))
			}
		}
	}

	var result strings.Builder
	result.WriteString("📚 ОТЧЕТ ПО СДАННЫМ ДОМАШНИМ ЗАДАНИЯМ СТУДЕНТАМИ\n\n")

	if len(lowSubmissionStudents) > 0 {
		result.WriteString("Студенты с процентом выполнения ниже 70%:\n")
		for i, student := range lowSubmissionStudents {
			result.WriteString(fmt.Sprintf("%d. %s\n", i+1, student))
		}
	} else {
		result.WriteString("✅ У всех студентов процент выполнения 70% и выше")
	}

	return result.String(), nil
}

func max(nums ...int) int {
	m := nums[0]
	for _, n := range nums {
		if n > m {
			m = n
		}
	}
	return m
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
