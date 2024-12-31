package main

import (
	"fmt"
	"log"
	"net/smtp"
	"os"

	"github.com/joho/godotenv"
	"github.com/xuri/excelize/v2"
)

type Company struct {
	EmpName string
	Name    string
	Email   string
}

func readExcel(filename string) ([]Company, error) {
	file, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, fmt.Errorf("error opening file: %w", err)
	}

	defer func() {
		_ = file.Close()
	}()

	rows, err := file.GetRows("Sheet1")
	if err != nil {
		return nil, fmt.Errorf("error reading rows: %w", err)
	}

	var companies []Company
	for i, row := range rows {
		// Skip header and ensure row has at least two columns
		if i == 0 || len(row) < 3 {
			continue
		}
		companies = append(companies, Company{
			EmpName: row[0],
			Name:    row[1],
			Email:   row[2],
		})
	}
	return companies, nil
}

func sendMail(to, subject string, body string) bool {
	from := os.Getenv("EMAIL")
	password := os.Getenv("EMAIL_PASSWORD")

	if from == "" || password == "" {
		log.Println("Environment variables EMAIL or EMAIL_PASSWORD are not set.")
		return false
	}

	headers := "MIME-Version: 1.0\r\n" +
		"Content-Type: text/html; charset=UTF-8\r\n" +
		"From: Ayush Goyal <" + from + ">\r\n" +
		"To: " + to + "\r\n" +
		"Subject: " + subject + "\r\n"

	msg := []byte(headers + "\r\n" + body)

	err := smtp.SendMail("smtp.gmail.com:587",
		smtp.PlainAuth("", from, password, "smtp.gmail.com"),
		from, []string{to}, msg)

	if err != nil {
		log.Printf("Failed to send email to %s: %v", to, err)
		return false
	}
	return true
}

func main() {
	err := godotenv.Load()
	if err != nil {
		log.Fatalf("Error loading .env file")
	}

	fmt.Println("Starting sending mails...")

	companies, err := readExcel(os.Getenv("FILENAME"))
	if err != nil {
		log.Fatalf("Failed to read Excel file: %v", err)
	}

	for _, company := range companies {
		emailBody := composeEmail(company.Name, company.EmpName)
		emailSubject := "Application for Full Stack/Software Developer Engineer Position at " + company.Name
		success := sendMail(company.Email, emailSubject, emailBody)
		if success {
			log.Printf("Email sent to %s (%s)\n", company.Name, company.Email)
		} else {
			log.Fatalf("Failed to send email to %s (%s)\n", company.Name, company.Email)
		}
	}

	fmt.Println("Successfully sends all emails.")
}

func composeEmail(companyName string, employeeName string) string {
	return fmt.Sprintf(`
	<p>Dear <strong>%s</strong>,</p>
	<p>I hope this message finds you well. My name is Ayush Goyal, and I am a passionate <strong>Full Stack Developer</strong> with a strong focus on building scalable, user-centric web applications.</p>
	<p>With proficiency in <strong>React.js, Next.js, Node.js</strong>, and databases such as <strong>PostgreSQL, MySQL</strong> and <strong>MongoDB</strong>. You can explore my resume here: <a href="https://drive.google.com/file/d/1DcCCzmSXNA-SSKfEa3F8WFxF468EXzNZ/view?usp=sharing">Resume Link</a>.</p>
	<p>I am reaching out to express my interest in contributing to <strong>%s</strong>. Leveraging my technical skills and innovative approach, I am confident in my ability to drive impactful outcomes and align with your company's goals.</p>
	<p>I would be thrilled to discuss how my background and expertise can add value to your team. Please let me know a convenient time for a conversation.</p>
	<p>Looking forward to your response!!</p>
	<p>Best regards,<br><strong>Ayush Goyal</strong><br>
	Phone: +91 8178262999<br>
	Email: <a href="mailto:ayushgoyal8178@gmail.com">ayushgoyal8178@gmail.com</a><br>
	GitHub: <a href="https://github.com/ayuugoyal">github.com/ayuugoyal</a><br>
	LinkedIn: <a href="https://linkedin.com/in/ayuugoyal">linkedin.com/in/ayuugoyal</a><br>
	Portfolio: <a href="https://ayuugoyal.vercel.app/">ayuugoyal.vercel.app</a></p>
	`, employeeName, companyName)
}
