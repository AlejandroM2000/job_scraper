import maskpass, os
import LinkedinBot


if __name__ == "__main__":
    email = input("Enter your email: ")
    pw = maskpass.askpass(prompt="Enter your password: ", mask="*")
    file = input("Enter the file name you are creating: ")
    job_keywords = input("Enter the job industry you are looking for: ")
    country = input("Enter the country: ")
    bot = LinkedinBot()
    bot.run(email, pw, job_keywords, country)
    if not os.path.exists("../"+file):
        bot.create_workbook("../"+file)
    bot.excel_export()