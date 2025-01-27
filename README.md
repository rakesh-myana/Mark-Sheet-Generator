# Mark-Sheet Generator

This is an automation tool built using Python that generates an Excel sheet containing student results, including roll numbers, names, semester-wise SGPA, and overall CGPA. The tool achieves an 85% time reduction compared to the manual process, significantly improving efficiency and accuracy.

## Automating Result Retrieval and Storage with Python 🚀

I'm excited to share this project where I developed a Python program to automate the retrieval of student results from the **JNTUH Results** website and store them in an Excel sheet. This automation helps save time and reduces manual effort for educators and administrators.

### 🔍 What It Does:
- **Navigates to the Result Website:** The program opens the [JNTUH Results Website](https://jntuhresults.vercel.app/).
- **Inputs Roll Numbers:** It inputs each student's roll number one by one.
- **Extracts Results:** The program retrieves results, including the student's name, roll number, and scores.
- **Stores in Excel:** All extracted data is stored in an organized manner in an Excel file.

### 💻 Technologies Used:
- **Python:** The core programming language used for the script.
- **Libraries:**
  - **Selenium:** Used for web scraping and automating browser actions.
  - **xlsxwriter:** Utilized for creating and writing to Excel files.

### 🛠 Challenges & Solutions:
- **Handling Dynamic Web Content:** The website's dynamic nature required the use of Selenium's WebDriverWait and various locator strategies to reliably interact with elements.
  - **Solution:** Used explicit waits to ensure elements are present before interacting with them and dynamically located elements based on their attributes and structure.
  
- **Extracting Nested Elements:** Extracting deeply nested elements, such as specific grades and CGPA, posed a challenge.
  - **Solution:** Leveraged specific XPath and CSS selectors to accurately retrieve the desired information from the webpage.

### 📊 Results:
- **Roll Number:** Unique identifier for each student.
- **Name:** The student's name.
- **Semester-wise SGPA:** Detailed scores for each semester.
- **CGPA:** Cumulative Grade Point Average.


### 📈 Future Enhancements:
- **Parallel Processing:** Implementing multithreading or multiprocessing to handle multiple roll numbers simultaneously, speeding up the process.
- **Database Integration:** Storing results in a database for easier querying and long-term storage.
This project not only streamlined the process of result retrieval but also provided valuable insights into web automation and data handling using Python.
Feel free to ask any questions or provide feedback!

