# FOSSHACK
Natural Language to Pandas Query Converter

🚀 Project Overview

This project is designed to allow users to interact with Excel data using natural language queries. It uses Pandas, FuzzyWuzzy, and Tkinter to translate user input into valid Pandas operations, making it easier to filter, sort, and analyze data without needing advanced programming knowledge.

🎯 Features

Load an Excel file and extract column names automatically.

Convert natural language queries into Pandas DataFrame operations.

Support for sorting, filtering, and retrieving specific information from the dataset.

Fuzzy matching for column names to handle variations in input.

Save query results into a new Excel file for further analysis.

🏆 Hackathon Use Case

This project is developed for a GitHub Hackathon to provide an innovative way for users to interact with their data. It can be particularly useful in data analytics competitions, educational purposes, and business intelligence applications.

📜 How to Use

1️⃣ Install Dependencies

Make sure you have Python 3.x installed. Then install the required libraries:

pip install pandas fuzzywuzzy python-Levenshtein tk

2️⃣ Run the Script

python main.py

A file dialog will appear prompting you to select an Excel file.

After selecting the file, enter natural language queries in the pop-up input box.

The query result will be displayed and saved in an Excel file.

3️⃣ Example Queries

"sorted by age in ascending order" → Returns the dataset sorted by age.

"highest total marks" → Finds the student with the highest total marks.

"students from class B" → Filters students belonging to class B.

"marks more than 40" → Lists students scoring more than 40.

"list all students along with phone numbers" → Extracts student names and phone numbers.

💻 Code Structure

main.py → The core script that processes user queries and interacts with the Excel file.

requirements.txt → Lists required dependencies.

README.md → This file explaining the project.

🔥 Contributing

We welcome contributions! If you find any bugs or want to add new features:

Fork the repository.

Create a new branch: git checkout -b feature-name

Commit your changes: git commit -m "Add feature-name"

Push to the branch: git push origin feature-name

Open a Pull Request

📜 License

This project is open-source and available under the MIT License.

Developed with ❤️ for the FOSSHACK GitHub Hackathon 🚀
