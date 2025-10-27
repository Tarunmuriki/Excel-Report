# 📊 Excel Report Generator

A Python application that automates the generation of Excel reports from CSV data — including **pivot tables**, **charts**, and **summary statistics**.  
Built with **pandas**, **openpyxl**, and **matplotlib**, this tool simplifies data reporting and visualization through a clean GUI interface.

---

## 🚀 Features

- 📂 Import CSV files easily  
- 📈 Generate dynamic pivot tables  
- 📊 Create charts using **matplotlib**  
- 🎨 Export styled Excel reports using **openpyxl**  
- 🧮 View summary statistics (mean, median, totals, etc.)  
- 🖥️ Simple, user-friendly **GUI (Tkinter)**  

---

## 🧰 Requirements

- Python **3.8+**
- Libraries:
  - `pandas`
  - `openpyxl`
  - `matplotlib`
  - `tkinter` (usually comes preinstalled)

Install all dependencies using:
```bash
pip install -r requirements.txt
🏗️ Installation
Clone the repository:

bash
Copy code
git clone https://github.com/your-username/Excel-Report-Generator.git
cd Excel-Report-Generator
(Optional) Create and activate a virtual environment:

bash
Copy code
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
Install dependencies:

bash
Copy code
pip install -r requirements.txt
▶️ Usage
Run the main program:

bash
Copy code
python src/main.py
In the GUI:

Browse and select your input .csv file

Choose report options (summary, pivot, charts)

Generate and save your Excel report

Output reports will be saved as .xlsx files in the selected directory.

📂 Project Structure
bash
Copy code
.
├── src/
│   ├── main.py             # GUI entry point
│   ├── report_generator.py # Core Excel generation logic
│   └── utils.py            # Helper functions
├── data/
│   └── sample_data.csv     # Example CSV input
├── requirements.txt        # Project dependencies
└── README.md               # Documentation
📸 Example Output
An Excel file with formatted tables, pivot summaries, and charts automatically created from your CSV data.

🤝 Contributing
Contributions are welcome!
If you'd like to improve this project:

Fork the repository

Create a new branch

Commit your changes

Open a pull request 🚀

🪪 License
This project is released under the MIT License — feel free to use and modify it.
