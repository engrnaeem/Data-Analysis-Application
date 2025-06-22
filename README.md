
# 📊 Excel Data Viewer and Analyzer

This is a simple and interactive desktop GUI application built with Python `Tkinter`, `Pandas`, and `Matplotlib`. It allows users to load Excel files (`.xls` or `.xlsx`), view the contents, clean data, perform basic analysis, and visualize selected columns using various plot types.

## 🛠 Features

- 📁 Load Excel files and display columns dynamically.
- 🧾 View the **head** and **tail** of the dataset.
- 📉 Perform **basic statistical analysis** using `pandas.describe()`.
- 🧹 **Clean data** by removing rows with missing values.
- 📊 Visualize selected columns using:
  - Line Plot
  - Scatter Plot
  - Bar Chart
  - Histogram

## 🖼 GUI Preview

> (Add screenshots or a GIF of the GUI here)

## 🐍 Requirements

Make sure you have Python 3.7+ installed and then install the required packages:

```bash
pip install pandas matplotlib openpyxl
```

## ▶️ How to Run

Clone the repository and run the main Python script:

```bash
git clone https://github.com/yourusername/excel-data-analyzer.git
cd excel-data-analyzer
python app.py
```

## 📂 How to Use

1. **Launch the app**.
2. Click on **"Choose Excel File"** and select a `.xlsx` or `.xls` file.
3. Use the **View Head & Tail** button to inspect the top and bottom rows.
4. Click **Analyze Data** to view descriptive statistics.
5. Use **Clean Data** to remove rows with missing values.
6. Select columns via checkboxes, then click **Plot Data** and choose a plot type.

## 📄 File Structure

```
excel-data-analyzer/
├── app.py             # Main application code
├── README.md          # Documentation
└── instructions.pdf   # (Optional) User instructions or assignment details
```

> Note: Make sure to place `instructions.pdf` in the repo if you're distributing usage guidelines.

## 🧠 Future Improvements

- Add CSV support.
- Export cleaned or analyzed data.
- Add more chart types (e.g., boxplot, pie chart).
- Enable column-wise data filtering.

## 📚 License

This project is open-source and available under the [MIT License](LICENSE).

## 👤 Author

**Muhammad Naeem Mumtaz**  
📧 engrnaeem.9298@gmail.com  
🔗 [LinkedIn](https://www.linkedin.com/in/muhammad-naeem-mumtaz-awan-59bab9202/)
