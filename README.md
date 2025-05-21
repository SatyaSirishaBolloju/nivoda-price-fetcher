# 💎 Nivoda Excel Price Fetcher

This Python automation tool reads diamond inventory from an Excel file, sends queries to the [Nivoda](https://www.nivoda.net/) API for market prices, and writes the price and price per carat back into the Excel file.

---

## 🚀 Features

- 🔐 Authenticates via token with Nivoda GraphQL API
- 📊 Fetches pricing for each diamond row (shape, size, color, clarity)
- 📁 Updates the same Excel file with market pricing
- ⏱️ Significantly faster than manual search

---

## 📂 Project Structure

```
nivoda-price-fetcher/
├── nivoda_price_fetcher.py                  # Main script
├── nivoda_input_sample.txt          # Sample input config (safe to share)
├── .gitignore                       # Files to ignore
└── README.md                        # You're reading it

```
---

## ⚙️ How to Run

1. Install dependencies:
   ```bash
   pip install requests openpyxl
   
2. Create a config file named nivoda_excel.txt:
    ```bash
    file_path = NIVODA_EXCEL_FORMAT.xlsx
    diamond = 10
    token = YOUR_ACTUAL_API_TOKEN_HERE
    ```
    
3. Run the script
    ```bash
    python nivoda_price_fetcher.py
    ```
    
4. Excel will be updated will price.

---

## 🔐 Security Note

- Do NOT commit nivoda_input.txt to GitHub.
- Use .gitignore to exclude .xlsx and config files.
- Use nivoda_excel_sample.txt for safe sharing.
  
---

## 👤 About the Author

**Satya Sirisha Bolloju**   
🔗 [LinkedIn](www.linkedin.com/in/satya-sirisha-bolloju-031b33239)

---

## 📄 License

MIT License — free to use, modify, and contribute.

---

## ⭐ Feedback & Contributions

If you find this useful:
- Star the repo
- Fork the code
- Or open an Issue / Pull Request 

---
