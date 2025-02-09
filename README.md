# 📌 MsExcel-Language-Translator-Tool  
A VBA-based tool for **translating entire Excel workbooks** while preserving formatting, using the **Google Translate API**.

## 🚀 Features  
✅ Translates **entire Excel sheets** with a single click  
✅ **Maintains formatting & styling** of cells  
✅ Uses **Google Translate API** for accurate translation  
✅ Supports **any language** translation  
✅ **Free & open-source** VBA macro  

---

## 📌 Installation Guide  

### **Step 1: Open the VBA Editor**
1. Open **Microsoft Excel**.  
2. Press **`ALT + F11`** to open the **VBA Editor**.  
3. In the VBA Editor, go to **`File > Import File...`**.  

### **Step 2: Import the VBA Modules**
1. Download the repository files.  
2. Import the following two `.bas` files into your VBA editor:  
   - **`ExcelLanguageTranslator.bas`** (Main script)  
   - **`JsonConverter.bas`** (Required for parsing JSON responses)  

### **Step 3: Enable Required References**
1. In the **VBA Editor**, go to **`Tools > References...`**.  
2. Scroll down and check **"Microsoft Scripting Runtime"**.  
3. Click **OK**.  

---

## 📌 Usage Guide  

### **Step 1: Run the Macro**
1. Open your Excel file.  
2. Press **`ALT + F8`** to open the **Macro Window**.  
3. Select **`TranslateEntireWorkbook`**.  
4. Click **"Run"**.  

### **Step 2: Translation Process**
- The script will **automatically translate** all text in the workbook from **Japanese to English**.  
- A **message box** will appear once the translation is complete.  

---

## 📌 Changing the Translation Language  
To modify the source and target languages:  

1. Open the **VBA Editor (`ALT + F11`)**.  
2. Locate the **`TranslateGoogle`** function.  
3. Change the language codes in this line:  
   ```vb
   translatedText = TranslateGoogle(inputText, "ja", "en")

Use the following Google Translate language codes to specify source and target languages:

| Language      | Code |
|--------------|------|
| Japanese to English | `ja`, `en` |
| French to English   | `fr`, `en` |
| Spanish to German   | `es`, `de` |
| Chinese to English  | `zh`, `en` |

Replace with any valid Google Translate language codes as needed.

---

## 📌 Troubleshooting

### ❌ Translation Not Working?
✅ Ensure you have an active internet connection. The Google Translate API requires an internet connection to function properly.

### ❌ Macro Not Running?
✅ Enable macros in Excel:
1. Go to **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Macro Settings**.
2. Select **Enable all macros**.

---

## 📌 📂 Repository & Contributions

💡 Want to improve the tool? Fork the repository and contribute! 🚀

🔗 **GitHub Repository**: [MsExcel-Language-Translator-Tool](#)
