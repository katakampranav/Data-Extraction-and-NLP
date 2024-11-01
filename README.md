# Text Analysis and Data Extraction

## Overview
This project performs text analysis on article content extracted from provided URLs. It calculates readability and sentiment metrics such as Positive/Negative Score, Fog Index, and Polarity Score, and outputs the results in an Excel file. The approach includes extracting text using BeautifulSoup, preprocessing with NLTK, and computing metrics for detailed insights.

## Features
- Article content extraction from URLs
- Text preprocessing and cleaning
- Calculation of various readability metrics
- Sentiment analysis
- Custom stopword handling
- Results export to Excel

## Project Setup

### Prerequisites
The following libraries are required:
- pandas
- nltk
- requests
- beautifulsoup4
- openpyxl

### Installation
```bash
pip install pandas nltk requests beautifulsoup4 openpyxl
```

### NLTK Resource Setup
```python
import nltk
nltk.download('punkt')
nltk.download('stopwords')
nltk.download('wordnet')
```

### Directory Structure
Ensure the following folders and files are present in the same directory as the script:
```
project/
├── Input.xlsx              # Input file containing URLs and IDs
├── model.py               # Main script
├── text_files/           # Folder for storing extracted text
├── StopWords/
│   ├── StopWords_Auditor.txt
│   ├── StopWords_Currencies.txt
│   ├── StopWords_DatesandNumbers.txt
│   ├── StopWords_Generic.txt
│   ├── StopWords_GenericLong.txt
│   ├── StopWords_Geographic.txt
│   └── StopWords_Names.txt
└── MasterDictionary/
    ├── negative-words.txt
    └── positive-words.txt
```

## Usage
1. Place your input file (`Input.xlsx`) in the project directory
2. Run the script:
```bash
python model.py
```

## Input and Output
### Input Format
- File: `Input.xlsx`
- Required columns: URL_ID, URL

### Output Format
Results are saved in `Output.xlsx` with the following metrics:
- URL_ID
- URL
- POSITIVE SCORE
- NEGATIVE SCORE
- POLARITY SCORE
- SUBJECTIVITY SCORE
- AVG SENTENCE LENGTH
- PERCENTAGE OF COMPLEX WORDS
- FOG INDEX
- AVG NUMBER OF WORDS PER SENTENCE
- COMPLEX WORD COUNT
- WORD COUNT
- SYLLABLE PER WORD
- PERSONAL PRONOUNS
- AVG WORD LENGTH

## Project Workflow
1. **Data Extraction**
   - Extracts article content and titles from URLs
   - Uses requests and BeautifulSoup libraries
   
2. **Text Preprocessing**
   - Converts text to lowercase
   - Removes non-alphanumeric characters
   - Tokenizes text
   - Removes stopwords
   - Performs lemmatization
   
3. **Metric Calculation**
   - Calculates sentiment scores
   - Computes readability metrics
   - Counts personal pronouns
   - Analyzes text complexity
   
4. **Error Handling**
   - Handles inaccessible URLs
   - Continues processing remaining URLs on error
  
![image](https://github.com/user-attachments/assets/36ad7bd0-e943-4e4a-96ed-771cec8df4b5)

## Author

This Text Analysis and Data Extraction project was developed by :
-	[@katakampranav](https://github.com/katakampranav)
-	Repository : https://github.com/katakampranav/Text Analysis and Data Extraction

## Feedback

For any feedback or queries, please reach out to me at katakampranavshankar@gmail.com.
