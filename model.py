## Importing the Libraries
import os
import re
import pandas as pd
import nltk
import requests
from bs4 import BeautifulSoup
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.stem import WordNetLemmatizer

## Setup for NLP
nltk.download('punkt')
nltk.download('stopwords')
nltk.download('wordnet')
lemmatizer = WordNetLemmatizer()
stop_words = set(stopwords.words('english'))

## Load Input Data
input_data = pd.read_excel('Input.xlsx')

## Define directories for saving text files and outputs
output_text_dir = 'text_files'
if not os.path.exists(output_text_dir):
    os.makedirs(output_text_dir)

## Load custom stopwords
StopWords_Auditor = open('StopWords\\StopWords_Auditor.txt', 'r', encoding='ISO-8859-1')
StopWords_Currencies = open('StopWords\\StopWords_Currencies.txt', 'r', encoding='ISO-8859-1')
StopWords_DatesandNumbers = open('StopWords\\StopWords_DatesandNumbers.txt', 'r', encoding='ISO-8859-1')
StopWords_Generic = open('StopWords\\StopWords_Generic.txt', 'r', encoding='ISO-8859-1')
StopWords_GenericLong = open('StopWords\\StopWords_GenericLong.txt', 'r', encoding='ISO-8859-1')
StopWords_Geographic = open('StopWords\\StopWords_Geographic.txt', 'r', encoding='ISO-8859-1')
StopWords_Names = open('StopWords\\StopWords_Names.txt', 'r', encoding='ISO-8859-1')

## Function to extract article title and content
def extract_article_content(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')
    article_title = soup.find('title').text
    article_content = soup.find('div', class_='td-post-content').get_text(strip=True, separator=' ')
    return article_title, article_content

## Function to preprocess text
def preprocess_text(x):
    text_lower = str(x).lower()
    clean_text = re.sub('[^a-zA-Z]+',' ',text_lower).strip()
    tokens = word_tokenize(clean_text)
    token_word = [token for token in tokens if token not in (StopWords_Auditor, StopWords_Currencies, StopWords_DatesandNumbers, StopWords_Generic, StopWords_GenericLong, StopWords_Geographic, StopWords_Names)]
    filtered_tokens = [lemmatizer.lemmatize(w) for w in token_word]
    return filtered_tokens

## Function to count syllables in a word
def count_syllables(word):
    vowels = "aeiouy"
    word = word.lower()
    count = 0
    previous_char_was_vowel = False
    for char in word:
        if char in vowels:
            if not previous_char_was_vowel:
                count += 1
            previous_char_was_vowel = True
        else:
            previous_char_was_vowel = False
    return count

## Read Positive and Negative words for Sentiment Analysis
with open('MasterDictionary/negative-words.txt', 'r', encoding='ISO-8859-1') as neg_file:
    negative_words = set(neg_file.read().split())

with open('MasterDictionary/positive-words.txt', 'r', encoding='ISO-8859-1') as pos_file:
    positive_words = set(pos_file.read().split())

## Function to calculate text metrics
def calculate_text_metrics(tokens, content, url_id, url):
    word_count = len(tokens)
    positive_score = sum(1 for word in tokens if word in positive_words)
    negative_score = sum(1 for word in tokens if word in negative_words)
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (word_count + 0.000001)

    total_sentences = len(sent_tokenize(content))
    avg_sentence_length = round(word_count / total_sentences, 0) if total_sentences > 0 else 0 
    complex_words = sum(1 for word in tokens if count_syllables(word) >= 2)
    percentage_complex_words = (complex_words / word_count) * 100 if word_count > 0 else 0
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)

    avg_words_per_sentence = round(word_count / total_sentences, 0) if total_sentences > 0 else 0
    personal_pronouns = sum(1 for word in tokens if word.lower() in {'i', 'me', 'my', 'we', 'us', 'our', 'you', 'your'})
    avg_word_length = sum(len(word) for word in tokens) / word_count if word_count > 0 else 0

    return {
        'URL_ID': url_id,
        'URL': url,
        'POSITIVE SCORE': positive_score,
        'NEGATIVE SCORE': negative_score,
        'POLARITY SCORE': polarity_score,
        'SUBJECTIVITY SCORE': subjectivity_score,
        'AVG SENTENCE LENGTH': avg_sentence_length,
        'PERCENTAGE OF COMPLEX WORDS': percentage_complex_words,
        'FOG INDEX': fog_index,
        'AVG NUMBER OF WORDS PER SENTENCE': avg_words_per_sentence,
        'COMPLEX WORD COUNT': complex_words,
        'WORD COUNT': word_count,
        'SYLLABLE PER WORD': round(sum(count_syllables(word) for word in tokens) / word_count, 0) if word_count > 0 else 0,
        'PERSONAL PRONOUNS': personal_pronouns,
        'AVG WORD LENGTH': round(avg_word_length, 0)
    }

## Initialize list to store the final results
output_data = []

## Iterate through each URL in the input data
for index, row in input_data.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    try:
        # Extract the title and content from the URL
        title, content = extract_article_content(url)

        # Save the article content to a text file
        text_file_path = os.path.join(output_text_dir, f"{url_id}.txt")
        with open(text_file_path, 'w', encoding='utf-8') as file:
            file.write(f"Title: {title}\\n\\n{content}")

        # Preprocess the content for analysis
        processed_tokens = preprocess_text(content)

        # Calculate metrics with URL_ID and URL
        metrics = calculate_text_metrics(processed_tokens, content, url_id, url)
        output_data.append(metrics)

    except Exception as e:
        print(f"Failed to process URL {url_id}: {e}")

## Convert the result to a DataFrame and save as an Excel file
output_df = pd.DataFrame(output_data)
output_df.to_excel('Output.xlsx', index=False)