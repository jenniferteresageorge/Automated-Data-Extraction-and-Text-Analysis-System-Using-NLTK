import re
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize
import pandas as pd
import os

# Ensure NLTK resources are available
nltk.download('punkt')
nltk.download('stopwords')

# Get the directory where the script is located
base_dir = os.path.dirname(os.path.abspath(__file__))

# Function to load words from a file
def load_words(filepath):
    with open(filepath, 'r') as file:
        words = file.read().splitlines()
    return set(words)

# Load positive and negative words
positive_words = load_words(os.path.join(base_dir, 'MasterDictionary', 'positive-words.txt'))
negative_words = load_words(os.path.join(base_dir, 'MasterDictionary', 'negative-words.txt'))

# Load additional stop words
stop_words_files = [
    'StopWords_Auditor.txt',
    'StopWords_Currencies.txt',
    'StopWords_DatesandNumbers.txt',
    'StopWords_Generic.txt',
    'StopWords_GenericLong.txt',
    'StopWords_Geographic.txt',
    'StopWords_Names.txt'
]

additional_stop_words = set()
for filename in stop_words_files:
    additional_stop_words.update(load_words(os.path.join(base_dir, 'StopWords', filename)))

# Combine with NLTK stop words
stop_words = set(stopwords.words('english')).union(additional_stop_words)

# Function to clean and tokenize text
def clean_and_tokenize(text):
    text = re.sub(r'\W', ' ', text)  # Remove all non-word characters
    text = re.sub(r'\s+', ' ', text)  # Replace all whitespace with a single space
    text = text.lower()  # Convert text to lowercase
    words = word_tokenize(text)
    words = [word for word in words if word not in stop_words]  # Remove stop words
    return words

# Function to compute positive score
def positive_score(text):
    words = clean_and_tokenize(text)
    score = sum(1 for word in words if word in positive_words)
    return score

# Function to compute negative score
def negative_score(text):
    words = clean_and_tokenize(text)
    score = sum(1 for word in words if word in negative_words)
    return score

# Function to compute polarity score
def polarity_score(positive, negative):
    return (positive - negative) / ((positive + negative) + 0.000001)

# Function to compute subjectivity score
def subjectivity_score(positive, negative, total_words):
    return (positive + negative) / (total_words + 0.000001)

# Function to compute average sentence length
def average_sentence_length(text):
    sentences = sent_tokenize(text)
    words = clean_and_tokenize(text)
    avg_length = len(words) / len(sentences)
    return avg_length

# Function to compute percentage of complex words
def percentage_complex_words(text):
    words = clean_and_tokenize(text)
    complex_words = sum(1 for word in words if count_syllables(word) > 2)
    return (complex_words / len(words)) * 100

# Function to compute fog index
def fog_index(avg_sentence_length, percentage_complex_words):
    return 0.4 * (avg_sentence_length + percentage_complex_words)

# Function to compute average number of words per sentence
def average_number_of_words_per_sentence(text):
    sentences = sent_tokenize(text)
    words = clean_and_tokenize(text)
    avg_words = len(words) / len(sentences)
    return avg_words

# Function to count complex words
def complex_word_count(text):
    words = clean_and_tokenize(text)
    complex_words = sum(1 for word in words if count_syllables(word) > 2)
    return complex_words

# Function to compute word count
def word_count(text):
    words = clean_and_tokenize(text)
    return len(words)

# Function to count syllables in a word
def count_syllables(word):
    word = word.lower()
    vowels = "aeiouy"
    syllable_count = 0
    if word[0] in vowels:
        syllable_count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            syllable_count += 1
    if word.endswith("e"):
        syllable_count -= 1
    if syllable_count == 0:
        syllable_count += 1
    return syllable_count

# Function to count personal pronouns
def personal_pronouns(text):
    words = clean_and_tokenize(text)
    pronouns = ['i', 'we', 'my', 'ours', 'us']
    pronoun_count = sum(1 for word in words if word in pronouns)
    return pronoun_count

# Function to compute average word length
def average_word_length(text):
    words = clean_and_tokenize(text)
    avg_length = sum(len(word) for word in words) / len(words)
    return avg_length

# Directory containing the extracted articles
articles_dir = os.path.join(base_dir, 'extracted_articles_1')

# Load the input Excel file
input_file_path = os.path.join(base_dir, 'Input.xlsx')
input_df = pd.read_excel(input_file_path)

# Initialize a DataFrame to store the results
results_list = []

# Compute metrics for each article
for index, row in input_df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    try:
        with open(os.path.join(articles_dir, f'{url_id}.txt'), 'r', encoding='utf-8') as file:
            text = file.read()
            positive = positive_score(text)
            negative = negative_score(text)
            polarity = polarity_score(positive, negative)
            subjectivity = subjectivity_score(positive, negative, word_count(text))
            avg_sentence_len = average_sentence_length(text)
            perc_complex_words = percentage_complex_words(text)
            fog_idx = fog_index(avg_sentence_len, perc_complex_words)
            avg_words_per_sentence = average_number_of_words_per_sentence(text)
            complex_word_cnt = complex_word_count(text)
            word_cnt = word_count(text)
            syllables_per_word = sum(count_syllables(word) for word in clean_and_tokenize(text)) / word_cnt
            personal_pronoun_cnt = personal_pronouns(text)
            avg_word_len = average_word_length(text)
            
            results_list.append({
                'URL_ID': url_id, 'URL': url, 'POSITIVE SCORE': positive,
                'NEGATIVE SCORE': negative, 'POLARITY SCORE': polarity,
                'SUBJECTIVITY SCORE': subjectivity, 'AVG SENTENCE LENGTH': avg_sentence_len,
                'PERCENTAGE OF COMPLEX WORDS': perc_complex_words, 'FOG INDEX': fog_idx,
                'AVG NUMBER OF WORDS PER SENTENCE': avg_words_per_sentence,
                'COMPLEX WORD COUNT': complex_word_cnt, 'WORD COUNT': word_cnt,
                'SYLLABLE PER WORD': syllables_per_word, 'PERSONAL PRONOUNS': personal_pronoun_cnt,
                'AVG WORD LENGTH': avg_word_len
            })
            print(f'Successfully processed article {url_id}')
    except Exception as e:
        print(f'Failed to process article {url_id}: {e}')

# Convert the results list to a DataFrame
results_df = pd.DataFrame(results_list)

# Save the results to an Excel file
output_file_path = os.path.join(base_dir, 'Output_Data_Structure.xlsx')
results_df.to_excel(output_file_path, index=False)
