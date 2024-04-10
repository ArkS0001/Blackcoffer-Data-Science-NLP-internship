#Team Perfect Cube:Aakarshit Srivastava,Bhaskar Banerjee,Ayush Verma 
import pandas as pd
from bs4 import BeautifulSoup
import requests
from textblob import TextBlob
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize

# Download NLTK resources
nltk.download('punkt')
nltk.download('stopwords')

# Read input Excel file
input_file = "Input.xlsx"
output_file = "Output Data Main.xlsx"

input_df = pd.read_excel(input_file)

# Initialize lists to store output data
output_data = {
    "URL_ID": [],
    "URL": [],
    "ARTICLE_TITLE": [],
    "ARTICLE_TEXT": [],
    "POSITIVE_SCORE": [],
    "NEGATIVE_SCORE": [],
    "POLARITY_SCORE": [],
    "SUBJECTIVITY_SCORE": [],
    "AVG_SENTENCE_LENGTH": [],
    "PERCENTAGE_OF_COMPLEX_WORDS": [],
    "FOG_INDEX": [],
    "AVG_NUMBER_OF_WORDS_PER_SENTENCE": [],
    "COMPLEX_WORD_COUNT": [],
    "WORD_COUNT": [],
    "SYLLABLE_PER_WORD": [],
    "PERSONAL_PRONOUNS": [],
    "AVG_WORD_LENGTH": []
}

# Function to extract text from URL
def extract_text_from_url(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # Find article title
    article_title = soup.find('title').text.strip()
    
    # Find article text
    article_text = ""
    for paragraph in soup.find_all('p'):
        article_text += paragraph.text.strip() + "\n"
    
    return article_title, article_text

# Function to calculate syllables in a word
def syllable_count(word):
    word = word.lower()
    count = 0
    vowels = 'aeiouy'
    if word[0] in vowels:
        count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1
    if word.endswith('e'):
        count -= 1
    if count == 0:
        count += 1
    return count

# Function to compute text analysis
def analyze_text(text):
    blob = TextBlob(text)
    sentences = sent_tokenize(text)
    words = word_tokenize(text)
    stop_words = set(stopwords.words('english'))

    positive_score = sum(1 for word in blob.words if TextBlob(word).sentiment.polarity > 0)
    negative_score = sum(1 for word in blob.words if TextBlob(word).sentiment.polarity < 0)
    polarity_score = blob.sentiment.polarity
    subjectivity_score = blob.sentiment.subjectivity
    avg_sentence_length = sum(len(sent.split()) for sent in sentences) / len(sentences)
    complex_words = [word for word in words if syllable_count(word) > 2 and word.lower() not in stop_words]
    percentage_complex_words = (len(complex_words) / len(words)) * 100 if len(words) > 0 else 0
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
    avg_words_per_sentence = len(words) / len(sentences)
    complex_word_count = len(complex_words)
    word_count = len(words)
    syllables = sum(syllable_count(word) for word in words)
    syllable_per_word = syllables / word_count if word_count > 0 else 0
    personal_pronouns = sum(1 for word in words if word.lower() in ['i', 'me', 'my', 'mine', 'myself', 'we', 'us', 'our', 'ours', 'ourselves', 'you', 'your', 'yours', 'yourself', 'yourselves'])
    avg_word_length = sum(len(word) for word in words) / len(words) if len(words) > 0 else 0
    
    return (positive_score, negative_score, polarity_score, subjectivity_score, 
            avg_sentence_length, percentage_complex_words, fog_index, avg_words_per_sentence, 
            complex_word_count, word_count, syllable_per_word, personal_pronouns, avg_word_length)

# Iterate over each row in input Excel file
for index, row in input_df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    
    # Extract text from URL
    article_title, article_text = extract_text_from_url(url)
    
    # Perform text analysis
    (positive_score, negative_score, polarity_score, subjectivity_score, avg_sentence_length, 
     percentage_complex_words, fog_index, avg_words_per_sentence, complex_word_count, word_count, 
     syllable_per_word, personal_pronouns, avg_word_length) = analyze_text(article_text)
    
    # Store data in output dictionary
    output_data["URL_ID"].append(url_id)
    output_data["URL"].append(url)
    output_data["ARTICLE_TITLE"].append(article_title)
    output_data["ARTICLE_TEXT"].append(article_text)
    output_data["POSITIVE_SCORE"].append(positive_score)
    output_data["NEGATIVE_SCORE"].append(negative_score)
    output_data["POLARITY_SCORE"].append(polarity_score)
    output_data["SUBJECTIVITY_SCORE"].append(subjectivity_score)
    output_data["AVG_SENTENCE_LENGTH"].append(avg_sentence_length)
    output_data["PERCENTAGE_OF_COMPLEX_WORDS"].append(percentage_complex_words)
    output_data["FOG_INDEX"].append(fog_index)
    output_data["AVG_NUMBER_OF_WORDS_PER_SENTENCE"].append(avg_words_per_sentence)
    output_data["COMPLEX_WORD_COUNT"].append(complex_word_count)
    output_data["WORD_COUNT"].append(word_count)
    output_data["SYLLABLE_PER_WORD"].append(syllable_per_word)
    output_data["PERSONAL_PRONOUNS"].append(personal_pronouns)
    output_data["AVG_WORD_LENGTH"].append(avg_word_length)

# Create DataFrame from output data
output_df = pd.DataFrame(output_data)

# Write output DataFrame to Excel file
output_df.to_excel(output_file, index=False)