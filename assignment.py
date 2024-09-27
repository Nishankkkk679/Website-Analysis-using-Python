import requests
from bs4 import BeautifulSoup
import pandas as pd
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords
import re

# Download NLTK resources
#nltk.download('punkt')
#nltk.download('stopwords')

# Function to extract article text from a given URL
def extract_article_text(url):
    try:
        # Send a GET request to the URL
        response = requests.get(url)
        # Parse the HTML content of the page
        soup = BeautifulSoup(response.content, "html.parser")
        
        # Find the article title
        title = soup.find("title").text.strip()
        
        # Find the article text (assuming it's within a specific HTML tag)
        article_text = ""
        # Replace 'article_tag' with the appropriate HTML tag where the article text is located
        article_tags = soup.find_all("article")
        for tag in article_tags:
            article_text += tag.text.strip() + "\n\n"
        
        return title, article_text
    except Exception as e:
        print(f"Error occurred while extracting article from {url}: {e}")
        return None, None

# Load positive and negative words
with open("C:\\Users\\nisha\\Downloads\\negative-words.txt", "r") as file:
    positive_words = set(file.read().splitlines())
with open("C:\\Users\\nisha\\Downloads\\positive-words.txt", "r") as file:
    negative_words = set(file.read().splitlines())


def load_stop_words(personal_stop_word_files):
    stop_words = set(stopwords.words('english'))
    for file_path in personal_stop_word_files:
        with open(file_path, "r") as file:
            stop_words.update(file.read().splitlines())
    return stop_words

personal_stop_word_files = ["C:\\Users\\nisha\\Downloads\\StopWords_Auditor.txt", "C:\\Users\\nisha\\Downloads\\StopWords_Currencies.txt", "C:\\Users\\nisha\\Downloads\\StopWords_DatesandNumbers.txt","C:\\Users\\nisha\\Downloads\\StopWords_Generic.txt","C:\\Users\\nisha\\Downloads\\StopWords_GenericLong.txt","C:\\Users\\nisha\\Downloads\\StopWords_Geographic.txt","C:\\Users\\nisha\\Downloads\\StopWords_Names.txt"]

# Load stop words
stop_words = load_stop_words(personal_stop_word_files)

# Function to clean and tokenize text
def clean_text(text):
    # Remove punctuations and convert to lowercase
    text = re.sub(r'[^\w\s]', '', text.lower())
    # Tokenize the text
    tokens = word_tokenize(text)
    # Remove stop words
    tokens = [token for token in tokens if token not in stop_words]
    return tokens

# Function to calculate syllable count per word
def syllable_count(word):
    word = word.lower()
    vowels = "aeiouy"
    count = 0
    if word[0] in vowels:
        count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1
    if word.endswith("e"):
        count -= 1
    if count == 0:
        count += 1
    return count

# Function to perform sentimental analysis
def sentimental_analysis(tokens):
    positive_score = sum(1 for word in tokens if word in positive_words)
    negative_score = sum(1 for word in tokens if word in negative_words)
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(tokens) + 0.000001)
    return positive_score, negative_score, polarity_score, subjectivity_score

# Function to perform readability analysis
def readability_analysis(text):
    sentences = sent_tokenize(text)
    words = clean_text(text)
    total_words = len(words)
    total_sentences = len(sentences)
    avg_sentence_length = total_words / total_sentences
    complex_word_count = sum(1 for word in words if syllable_count(word) > 2)
    percentage_of_complex_words = (complex_word_count / total_words) * 100
    fog_index = 0.4 * (avg_sentence_length + percentage_of_complex_words)
    avg_words_per_sentence = total_words / total_sentences
    syllable_per_word = sum(syllable_count(word) for word in words) / total_words
    personal_pronouns_count = sum(1 for word in words if word.lower() in ["i", "we", "my", "ours", "us"])
    avg_word_length = sum(len(word) for word in words) / total_words
    return avg_sentence_length, percentage_of_complex_words, fog_index, avg_words_per_sentence, \
           complex_word_count, total_words, syllable_per_word, personal_pronouns_count, avg_word_length

# Read the input.xlsx file to get the URLs
input_df = pd.read_excel("C:\\Users\\nisha\\Downloads\\Input.xlsx")

# Create a list to store dictionaries
output_data = []

# Loop through each row in the input DataFrame
for index, row in input_df.iterrows():
    url_id = row["URL_ID"]
    url = row["URL"]
    
    # Extract article title and text
    title, article_text = extract_article_text(url)
    
    if title and article_text:
        # Perform sentimental analysis
        tokens = clean_text(article_text)
        positive_score, negative_score, polarity_score, subjectivity_score = sentimental_analysis(tokens)
        
        # Perform readability analysis
        avg_sentence_length, percentage_of_complex_words, fog_index, avg_words_per_sentence, \
        complex_word_count, total_words, syllable_per_word, personal_pronouns_count, avg_word_length = readability_analysis(article_text)
        
        # Append the data to the list
        output_data.append({"URL_ID": url_id, "Title": title, "Article_Text": article_text, 
                            "POSITIVE SCORE": positive_score, "NEGATIVE SCORE": negative_score, 
                            "POLARITY SCORE": polarity_score, "SUBJECTIVITY SCORE": subjectivity_score, 
                            "AVG SENTENCE LENGTH": avg_sentence_length, 
                            "PERCENTAGE OF COMPLEX WORDS": percentage_of_complex_words, 
                            "FOG INDEX": fog_index, "AVG NUMBER OF WORDS PER SENTENCE": avg_words_per_sentence, 
                            "COMPLEX WORD COUNT": complex_word_count, "WORD COUNT": total_words, 
                            "SYLLABLE PER WORD": syllable_per_word, 
                            "PERSONAL PRONOUNS": personal_pronouns_count, "AVG WORD LENGTH": avg_word_length})
        print(f"Article extracted from {url}")
    else:
        print(f"Failed to extract article from {url}")

# Convert the list of dictionaries to a DataFrame
output_df = pd.DataFrame(output_data)

# Save the output DataFrame to an Excel file
output_df.to_excel("C:\\Users\\nisha\\Desktop\\Outputs2.xlsx", index=False)
print("Extraction completed. Output saved to Output.xlsx")
