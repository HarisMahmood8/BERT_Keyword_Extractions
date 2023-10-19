import docx
from transformers import BertTokenizer, BertForSequenceClassification
import torch
from keywords import keywords  # Import the list of keywords from keywords.py

# Load the BERT model
model = BertForSequenceClassification.from_pretrained("bert-base-uncased")
tokenizer = BertTokenizer.from_pretrained("bert-base-uncased")

# Load the DOCX file
doc = docx.Document("equifax_cc.docx")

# Define a function for sentiment analysis
def analyze_sentiment(text):
    inputs = tokenizer(text, return_tensors="pt", truncation=True, padding=True)
    outputs = model(**inputs)
    positive_logit = outputs.logits[0][1]
    negative_logit = outputs.logits[0][0]
    sentiment_score = positive_logit - negative_logit
    return sentiment_score


# Load the list of keywords from keywords.py (assuming you have a keywords.py file)
from keywords import keywords

# Create a text document for storing the results
result_file = open("sentiment_results.txt", "w")

# Process the document
for paragraph in doc.paragraphs:
    text = paragraph.text
    sentiment = analyze_sentiment(text)
    
    # Check if any of the keywords are found in the paragraph
    found_keywords = [keyword for keyword in keywords if keyword in text]
    
    # Write the paragraph, keywords, and sentiment to the result file
    if found_keywords:
        result_file.write(f"Paragraph: {text}\n")
        result_file.write(f"Keywords: {', '.join(found_keywords)}\n")
        result_file.write(f"Sentiment Score: {sentiment}\n\n")

result_file.close()

print("Sentiment analysis and keyword extraction completed. Results saved in 'sentiment_results.txt'.")
