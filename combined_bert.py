import docx
from transformers import BertTokenizer, BertForSequenceClassification
import torch
import openpyxl
# python program that reads the .xlsx file containing keywords & topics
from test_keyword_read import read_keyword_topics, export_to_excel

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


# Get dictionary of keywords and topics
workbook = openpyxl.load_workbook("keywords_topics.xlsx")
categories_dict = read_keyword_topics(workbook["Key Words_Topics"])
keywords = [item for sublist in categories_dict.values() for item in sublist]

# Create a list storing results, which will be used to create the excel output
result_list = []

# Process the document for sentiment
for paragraph in doc.paragraphs:
    text = paragraph.text
    sentiment = analyze_sentiment(text)

    found_keywords = [keyterm for keyterm in keywords if any(keyword in text for keyword in keyterm.split())]

    if found_keywords:
        category = next((category for category, keyterm in categories_dict.items() if all(keyword in keyterm for keyword in found_keywords)), None)
        result_list.append([text, float(sentiment), ', '.join(found_keywords), category])


# Output results to excel sheet
export_to_excel(result_list, "sentiment_results.xlsx")

workbook.close()

print("Sentiment analysis and keyword extraction completed. Results saved in 'sentiment_results.xlsx'.")
