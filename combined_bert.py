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

# Create a list storing results, which will be used to create the excel output
result_list = []

# Process the document for sentiment
for category in categories_dict:
    category_items = 0
    total_sentiment = 0.0

    keywords = categories_dict[category]
    for paragraph in doc.paragraphs:
        text = paragraph.text
        sentiment = analyze_sentiment(text)

        found_keywords = [keyword for keyword in keywords if keyword in text]

        if found_keywords:
            result_list.append([category, float(sentiment), text, ', '.join(found_keywords)])
            category_items += 1
            total_sentiment += float(sentiment)

    if category_items > 0:
        avg_sentiment = total_sentiment / category_items
        result_list.insert(len(result_list) - category_items, [category, avg_sentiment, "", ', '.join(keywords)])


# Output results to excel sheet
export_to_excel(result_list, "sentiment_results.xlsx")

workbook.close()

print("Sentiment analysis and keyword extraction completed. Results saved in 'sentiment_results.xlsx'.")
