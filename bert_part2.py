import docx
from transformers import BertTokenizer, BertForSequenceClassification
import torch
from keywords import keywords 

# Load the BERT model
model = BertForSequenceClassification.from_pretrained("fine_tuned_bert_model")
tokenizer = BertTokenizer.from_pretrained("bert-base-uncased")

# Load the DOCX file
doc = docx.Document("equifax_cc.docx")

def analyze_sentiment(text):
    inputs = tokenizer(text, return_tensors="pt", truncation=True, padding=True)
    outputs = model(**inputs)
    sentiment_score = outputs.logits[0][1] - outputs.logits[0][0]
    return sentiment_score

result_file = open("sentiment_results.txt", "w")

# Process the document
for paragraph in doc.paragraphs:
    text = paragraph.text
    sentiment = analyze_sentiment(text)
    
    found_keywords = [keyword for keyword in keywords if keyword in text]
    
    if found_keywords:
        result_file.write(f"Paragraph: {text}\n")
        result_file.write(f"Keywords: {', '.join(found_keywords)}\n")
        result_file.write(f"Sentiment Score: {sentiment}\n\n")

result_file.close()

print("Sentiment analysis and keyword extraction completed. Results saved in 'sentiment_results.txt'.")
