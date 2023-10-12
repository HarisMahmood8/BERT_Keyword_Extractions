import pandas as pd
import torch
from sklearn.model_selection import train_test_split
from torch.utils.data import DataLoader, TensorDataset
from transformers import BertTokenizer, BertForSequenceClassification, AdamW
from tqdm import tqdm
import warnings

warnings.filterwarnings("ignore", category=FutureWarning)

# Load your CSV dataset
data = pd.read_csv("test1.csv")

# Split the data into training and testing sets
train_data, test_data = train_test_split(data, test_size=0.2, random_state=42)

tokenizer = BertTokenizer.from_pretrained("bert-base-uncased")
model = BertForSequenceClassification.from_pretrained("bert-base-uncased")

train_texts = train_data["text"].tolist()
train_labels = train_data["label"].tolist()

def preprocess_text(texts, labels):
    inputs = tokenizer(texts, return_tensors="pt", truncation=True, padding=True)
    labels = torch.tensor(labels)
    return TensorDataset(inputs["input_ids"], inputs["attention_mask"], labels)

train_inputs = preprocess_text(train_texts, train_labels)

def fine_tune_bert(model, train_data, epochs=1, batch_size=8, learning_rate=2e-5):
    dataloader = DataLoader(train_data, batch_size=batch_size, shuffle=True)
    optimizer = AdamW(model.parameters(), lr=learning_rate)
    for epoch in range(epochs):
        model.train()
        total_loss = 0
        for batch in tqdm(dataloader, desc=f"Epoch {epoch+1}"):
            optimizer.zero_grad()
            input_ids, attention_mask, labels = batch
            outputs = model(input_ids, attention_mask=attention_mask, labels=labels)
            loss = outputs.loss
            loss.backward()
            optimizer.step()
            total_loss += loss.item()
        print(f"Epoch {epoch+1} Loss: {total_loss / len(dataloader)}")

fine_tune_bert(model, train_inputs)

model.save_pretrained("fine_tuned_bert_model")

model = BertForSequenceClassification.from_pretrained("fine_tuned_bert_model")

def analyze_sentiment(text):
    inputs = tokenizer(text, return_tensors="pt", truncation=True, padding=True)
    outputs = model(**inputs)
    sentiment = "Positive" if outputs.logits[0][1] > outputs.logits[0][0] else "Negative"
    return sentiment
