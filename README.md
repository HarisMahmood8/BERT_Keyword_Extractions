Clone the repository to your local machine:

    git clone https://github.com/HarisMahmood8/BERT_Keyword_Extractions.git

Install the required Python packages using pip:

    pip install docx openpyxl transformers torch pandas

The sentiment engine takes in a .docx file with the following name format:

    (company name)_Q(quarter number)_(year).docx

So for a company call from Equifax in the 3rd quarter of 2021, the docx would be named:

    Equifax_Q3_2021.docx

Place the .docx file in the same directory as the python program. 

Run the roberta_model.py program:

    py roberta_model.py

or

    python roberta_model.py

The program will prompt you with the name of the conference document file without the extension. Note that this is case-sensitive.

    Equifax_Q3_2021

The program will output an .xlsx file.
