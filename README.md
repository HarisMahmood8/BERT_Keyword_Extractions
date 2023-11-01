Clone the repository to your local machine:

    git clone https://github.com/HarisMahmood8/BERT_Keyword_Extractions.git

Install the required Python packages using pip:

    pip install docx openpyxl transformers torch pandas

The sentiment engine takes in a .docx file with the following name format:

    (company name)_Q(quarter number).docx

So for a company call from Equifax in the 3rd quarter, the docx would be named:

    Equifax_Q1.docx

Place the .docx file in the same directory as the python program. 

Run the roberta_model.py program:

    py roberta_model.py

or

    python roberta_model.py

The program will prompt you with the company name and quarter. Note that this is case-sensitive.

    Enter name of company: Equifax
    Enter number of quarters: 3

The program will output an .xlsx file.

To compile results by category and keyword over multiple quarters, run compile_results.py

    py compile_results.py

or

    python compile_results.py

The program will prompt you with the company name and the number of quarters to compile.

This is also case-sensitive, and compiles quarters starting from quarter 1.

    Enter name of company: Equifax
    Enter number of quarters: 3

This will compile the sentiment results from quarters 1, 2 and 3.

The output is saved to an .xlsx file.