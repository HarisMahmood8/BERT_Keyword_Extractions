import openpyxl

# Open result spreadsheets
workbook_1 = openpyxl.load_workbook("sentiment_results_Q1.xlsx")
workbook_2 = openpyxl.load_workbook("sentiment_results_Q2.xlsx")
workbook_3 = openpyxl.load_workbook("sentiment_results_Q3.xlsx")

res_dict = {}

workbooks = [workbook_1, workbook_2, workbook_3]

# Iterate through each sheet
for workbook in workbooks:
    sub_dict = {}
    sheet = workbook.active

    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, values_only=True):
        # List of categories found in the row's paragraph
        categories = row[3].split(", ")

        # Get list of sentiment values assigned to each category
        for category in categories:
            key, value = category, float(row[1])

            if key not in sub_dict.keys():
                sub_dict[key] = []
            sub_dict[key].append(value)

    # Get average sentiment for each category and add them to the result dictionary
    for category, values in sub_dict.items():
        average = sum(values)/len(values)

        if category not in res_dict.keys():
            res_dict[category] = []
        res_dict[category].append(average)

    workbook.close()

# Print results
for category, scores in res_dict.items():
    print(f"{category}: Q1 = {scores[0]}, Q2 = {scores[1]}, Q3 = {scores[2]}")