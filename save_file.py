import pandas as pd


def save_to_excel(tasks):
    # Exporting tasks to an Excel file
    rows = []
    for day, task_list in tasks.items():
        for task in task_list:
            rows.append({'Day': day, 'Task': task})

    df = pd.DataFrame(rows)
    df.to_excel('my_to_do.xlsx', index=False)
