import pandas as pd
import math

def calculate_final_grade(average):
    if average < 50:
        return "Reprovado por Nota", 0
    if 50 <= average < 70:
        return "Exame Final", 100 - average
    return "Aprovado", 0

def main():
    FILE_PATH = 'Engenharia de Software – Desafio GUILHERME DIEL.xlsx'

    # Load the Excel file into a DataFrame
    df = pd.read_excel(FILE_PATH)

    # Extract the number of classes to determine maximum absence
    classes = int(df.iloc[0, 0].split()[-1])
    maximum_absence = int(classes * 0.25)

    # Set column names
    df.columns = df.iloc[1].to_list()

    # Iterate through each row to calculate student situation
    for index, row in df.iterrows():
        # Skip rows where 'P1' is not a valid integer
        if not isinstance(row["P1"], int):
            continue

        naf = 0

        # Check if the student is 'Reprovado por Falta'
        if row["Faltas"] > maximum_absence:
            df.at[index, 'Situação'] = "Reprovado por Falta"
        else:
            # Calculate the average of three exams
            average = (row["P1"] + row["P2"] + row["P3"]) / 3

            # Determine the situation and NAF if 'Exame Final'
            situation, naf = calculate_final_grade(average)
            df.at[index, 'Situação'] = situation

        # Round NAF and update the DataFrame
        df.at[index, 'Nota para Aprovação Final'] = math.ceil(naf)

        print(f"\t\tMédia: {average:.2f}\n{row}\n")


if __name__ == "__main__":
    main()
